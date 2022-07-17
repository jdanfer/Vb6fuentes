VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_medpendv 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Medicación pendiente de entrega que está vencido el plazo"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11145
   Icon            =   "frm_medpendv.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   11145
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ver TODO lo pendiente"
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
      Left            =   6960
      TabIndex        =   6
      ToolTipText     =   "Muestra todo lo pendiente aún sin vencimiento"
      Top             =   4920
      Width           =   3735
   End
   Begin VB.Data data_parse 
      Caption         =   "data_parse"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_motivos 
      Caption         =   "data_motivos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos para baja"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   6495
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   5
         Top             =   840
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Baja por vencimiento"
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
         Left            =   120
         Picture         =   "frm_medpendv.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   4800
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10320
      Picture         =   "frm_medpendv.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir pendientes de entrega vencidos "
      Top             =   4080
      Width           =   615
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6800
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Matrícula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   5362
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha inicial"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha final"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Medicación"
         Object.Width           =   5186
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tipo Prescrip"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Fecha HC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Médico"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frm_medpendv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Dim Xnom As String
    Dim Xcount As Long
    
    Data1.DatabaseName = App.path & "\selec.mdb"
    Data1.RecordSource = "selec"
    Data1.Refresh
    If Data1.Recordset.RecordCount > 0 Then
       Data1.Recordset.MoveFirst
       Do While Not Data1.Recordset.EOF
          Data1.Recordset.Delete
          Data1.Recordset.MoveNext
       Loop
    End If

If Check1.Value = 1 Then
    
    frm_medpendv.MousePointer = 11
    Data2.Connect = "odbc;dsn=sappnew;"
    Data2.RecordSource = "select hc_prescrip.id,hc_prescrip.hc_fecha,hc_prescrip.hc_descrip,hc_prescrip.hc_tippresd," & _
    "hc_prescrip.hc_comfec,hc_prescrip.hc_hastaf,hc_prescrip.hc_indicanom,hc_prescrip.hc_mat,clientes.cl_codigo,clientes.cl_apellid " & _
    "from hc_prescrip inner join clientes on hc_prescrip.hc_mat=clientes.cl_codigo where " & _
    "hc_prescrip.hc_tippresd in ('MEDICACION','RECETA PACIENTE CRONICO')" & _
    " and hc_prescrip.hc_fecentrega is null and hc_prescrip.hc_codmedica is not null and hc_fecha >=#" & Format("01/07/2020", "yyyy/mm/dd") & "# order by hc_prescrip.hc_comfec"
    Data2.Refresh
    
    Xcount = 1
    ListView1.ListItems.Clear
    If Data2.Recordset.RecordCount <> 0 Then
       Data2.Recordset.MoveFirst
       Do While Not Data2.Recordset.EOF
'          If Format(Data2.Recordset("hc_hastaf"), "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                ListView1.ListItems.Add Xcount, , Data2.Recordset("hc_mat")
                If IsNull(Data2.Recordset("cl_apellid")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("cl_apellid")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
                End If
                If IsNull(Data2.Recordset("hc_comfec")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(Data2.Recordset("hc_comfec"), "dd/mm/yyyy")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
                End If
                If IsNull(Data2.Recordset("hc_hastaf")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(Data2.Recordset("hc_hastaf"), "dd/mm/yyyy")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
                End If
                If IsNull(Data2.Recordset("hc_descrip")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("hc_descrip")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
                End If
                If IsNull(Data2.Recordset("hc_tippresd")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("hc_tippresd")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/D"
                End If
                If IsNull(Data2.Recordset("hc_fecha")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(Data2.Recordset("hc_fecha"), "dd/mm/yyyy")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
                End If
                If IsNull(Data2.Recordset("hc_indicanom")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("hc_indicanom")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/M"
                End If
                If IsNull(Data2.Recordset("id")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("id")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
                End If
                
                Xcount = Xcount + 1
          'End If
          Data2.Recordset.MoveNext
       Loop
    Else
       frm_medpendv.MousePointer = 0
        MsgBox "No existe prescripción de medicación para este socio.", vbInformation, "Medicación"
    End If
    frm_medpendv.MousePointer = 0

Else
    frm_medpendv.MousePointer = 11

    Data2.Connect = "odbc;dsn=sappnew;"
    Data2.RecordSource = "select hc_prescrip.id,hc_prescrip.hc_fecha,hc_prescrip.hc_descrip,hc_prescrip.hc_tippresd," & _
    "hc_prescrip.hc_comfec,hc_prescrip.hc_hastaf,hc_prescrip.hc_indicanom,hc_prescrip.hc_mat,clientes.cl_codigo,clientes.cl_apellid " & _
    "from hc_prescrip inner join clientes on hc_prescrip.hc_mat=clientes.cl_codigo where " & _
    "hc_prescrip.hc_tippresd in ('MEDICACION','RECETA PACIENTE CRONICO')" & _
    " and hc_prescrip.hc_fecentrega is null and hc_prescrip.hc_codmedica is not null and hc_fecha >=#" & Format("01/07/2020", "yyyy/mm/dd") & "# order by hc_prescrip.hc_comfec"
    Data2.Refresh
    
    Xcount = 1
    ListView1.ListItems.Clear
    If Data2.Recordset.RecordCount <> 0 Then
       Data2.Recordset.MoveFirst
       Do While Not Data2.Recordset.EOF
          If Format(Data2.Recordset("hc_hastaf"), "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                ListView1.ListItems.Add Xcount, , Data2.Recordset("hc_mat")
                If IsNull(Data2.Recordset("cl_apellid")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("cl_apellid")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
                End If
                If IsNull(Data2.Recordset("hc_comfec")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(Data2.Recordset("hc_comfec"), "dd/mm/yyyy")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
                End If
                If IsNull(Data2.Recordset("hc_hastaf")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(Data2.Recordset("hc_hastaf"), "dd/mm/yyyy")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
                End If
                If IsNull(Data2.Recordset("hc_descrip")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("hc_descrip")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
                End If
                If IsNull(Data2.Recordset("hc_tippresd")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("hc_tippresd")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/D"
                End If
                If IsNull(Data2.Recordset("hc_fecha")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(Data2.Recordset("hc_fecha"), "dd/mm/yyyy")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
                End If
                If IsNull(Data2.Recordset("hc_indicanom")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("hc_indicanom")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/M"
                End If
                If IsNull(Data2.Recordset("id")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("id")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
                End If
                
                Xcount = Xcount + 1
          End If
          Data2.Recordset.MoveNext
       Loop
    Else
        frm_medpendv.MousePointer = 0
        MsgBox "No existe prescripción de medicación para este socio.", vbInformation, "Medicación"
    End If
    frm_medpendv.MousePointer = 0

End If
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "OTROS" Then
   Text1.Visible = True
   Text1.SetFocus
Else
   Text1.Text = ""
   Text1.Visible = False
End If

End Sub

Private Sub Command1_Click()
Data2.RecordSource = "select hc_prescrip.id,hc_prescrip.hc_fecha,hc_prescrip.hc_descrip,hc_prescrip.hc_tippresd," & _
"hc_prescrip.hc_comfec,hc_prescrip.hc_hastaf,hc_prescrip.hc_indicanom,hc_prescrip.hc_mat,clientes.cl_codigo,clientes.cl_apellid " & _
"from hc_prescrip inner join clientes on hc_prescrip.hc_mat=clientes.cl_codigo where " & _
"hc_prescrip.hc_tippresd in ('MEDICACION','RECETA PACIENTE CRONICO')" & _
" and hc_prescrip.hc_fecentrega is null and hc_prescrip.hc_codmedica is not null order by hc_prescrip.hc_comfec"
Data2.Refresh
If Data3.Recordset.RecordCount > 0 Then
   Data3.Recordset.MoveFirst
   Do While Not Data3.Recordset.EOF
      Data3.Recordset.Delete
      Data3.Recordset.MoveNext
   Loop
End If
If Data2.Recordset.RecordCount > 0 Then
   Data2.Recordset.MoveFirst
   Do While Not Data2.Recordset.EOF
      If Format(Data2.Recordset("hc_hastaf"), "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
         Data3.Recordset.AddNew
         Data3.Recordset("fecha") = Data2.Recordset("hc_fecha")
         Data3.Recordset("realizada") = Data2.Recordset("hc_hastaf")
         Data3.Recordset("nom_prod") = Mid(Data2.Recordset("hc_descrip"), 1, 50)
         Data3.Recordset("cod_cli") = Data2.Recordset("hc_mat")
         Data3.Recordset("nom_cli") = Mid(Data2.Recordset("cl_apellid"), 1, 30)
         Data3.Recordset("nom_med_a") = Mid(Data2.Recordset("hc_indicanom"), 1, 40)
         Data3.Recordset.Update
      End If
      Data2.Recordset.MoveNext
   Loop
   MsgBox "Terminado"
   
   cr1.ReportFileName = App.path & "\infmedpedv.rpt"
   cr1.ReportTitle = "Medicación pendiente de entrega que está vencida"
   cr1.Action = 1
Else
   MsgBox "No hay registros, verifique"
End If

End Sub

Private Sub Command2_Click()


End Sub

Private Sub Command3_Click()
Dim seguroque As String
Dim Xind, Xcuantos As Integer
Dim Xdias As Integer

Xcuantos = 0
Xind = 0
If Combo1.ListIndex >= 0 Then
    Command3.Enabled = False
    For Xind = 1 To ListView1.ListItems.count
        ListView1.ListItems(Xind).Selected = True
        If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
    '           Xdias = DateDiff("d", Format(ListView1.SelectedItem.ListSubItems(6).Text, "yyyy/mm/dd"), Date)
           If WElusuario = "JFERNAN" Then
                 Data1.Recordset.AddNew
                 Data1.Recordset("idsel") = Val(ListView1.SelectedItem.ListSubItems(8).Text)
                 Data1.Recordset("fecha") = Format(ListView1.SelectedItem.ListSubItems(6).Text, "dd/mm/yyyy")
                 Data1.Recordset("base") = data_parse.Recordset("base")
                 Data1.Recordset("mat") = Val(ListView1.SelectedItem.Text)
                 Data1.Recordset.Update
                 Xcuantos = Xcuantos + 1
           
           Else
              If Format(ListView1.SelectedItem.ListSubItems(3).Text, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                 Data1.Recordset.AddNew
                 Data1.Recordset("idsel") = Val(ListView1.SelectedItem.ListSubItems(8).Text)
                 Data1.Recordset("fecha") = Format(ListView1.SelectedItem.ListSubItems(6).Text, "dd/mm/yyyy")
                 Data1.Recordset("base") = data_parse.Recordset("base")
                 Data1.Recordset("mat") = Val(ListView1.SelectedItem.Text)
                 Data1.Recordset.Update
                 Xcuantos = Xcuantos + 1
              Else
                 MsgBox "Hay un registro seleccionado que aún no está vencido, no se procesará dicho registro", vbInformation
              End If
           End If
        End If
    Next Xind
    
    Data1.Refresh
    Xind = 0
    If Xcuantos <= 0 Then
       MsgBox "No hay registros seleccionados"
    Else
       seguroque = MsgBox("Desea bajar la medicación seleccionada?", vbInformation + vbYesNo)
       If seguroque = vbYes Then
          frm_medpendv.MousePointer = 11
          If Data1.Recordset.RecordCount > 0 Then
             Data1.Recordset.MoveFirst
             Do While Not Data1.Recordset.EOF
                Data2.RecordSource = "select * from hc_prescrip where id =" & Data1.Recordset("idsel") & " and hc_mat =" & Data1.Recordset("mat") & " and hc_tippresd in ('MEDICACION','RECETA PACIENTE CRONICO') and hc_fecentrega is null and hc_codmedica is not null order by hc_comfec"
                Data2.Refresh
                If Data2.Recordset.RecordCount > 0 Then
                   Data2.Recordset.Edit
                   Data2.Recordset("hc_fecentrega") = Date
                   Data2.Recordset("hc_baseent") = data_parse.Recordset("base")
                   Data2.Recordset("hc_usuarioent") = "AUT/" & Mid(WElusuario, 1, 41)
                   If Combo1.Text = "OTROS" Then
                      Data2.Recordset("motivo_cance") = Combo1.Text & "/" & Text1.Text
                   Else
                      Data2.Recordset("motivo_cance") = Combo1.Text
                   End If
                   Data2.Recordset.Update
                End If
                Data1.Recordset.MoveNext
             Loop
          End If
          frm_medpendv.MousePointer = 0
          MsgBox "Proceso terminado"
       End If
    End If
    
    Data1.DatabaseName = App.path & "\selec.mdb"
    Data1.RecordSource = "selec"
    Data1.Refresh
    If Data1.Recordset.RecordCount > 0 Then
       Data1.Recordset.MoveFirst
       Do While Not Data1.Recordset.EOF
          Data1.Recordset.Delete
          Data1.Recordset.MoveNext
       Loop
    End If
    Data2.Connect = "odbc;dsn=sappnew;"
    Data2.RecordSource = "select hc_prescrip.id,hc_prescrip.hc_fecha,hc_prescrip.hc_descrip,hc_prescrip.hc_tippresd," & _
    "hc_prescrip.hc_comfec,hc_prescrip.hc_hastaf,hc_prescrip.hc_indicanom,hc_prescrip.hc_mat,clientes.cl_codigo,clientes.cl_apellid " & _
    "from hc_prescrip inner join clientes on hc_prescrip.hc_mat=clientes.cl_codigo where " & _
    "hc_prescrip.hc_tippresd in ('MEDICACION','RECETA PACIENTE CRONICO')" & _
    " and hc_prescrip.hc_fecentrega is null and hc_prescrip.hc_codmedica is not null and hc_fecha >=#" & Format("01/07/2020", "yyyy/mm/dd") & "# order by hc_prescrip.hc_comfec"
    Data2.Refresh
    
    Dim Xcount As Long
    Xcount = 1
    ListView1.ListItems.Clear
    If Data2.Recordset.RecordCount <> 0 Then
       Data2.Recordset.MoveFirst
       Do While Not Data2.Recordset.EOF
          If Format(Data2.Recordset("hc_hastaf"), "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
            ListView1.ListItems.Add Xcount, , Data2.Recordset("hc_mat")
            If IsNull(Data2.Recordset("cl_apellid")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("cl_apellid")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
            End If
            If IsNull(Data2.Recordset("hc_comfec")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(Data2.Recordset("hc_comfec"), "dd/mm/yyyy")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
            End If
            If IsNull(Data2.Recordset("hc_hastaf")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(Data2.Recordset("hc_hastaf"), "dd/mm/yyyy")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
            End If
            If IsNull(Data2.Recordset("hc_descrip")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("hc_descrip")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(Data2.Recordset("hc_tippresd")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("hc_tippresd")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/D"
            End If
            If IsNull(Data2.Recordset("hc_fecha")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(Data2.Recordset("hc_fecha"), "dd/mm/yyyy")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
            End If
            If IsNull(Data2.Recordset("hc_indicanom")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("hc_indicanom")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/M"
            End If
            If IsNull(Data2.Recordset("id")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("id")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
            End If
            Xcount = Xcount + 1
          End If
          Data2.Recordset.MoveNext
       Loop
    Else
        MsgBox "No existe prescripción de medicación para este socio.", vbInformation, "Medicación"
    End If
    Command3.Enabled = True
Else
    MsgBox "Debe seleccionar motivo de baja", vbInformation
    
End If

End Sub

Private Sub Form_Load()
Dim Xnom As String
data_motivos.Connect = "odbc;dsn=sappnew;"
data_motivos.RecordSource = "select * from motmedic"
data_motivos.Refresh
If data_motivos.Recordset.RecordCount > 0 Then
   data_motivos.Recordset.MoveFirst
   Do While Not data_motivos.Recordset.EOF
      Combo1.AddItem data_motivos.Recordset("descrip")
      data_motivos.Recordset.MoveNext
   Loop
End If

data_parse.DatabaseName = App.path & "\PARSE.mdb"
data_parse.RecordSource = "parsec0"
data_parse.Refresh

Data3.DatabaseName = App.path & "\informes.mdb"
Data3.RecordSource = "infvtas"
Data3.Refresh

Data1.DatabaseName = App.path & "\selec.mdb"
Data1.RecordSource = "selec"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      Data1.Recordset.Delete
      Data1.Recordset.MoveNext
   Loop
End If

Data2.Connect = "odbc;dsn=sappnew;"
Data2.RecordSource = "select hc_prescrip.id,hc_prescrip.hc_fecha,hc_prescrip.hc_descrip,hc_prescrip.hc_tippresd," & _
"hc_prescrip.hc_comfec,hc_prescrip.hc_hastaf,hc_prescrip.hc_indicanom,hc_prescrip.hc_mat,clientes.cl_codigo,clientes.cl_apellid " & _
"from hc_prescrip inner join clientes on hc_prescrip.hc_mat=clientes.cl_codigo where " & _
"hc_prescrip.hc_tippresd in ('MEDICACION','RECETA PACIENTE CRONICO')" & _
" and hc_prescrip.hc_fecentrega is null and hc_prescrip.hc_codmedica is not null and hc_fecha >=#" & Format("01/07/2020", "yyyy/mm/dd") & "# order by hc_prescrip.hc_comfec"
Data2.Refresh

Dim Xcount As Long
Xcount = 1
ListView1.ListItems.Clear
If Data2.Recordset.RecordCount <> 0 Then
   Data2.Recordset.MoveFirst
   Do While Not Data2.Recordset.EOF
      If Format(Data2.Recordset("hc_hastaf"), "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
            ListView1.ListItems.Add Xcount, , Data2.Recordset("hc_mat")
            If IsNull(Data2.Recordset("cl_apellid")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("cl_apellid")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
            End If
            If IsNull(Data2.Recordset("hc_comfec")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(Data2.Recordset("hc_comfec"), "dd/mm/yyyy")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
            End If
            If IsNull(Data2.Recordset("hc_hastaf")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(Data2.Recordset("hc_hastaf"), "dd/mm/yyyy")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
            End If
            If IsNull(Data2.Recordset("hc_descrip")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("hc_descrip")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(Data2.Recordset("hc_tippresd")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("hc_tippresd")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/D"
            End If
            If IsNull(Data2.Recordset("hc_fecha")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(Data2.Recordset("hc_fecha"), "dd/mm/yyyy")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/F"
            End If
            If IsNull(Data2.Recordset("hc_indicanom")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("hc_indicanom")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/M"
            End If
            If IsNull(Data2.Recordset("id")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("id")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
            End If
            
            Xcount = Xcount + 1
      End If
      Data2.Recordset.MoveNext
   Loop
Else
    MsgBox "No existe prescripción de medicación para este socio.", vbInformation, "Medicación"
End If

End Sub
