VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_buscamedica 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar medicación"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9885
   Icon            =   "frm_buscamedica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_parse 
      Caption         =   "data_parse"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_motivos 
      Caption         =   "data_motivos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos para la baja"
      Height          =   1455
      Left            =   2280
      TabIndex        =   6
      Top             =   4080
      Width           =   7455
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   4095
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
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton Command4 
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
         Left            =   360
         Picture         =   "frm_buscamedica.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Autorizar excepción"
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
      Left            =   2520
      Picture         =   "frm_buscamedica.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_buscamedica.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Obtener los medicamentos seleccionados"
      Top             =   4080
      Width           =   615
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   6376
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FEC.INICIAL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "FEC.FINAL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "MEDICACIÓN"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "TIPO PRESCRIP"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "FECHA HC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "MEDICO"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "odbc;dsn=sappnew;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   2295
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
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox t_busca 
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
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Buscar por genérico:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frm_buscamedica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DBGrid1_DblClick()
If IsNull(Data1.Recordset("nombre")) = False Then
   Xdescmedic = Data1.Recordset("nombre")
   XcodelMedicamento = Data1.Recordset("id")
   IdMatMed = Val(frm_factura.labmatri.Caption)
   IdTablaPres = 0
Else
  Xdescmedic = ""
  XcodelMedicamento = 0
  IdMatMed = 0
  IdTablaPres = 0
End If

Unload Me

End Sub

Private Sub DBGrid2_DblClick()
Dim MedDesde, MedHasta, Medretiraantes As Date

MedDesde = Date - 5
MedHasta = Date + 1

If IsNull(Data2.Recordset("hc_comfec")) = False Then
   Medretiraantes = Data2.Recordset("hc_comfec") - 5
   If Format(Data2.Recordset("hc_comfec"), "yyyy/mm/dd") >= Format(MedDesde, "yyyy/mm/dd") Then
      Xdescmedic = Data2.Recordset("hc_descrip")
      If IsNull(Data2.Recordset("hc_codmedica")) = False Then
         XcodelMedicamento = Data2.Recordset("hc_codmedica")
         IdTablaPres = Data2.Recordset("id")
         IdMatMed = Data2.Recordset("hc_mat")
      Else
         XcodelMedicamento = 0
         IdTablaPres = 0
         IdMatMed = 0
      End If
   Else
      If IsNull(Data2.Recordset("hc_hastaf")) = False Then
         If Format(Data2.Recordset("hc_hastaf"), "yyyy/mm/dd") <= Format(MedHasta, "yyyy/mm/dd") Then
            Xdescmedic = Data2.Recordset("hc_descrip")
            If IsNull(Data2.Recordset("hc_codmedica")) = False Then
               XcodelMedicamento = Data2.Recordset("hc_codmedica")
               IdTablaPres = Data2.Recordset("id")
               IdMatMed = Data2.Recordset("hc_mat")
            Else
               XcodelMedicamento = 0
               IdTablaPres = 0
               IdMatMed = 0
            End If
         Else
            MsgBox "No está en fecha para retiro. Verifique.", vbInformation
            Xdescmedic = ""
            XcodelMedicamento = 0
            IdTablaPres = 0
            IdMatMed = 0
         End If
      Else
         MsgBox "No figura fecha de retiro. Consulte con el médico", vbInformation
         Xdescmedic = ""
         XcodelMedicamento = 0
         IdTablaPres = 0
         IdMatMed = 0
      End If
   End If
Else
   Xdescmedic = ""
   XcodelMedicamento = 0
   IdTablaPres = 0
   IdMatMed = 0
End If

Unload Me

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
Dim Xind, Xcuantos, Xdias As Integer

Xind = 0
Xcuantos = 0
'If Data1.Recordset.RecordCount > 0 Then
 '  Data1.Recordset.MoveFirst
 '  Do While Not Data1.Recordset.EOF
   '   Data1.Recordset.Delete
  '    Data1.Recordset.MoveNext
   'Loop
'End If

For Xind = 1 To ListView1.ListItems.count
    ListView1.ListItems(Xind).Selected = True
    If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
       Xdias = DateDiff("d", Format(ListView1.SelectedItem.ListSubItems(1).Text, "dd/mm/yyyy"), Date)
        
'       If Format(ListView1.SelectedItem.ListSubItems(5).Text, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
       If Xdias > -6 And Xdias <= 30 Then
          Data1.Recordset.AddNew
          Data1.Recordset("idsel") = Val(ListView1.SelectedItem.Text)
          Data1.Recordset("fecha") = Format(ListView1.SelectedItem.ListSubItems(5).Text, "dd/mm/yyyy")
          Data1.Recordset("base") = frm_menu.data_parse.Recordset("base")
          Data1.Recordset("mat") = Val(frm_factura.labmatri.Caption)
          Data1.Recordset.Update
          Xconvprom = ListView1.SelectedItem.ListSubItems(3).Text
          Xcuantos = Xcuantos + 1
       Else
          MsgBox "Hay medicación que no está en fecha de retiro, no se procesará. Verifique!", vbExclamation
       End If
    End If
Next Xind
Data1.Refresh
Xind = 0
If Xcuantos <= 0 Then
   MsgBox "No hay medicación seleccionada", vbInformation
   Xconvprom = ""

End If
Unload Me



End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
Dim Xelqueautoriza As String
Dim seguroque As String
Dim Xind, Xcuantos As Integer

Xcuantos = 0
Xind = 0

Xelqueautoriza = InputBox("Ingrese nombre de responsable que autoriza")
If Trim(Xelqueautoriza) <> "" Then


End If

End Sub

Private Sub Command4_Click()
Dim seguroque As String
Dim Xind, Xcuantos As Integer
Dim Xdias As Integer

Xcuantos = 0
Xind = 0

If Combo1.ListIndex >= 0 Then
    Command4.Enabled = False
    For Xind = 1 To ListView1.ListItems.count
        ListView1.ListItems(Xind).Selected = True
        If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
              If Format(ListView1.SelectedItem.ListSubItems(2).Text, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
                 Data1.Recordset.AddNew
                 Data1.Recordset("idsel") = Val(ListView1.SelectedItem.Text)
                 Data1.Recordset("fecha") = Format(ListView1.SelectedItem.ListSubItems(5).Text, "dd/mm/yyyy")
                 Data1.Recordset("base") = data_parse.Recordset("base")
                 Data1.Recordset("mat") = Val(frm_factura.labmatri.Caption)
                 Data1.Recordset.Update
                 Xcuantos = Xcuantos + 1
              Else
                 MsgBox "Hay un registro seleccionado que aún no está vencido, no se procesará dicho registro", vbInformation
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
          frm_buscamedica.MousePointer = 11
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
          frm_buscamedica.MousePointer = 0
          MsgBox "Proceso terminado"
       End If
    End If
    
    If Data1.Recordset.RecordCount > 0 Then
       Data1.Recordset.MoveFirst
       Do While Not Data1.Recordset.EOF
          Data1.Recordset.Delete
          Data1.Recordset.MoveNext
       Loop
    End If
    
    Data2.Connect = "odbc;dsn=sappnew;"
    Data2.RecordSource = "select * from hc_prescrip where hc_mat =" & Val(frm_factura.labmatri.Caption) & " and hc_tippresd in ('MEDICACION','RECETA PACIENTE CRONICO') and hc_fecentrega is null and hc_codmedica is not null and hc_fecha >=#" & Format("01/07/2020", "yyyy/mm/dd") & "# order by hc_comfec"
    Data2.Refresh
    
    Dim Xcount, Xsaldo As Long
    Xcount = 1
    ListView1.ListItems.Clear
    If Data2.Recordset.RecordCount <> 0 Then
       Data2.Recordset.MoveFirst
       Do While Not Data2.Recordset.EOF
          ListView1.ListItems.Add Xcount, , Data2.Recordset("id")
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
           Xcount = Xcount + 1
           Data2.Recordset.MoveNext
       Loop
    Else
        MsgBox "No existe prescripción de medicación para este socio.", vbInformation, "Medicación"
    End If
    Command4.Enabled = True
Else
    MsgBox "Debe seleccionar motivo de baja", vbInformation
    
End If
End Sub

Private Sub Form_Load()

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

Data1.DatabaseName = App.path & "\selec.mdb"
Data1.RecordSource = "selec"
Data1.Refresh
'If Data1.Recordset.RecordCount > 0 Then
'   Data1.Recordset.MoveFirst
'   Do While Not Data1.Recordset.EOF
'      Data1.Recordset.Delete
'      Data1.Recordset.MoveNext
'   Loop
'End If

Data2.Connect = "odbc;dsn=sappnew;"
Data2.RecordSource = "select * from hc_prescrip where hc_mat =" & Val(frm_factura.labmatri.Caption) & " and hc_tippresd in ('MEDICACION','RECETA PACIENTE CRONICO') and hc_fecentrega is null and hc_codmedica is not null and hc_fecha >=#" & Format("01/07/2020", "yyyy/mm/dd") & "# order by hc_comfec"
Data2.Refresh
Xconvprom = ""

Dim Xcount, Xsaldo As Long
Dim Xdias As Integer
Xcount = 1
ListView1.ListItems.Clear
If Data2.Recordset.RecordCount <> 0 Then
   Data2.Recordset.MoveFirst
   Do While Not Data2.Recordset.EOF
       Xdias = DateDiff("d", Format(Data2.Recordset("hc_comfec"), "dd/mm/yyyy"), Date)
       'If Xdias > -6 And Xdias <= 30 Then
            ListView1.ListItems.Add Xcount, , Data2.Recordset("id")
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
            Xcount = Xcount + 1
       'End If
       Data2.Recordset.MoveNext
   Loop
Else
'    MsgBox "No existe prescripción de medicación para este socio.", vbInformation, "Medicación"
End If

End Sub

Private Sub t_busca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If t_busca.Text <> "" Then
       Data2.RecordSource = "select * from hc_prescrip where hc_mat =" & Val(frm_factura.labmatri.Caption) & " and hc_descrip >='" & t_busca.Text & "' and hc_tippresd in ('MEDICACION','RECETA PACIENTE CRONICO') and hc_fecentrega is null and hc_codmedica is not null order by hc_comfec"
       Data2.Refresh
    
        Dim Xcount As Long
        Xcount = 1
        ListView1.ListItems.Clear
        If Data2.Recordset.RecordCount > 0 Then
           Data2.Recordset.MoveFirst
           Do While Not Data2.Recordset.EOF
               ListView1.ListItems.Add Xcount, , Data2.Recordset("id")
               If IsNull(Data2.Recordset("hc_fecha")) = False Then
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(Data2.Recordset("hc_fecha"), "dd/mm/yyyy")
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
               If IsNull(Data2.Recordset("hc_indicanom")) = False Then
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("hc_indicanom")
               Else
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/M"
               End If
               Xcount = Xcount + 1
               Data2.Recordset.MoveNext
           Loop
        Else
            MsgBox "No existe prescripción de medicación para este socio.", vbInformation, "Medicación"
        End If
    
    End If
End If

         
End Sub
