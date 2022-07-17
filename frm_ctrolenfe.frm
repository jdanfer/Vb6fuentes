VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_ctrolenfe 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Actos de enfermería"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_ctrolenfe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   10215
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Cerrar"
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSComctlLib.ListView ltv1 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "Nombre"
         Object.Width           =   6244
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "b"
         Text            =   "Servicio"
         Object.Width           =   6421
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "c"
         Text            =   "Cod_Serv."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "d"
         Text            =   "No.Fact."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "e"
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "f"
         Text            =   "Hora"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "g"
         Text            =   "Matrícula"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Para chequear el acto cómo realizado, PRIMERO: SELECCIONE EL REGISTRO y luego MARQUE EL VISTO EN LA CASILLA CORRESPONDIENTE."
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   8175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "PACIENTES PENDIENTES PARA ENFERMERIA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frm_ctrolenfe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
Dim Xcount As Long
Dim Xedtiene As Long
Xcount = 1
Data1.DatabaseName = App.Path & "\sapp.mdb"
Data2.DatabaseName = App.Path & "\sapp.mdb"

Data1.RecordSource = "Select * from linmmdd where base =" & frm_menu.data_parse.Recordset("base") & " and servicio =" & 0 & " order by hora"
Data1.Refresh
ListView1.ListItems.Clear
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      If Data1.Recordset("cod_prod") >= 20001 And Data1.Recordset("cod_prod") <= 20020 Then
         If IsNull(Data1.Recordset("cod_cli")) = False Then
            ListView1.ListItems.Add Xcount, , Data1.Recordset("nom_cli")
         Else
            ListView1.ListItems.Add Xcount, , "NN"
         End If
         If IsNull(Data1.Recordset("nom_prod")) = True Then
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SIN DATOS"
         Else
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nom_prod")
         End If
         If IsNull(Data1.Recordset("cod_prod")) = True Then
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
         Else
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("cod_prod")
         End If
         If IsNull(Data1.Recordset("factura")) = True Then
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
         Else
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("factura")
         End If
         If IsNull(Data1.Recordset("fecha")) = True Then
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
         Else
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("fecha")
         End If
         If IsNull(Data1.Recordset("hora")) = True Then
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
         Else
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("hora")
         End If
         If IsNull(Data1.Recordset("cod_cli")) = True Then
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
         Else
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("cod_cli")
         End If
         Xcount = Xcount + 1
      End If
      Data1.Recordset.MoveNext
   Loop
   Label3.Caption = Xcount - 1
Else
   MsgBox "No existe historial", vbInformation, "Ver historial"
   Label3.Caption = 0
End If
'End If

End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'MsgBox "El item está cheq"
Dim XNfac, XNmat, XNcod As Long
Dim XNfe As Date

If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then
'   MsgBox "Esta"
'   MsgBox "ES:" & ListView1.SelectedItem.ListSubItems(1).Text
   XNcod = ListView1.SelectedItem.ListSubItems(2).Text
   XNfac = ListView1.SelectedItem.ListSubItems(3).Text
   XNfe = ListView1.SelectedItem.ListSubItems(4).Text
   XNmat = ListView1.SelectedItem.ListSubItems(5).Text
   Data2.RecordSource = "Select * from linmmdd where base =" & frm_menu.data_parse.Recordset("base") & " And fecha =#" & Format(XNfe, "yyyy/mm/dd") & "# And factura =" & XNfac & " And cod_prod =" & XNcod
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      If IsNull(Data2.Recordset("servicio")) = True Then
         Data2.Recordset.Edit
         Data2.Recordset("servicio") = 1
         Data2.Recordset("nom_superv") = Format(Time, "HH:mm")
         Data2.Recordset.Update
      Else
         If Data2.Recordset("servicio") = 1 Then
         Else
            Data2.Recordset.Edit
            Data2.Recordset("servicio") = 1
            Data2.Recordset("nom_superv") = Format(Time, "HH:mm")
            Data2.Recordset.Update
         End If
      End If
   End If
Else
   XNcod = ListView1.SelectedItem.ListSubItems(2).Text
   XNfac = ListView1.SelectedItem.ListSubItems(3).Text
   XNfe = ListView1.SelectedItem.ListSubItems(4).Text
   XNmat = ListView1.SelectedItem.ListSubItems(5).Text
   Data2.RecordSource = "Select * from linmmdd where base =" & frm_menu.data_parse.Recordset("base") & " And fecha =#" & Format(XNfe, "yyyy/mm/dd") & "# And factura =" & XNfac & " And cod_prod =" & XNcod
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      If IsNull(Data2.Recordset("servicio")) = True Then
         Data2.Recordset.Edit
         Data2.Recordset("servicio") = 0
         Data2.Recordset("nom_superv") = Format(Time, "99:99")
         Data2.Recordset.Update
      Else
         If Data2.Recordset("servicio") = 0 Then
         Else
            Data2.Recordset.Edit
            Data2.Recordset("servicio") = 0
            Data2.Recordset("nom_superv") = Format(Time, "99:99")
            Data2.Recordset.Update
         End If
      End If
   End If

End If


End Sub

Private Sub ltv1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub ltv1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'MsgBox "El item está cheq"
Dim XNfac, XNmat, XNcod As Long
Dim XNfe As Date

If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then
'   MsgBox "Esta"
'   MsgBox "ES:" & ListView1.SelectedItem.ListSubItems(1).Text
   XNcod = ListView1.SelectedItem.ListSubItems(2).Text
   XNfac = ListView1.SelectedItem.ListSubItems(3).Text
   XNfe = ListView1.SelectedItem.ListSubItems(4).Text
   XNmat = ListView1.SelectedItem.ListSubItems(5).Text
   Data2.RecordSource = "Select * from linmmdd where base =" & frm_menu.data_parse.Recordset("base") & " And fecha =#" & Format(XNfe, "yyyy/mm/dd") & "# And factura =" & XNfac & " And cod_prod =" & XNcod
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      If IsNull(Data2.Recordset("servicio")) = True Then
         Data2.Recordset.Edit
         Data2.Recordset("servicio") = 1
         Data2.Recordset("nom_superv") = Format(Time, "HH:mm")
         Data2.Recordset.Update
      Else
         If Data2.Recordset("servicio") = 1 Then
         Else
            Data2.Recordset.Edit
            Data2.Recordset("servicio") = 1
            Data2.Recordset("nom_superv") = Format(Time, "HH:mm")
            Data2.Recordset.Update
         End If
      End If
   End If
Else
   XNcod = ListView1.SelectedItem.ListSubItems(2).Text
   XNfac = ListView1.SelectedItem.ListSubItems(3).Text
   XNfe = ListView1.SelectedItem.ListSubItems(4).Text
   XNmat = ListView1.SelectedItem.ListSubItems(5).Text
   Data2.RecordSource = "Select * from linmmdd where base =" & frm_menu.data_parse.Recordset("base") & " And fecha =#" & Format(XNfe, "yyyy/mm/dd") & "# And factura =" & XNfac & " And cod_prod =" & XNcod
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      If IsNull(Data2.Recordset("servicio")) = True Then
         Data2.Recordset.Edit
         Data2.Recordset("servicio") = 0
         Data2.Recordset("nom_superv") = Format(Time, "99:99")
         Data2.Recordset.Update
      Else
         If Data2.Recordset("servicio") = 0 Then
         Else
            Data2.Recordset.Edit
            Data2.Recordset("servicio") = 0
            Data2.Recordset("nom_superv") = Format(Time, "99:99")
            Data2.Recordset.Update
         End If
      End If
   End If

End If
End Sub
