VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_pendcmtpol 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pendientes CMT Policlínica Med.Gral."
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12630
   Icon            =   "frm_pendcmtpol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   12630
   StartUpPosition =   1  'CenterOwner
   Begin MSMask.MaskEdBox md 
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   5160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00800000&
      Caption         =   "Ver CMT Polic."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   9720
      TabIndex        =   16
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Data data_llam 
      Caption         =   "data_llam"
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
      Top             =   4800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00800000&
      Caption         =   "Ver sólo CMT Despacho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   9720
      TabIndex        =   13
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00800000&
      Caption         =   "Ver sólo Presencial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   9720
      TabIndex        =   12
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox t_base 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   4200
      Width           =   975
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   5520
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11880
      Picture         =   "frm_pendcmtpol.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprimir pendientes de base"
      Top             =   4200
      Width           =   495
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11880
      Picture         =   "frm_pendcmtpol.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Actualizar pantalla"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8880
      Picture         =   "frm_pendcmtpol.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3600
      Width           =   4575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   5318
      View            =   3
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Hora"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Matrícula"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nombre"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Convenio"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Médico Asignado"
         Object.Width           =   3951
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "CodMed"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Contacto"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Consulta"
         Object.Width           =   3951
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Nro.Reg."
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Repetición?"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Base"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Zona"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      Caption         =   "FECHA DESDE:"
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
      TabIndex        =   17
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label labzon 
      Height          =   255
      Left            =   4320
      TabIndex        =   15
      Top             =   4560
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label labcodmedcmt 
      Height          =   255
      Left            =   7080
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      Caption         =   "BASE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label labnommed 
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label labcodmed 
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
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
      Left            =   240
      TabIndex        =   6
      Top             =   4680
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Seleccione el médico a agendar:"
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
      TabIndex        =   2
      Top             =   3600
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Pendientes CMT y Policlínica presencial de Med.Gral."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   12135
   End
End
Attribute VB_Name = "frm_pendcmtpol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check2_Click()
If Check2.Value = 1 Then
   If Check3.Value = 1 Then
      Check3.Value = 0
   End If
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
   If Check2.Value = 1 Then
      Check2.Value = 0
   End If
End If

End Sub

Private Sub Combo1_LostFocus()
Buscar_medico
Buscar_medicoCmt

End Sub

Private Sub Command1_Click()
Dim Xdeseacambiar As String
Dim Xmotivocmt As String
Dim Xcountt, Xind, Xnrodoc, Xmatcmt As Long
Dim Yalotiene As String

On Error GoTo Quepasaalmodif

If labcodmed.Caption <> "" Then
   Xdeseacambiar = MsgBox("Desea asignar el médico:" & labnommed.Caption & " Código:" & labcodmed.Caption & " a los registros seleccionados?", vbInformation + vbYesNo, "SAPP")
   If Xdeseacambiar = vbYes Then
''      Xmotivocmt = InputBox("Ingrese motivo del cambio", "SAPP")
      Xmotivocmt = "PASAR CMT"
      If Trim(Xmotivocmt) <> "" Then
         Xcountt = 1
         Xind = 0
         frm_pendcmtpol.MousePointer = 11
         For Xind = 1 To ListView1.ListItems.count
              ListView1.ListItems(Xind).Selected = True
              If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
        '     MsgBox "Chequeado"
                 If ListView1.SelectedItem.ListSubItems(8).Text = "CMT DESPACHO" Then
                    If Trim(labcodmedcmt.Caption) <> "" Then
                       Xnrodoc = Val(ListView1.SelectedItem.ListSubItems(9).Text)
                       Data2.RecordSource = "select * from llamado where nrolla =" & Xnrodoc
                       Data2.Refresh
                       If Data2.Recordset.RecordCount > 0 Then
                          If IsNull(Data2.Recordset("cmt_usproc")) = False Then
                             Yalotiene = MsgBox("Lo está llamando el médico: " & Data2.Recordset("cmt_usproc") & " Desea igual reasignar?", vbExclamation + vbYesNo, "CMT")
                             If Yalotiene = vbYes Then
                                If IsNull(Data2.Recordset("codmedcmt")) = False Then
                                   If Val(Data2.Recordset("codmedcmt")) <> Val(labcodmedcmt.Caption) Then
                                      Data2.Recordset.Edit
                                      Data2.Recordset("codmedcmt") = Val(labcodmedcmt.Caption)
                                      Data2.Recordset.Update
                                      Data3.RecordSource = "select * from cambios_cmt where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
                                      Data3.Refresh
                                      Data3.Recordset.AddNew
                                      Data3.Recordset("fecha") = Date
                                      Data3.Recordset("hora") = Format(Time, "HH:mm")
                                      Data3.Recordset("factura") = Xnrodoc
                                      Data3.Recordset("motivo") = Xmotivocmt & " U:" & WElusuario
                                      Data3.Recordset("base") = frm_menu.data_parse.Recordset("base")
                                      Data3.Recordset.Update
                                   End If
                                Else
                                   Data2.Recordset.Edit
                                   Data2.Recordset("codmedcmt") = Val(labcodmedcmt.Caption)
                                   Data2.Recordset.Update
                                   Data3.RecordSource = "select * from cambios_cmt where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
                                   Data3.Refresh
                                   Data3.Recordset.AddNew
                                   Data3.Recordset("fecha") = Date
                                   Data3.Recordset("hora") = Format(Time, "HH:mm")
                                   Data3.Recordset("factura") = Xnrodoc
                                   Data3.Recordset("motivo") = Xmotivocmt & " U:" & WElusuario
                                   Data3.Recordset("base") = frm_menu.data_parse.Recordset("base")
                                   Data3.Recordset.Update
                                End If
                             End If
                          Else
                             If IsNull(Data2.Recordset("codmedcmt")) = False Then
                                If Val(Data2.Recordset("codmedcmt")) <> Val(labcodmedcmt.Caption) Then
                                   Data2.Recordset.Edit
                                   Data2.Recordset("codmedcmt") = Val(labcodmedcmt.Caption)
                                   Data2.Recordset.Update
                                   Data3.RecordSource = "select * from cambios_cmt where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
                                   Data3.Refresh
                                   Data3.Recordset.AddNew
                                   Data3.Recordset("fecha") = Date
                                   Data3.Recordset("hora") = Format(Time, "HH:mm")
                                   Data3.Recordset("factura") = Xnrodoc
                                   Data3.Recordset("motivo") = Xmotivocmt & " U:" & WElusuario
                                   Data3.Recordset("base") = frm_menu.data_parse.Recordset("base")
                                   Data3.Recordset.Update
                                End If
                             Else
                                Data2.Recordset.Edit
                                Data2.Recordset("codmedcmt") = Val(labcodmedcmt.Caption)
                                Data2.Recordset.Update
                                Data3.RecordSource = "select * from cambios_cmt where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
                                Data3.Refresh
                                Data3.Recordset.AddNew
                                Data3.Recordset("fecha") = Date
                                Data3.Recordset("hora") = Format(Time, "HH:mm")
                                Data3.Recordset("factura") = Xnrodoc
                                Data3.Recordset("motivo") = Xmotivocmt & " U:" & WElusuario
                                Data3.Recordset("base") = frm_menu.data_parse.Recordset("base")
                                Data3.Recordset.Update
                             End If
                          End If
                       End If
                    Else
                       MsgBox "No se encuentra médico para CMT, No se puede actualizar en Despacho.", vbCritical
                    End If
                 Else
                    Xnrodoc = Val(ListView1.SelectedItem.ListSubItems(9).Text)
                    Xmatcmt = Val(ListView1.SelectedItem.ListSubItems(2).Text)
                    Data2.RecordSource = "Select * from linmmdd where factura =" & Xnrodoc & " and cod_cli =" & Xmatcmt
                    Data2.Refresh
                    If Data2.Recordset.RecordCount > 0 Then
                       If IsNull(Data2.Recordset("nro_med_a")) = False Then
                          If IsNull(Data2.Recordset("cmt_usproc")) = False Then
                             Yalotiene = MsgBox("Lo está llamando el médico: " & Data2.Recordset("cmt_usproc") & " Desea igual reasignar?", vbExclamation + vbYesNo, "CMT")
                             If Yalotiene = vbYes Then
                                If Data2.Recordset("nro_med_a") <> Val(labcodmed.Caption) Then
                                   Data2.Recordset.Edit
                                   If Trim(t_base.Text) <> "" Then
                                      If Data2.Recordset("tot_lin") > 0 Then
                                         MsgBox "ANOTE: El socio: " & Data2.Recordset("cod_cli") & " no se puede cambiar de base por tener costo la consulta.", vbCritical
                                      Else
                                         Data2.Recordset("base") = t_base.Text
                                      End If
                                   End If
                                   Data2.Recordset("fecha") = Date
                                   Data2.Recordset("nro_med_a") = Val(labcodmed.Caption)
                                   Data2.Recordset("nom_med_a") = Mid(Trim(labnommed.Caption), 1, 40)
                                   Data2.Recordset.Update
                                   Data3.RecordSource = "select * from cambios_cmt where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
                                   Data3.Refresh
                                   Data3.Recordset.AddNew
                                   Data3.Recordset("fecha") = Date
                                   Data3.Recordset("hora") = Format(Time, "HH:mm")
                                   Data3.Recordset("factura") = Xnrodoc
                                   Data3.Recordset("motivo") = Mid(Xmotivocmt, 1, 90)
                                   Data3.Recordset("base") = frm_menu.data_parse.Recordset("base")
                                   Data3.Recordset.Update
                                End If
                             End If
                          Else
                             If Data2.Recordset("nro_med_a") <> Val(labcodmed.Caption) Then
                                Data2.Recordset.Edit
                                If Trim(t_base.Text) <> "" Then
                                   If Data2.Recordset("tot_lin") > 0 Then
                                      MsgBox "ANOTE: El socio: " & Data2.Recordset("cod_cli") & " no se puede cambiar de base por tener costo la consulta.", vbCritical
                                   Else
                                      Data2.Recordset("base") = t_base.Text
                                   End If
                                End If
                                Data2.Recordset("fecha") = Date
                                Data2.Recordset("nro_med_a") = Val(labcodmed.Caption)
                                Data2.Recordset("nom_med_a") = Mid(Trim(labnommed.Caption), 1, 40)
                                Data2.Recordset.Update
                                Data3.RecordSource = "select * from cambios_cmt where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
                                Data3.Refresh
                                Data3.Recordset.AddNew
                                Data3.Recordset("fecha") = Date
                                Data3.Recordset("hora") = Format(Time, "HH:mm")
                                Data3.Recordset("factura") = Xnrodoc
                                Data3.Recordset("motivo") = Mid(Xmotivocmt, 1, 90)
                                Data3.Recordset("base") = frm_menu.data_parse.Recordset("base")
                                Data3.Recordset.Update
                             End If
                          End If
                       End If
                    End If
                 End If
              End If
         Next Xind
         frm_pendcmtpol.MousePointer = 0
         MsgBox "Proceso terminado"
         Command2_Click
      Else
         frm_pendcmtpol.MousePointer = 0
         MsgBox "Falta ingresar motivo", vbExclamation
      End If
   End If
Else
   MsgBox "Falta seleccionar médico.", vbExclamation
End If

Exit Sub

Quepasaalmodif:
            If Err.Number = 3306 Then
               MsgBox "Error 3306 " & Err.Description
            Else
               MsgBox "ERROR: " & Err.Description
            End If

End Sub

Private Sub Command2_Click()
Dim Xcount As Long
Dim Xven As Date

Xcount = 1

Data3.Connect = "odbc;dsn=" & Xconexrmt & ";"

Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"

Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

If Check3.Value = 1 Then
   ListView1.ListItems.Clear
   frm_pendcmtpol.MousePointer = 11
   Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/05/2022", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by fecha,hora"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         If md.Text = "__/__/____" Then
            Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "# and hora >='" & Data1.Recordset("hora") & "' and cednum =" & Data1.Recordset("ci") & " and tipo_consd in ('Orientación Telefónica')"
         Else
            Data2.RecordSource = "select * from cabezal_hcdig where fecha >=#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "# and hora >='" & Data1.Recordset("hora") & "' and cednum =" & Data1.Recordset("ci") & " and tipo_consd in ('Orientación Telefónica')"
         End If
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
         Else
            ListView1.ListItems.Add Xcount, , Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("hora")
            If IsNull(Data1.Recordset("matric")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("matric")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(Data1.Recordset("nombre")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nombre")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NN"
            End If
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("categ")
            If IsNull(Data1.Recordset("codmedcmt")) = False Then
               Data2.RecordSource = "select * from medicos_esp where id =" & Data1.Recordset("codmedcmt")
               Data2.Refresh
               If Data2.Recordset.RecordCount > 0 Then
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("nom_med")
               Else
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO REGISTRADO"
               End If
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO REGISTRADO"
            End If
            If IsNull(Data1.Recordset("codmedcmt")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("codmedcmt")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(Data1.Recordset("telef")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("telef")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "CMT DESPACHO"
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nrolla")
            If IsNull(Data1.Recordset("repite")) = False Then
               If Data1.Recordset("repite") = 1 Then
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SI"
               Else
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO"
               End If
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO"
            End If
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "19"
            If IsNull(Data1.Recordset("motmov")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("motmov")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "Sin Zona"
            End If
            Xcount = Xcount + 1
         End If
         Data1.Recordset.MoveNext
      Loop
   End If
   Label3.Caption = "Total registros:" & Trim(str(Xcount - 1))
Else
   frm_pendcmtpol.MousePointer = 11
   If Check2.Value = 1 Then
      If md.Text = "__/__/____" Then
         Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod in (10001,10003,10005) order by base"
      Else
         Data1.RecordSource = "select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cod_prod in (10001,10003,10005) order by base"
      End If
   Else
      If Check1.Value = 1 Then
         If md.Text = "__/__/____" Then
            Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod in (10050) order by base"
         Else
            Data1.RecordSource = "select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cod_prod in (10050) order by base"
         End If
      Else
         If md.Text = "__/__/____" Then
            Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod in (10050,10001,10003,10005) order by base"
         Else
            Data1.RecordSource = "select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cod_prod in (10050,10001,10003,10005) order by base"
         End If
''   Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod =" & 10050 & " and base =" & frm_menu.data_parse.Recordset("base")
      End If
   End If
   Data1.Refresh
   Data4.DatabaseName = App.path & "\informes.mdb"

   ListView1.ListItems.Clear
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         If md.Text = "__/__/____" Then
            Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "# and cednum =" & Data1.Recordset("ced_socio")
         Else
            Data2.RecordSource = "select * from cabezal_hcdig where fecha >=#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "# and cednum =" & Data1.Recordset("ced_socio")
         End If
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
         Else
            ListView1.ListItems.Add Xcount, , Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("hora")
            If IsNull(Data1.Recordset("cod_cli")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("cod_cli")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nom_cli")
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("convenio")
            If IsNull(Data1.Recordset("nom_med_a")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nom_med_a")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(Data1.Recordset("nro_med_a")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nro_med_a")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(Data1.Recordset("contact_tel")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("contact_tel")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nom_prod")
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("factura")
            If IsNull(Data1.Recordset("repetir")) = False Then
               If Data1.Recordset("repetir") = "S" Then
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SI"
               Else
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO"
               End If
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO"
            End If
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("base")
            Consulta_zonas
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , labzon.Caption
            
            Xcount = Xcount + 1
         End If
         Data1.Recordset.MoveNext
      Loop
   Else
      frm_pendcmtpol.MousePointer = 0
      MsgBox "No hay registros para CMT"
   End If
   
   If Check2.Value = 1 Or Check1.Value = 1 Then
   Else
        frm_pendcmtpol.MousePointer = 11
        Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/05/2022", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by fecha,hora"
        Data1.Refresh
        If Data1.Recordset.RecordCount > 0 Then
           Data1.Recordset.MoveFirst
           Do While Not Data1.Recordset.EOF
              If md.Text = "__/__/____" Then
                 Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "# and hora >='" & Data1.Recordset("hora") & "' and cednum =" & Data1.Recordset("ci") & " and tipo_consd in ('Orientación Telefónica')"
              Else
                 Data2.RecordSource = "select * from cabezal_hcdig where fecha >=#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "# and hora >='" & Data1.Recordset("hora") & "' and cednum =" & Data1.Recordset("ci") & " and tipo_consd in ('Orientación Telefónica')"
              End If
              Data2.Refresh
              If Data2.Recordset.RecordCount > 0 Then
              Else
                ListView1.ListItems.Add Xcount, , Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
                ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("hora")
                If IsNull(Data1.Recordset("matric")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("matric")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
                End If
                If IsNull(Data1.Recordset("nombre")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nombre")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NN"
                End If
                ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("categ")
                If IsNull(Data1.Recordset("codmedcmt")) = False Then
                   Data2.RecordSource = "select * from medicos_esp where id =" & Data1.Recordset("codmedcmt")
                   Data2.Refresh
                   If Data2.Recordset.RecordCount > 0 Then
                      ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("nom_med")
                   Else
                      ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO REGISTRADO"
                   End If
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO REGISTRADO"
                End If
                If IsNull(Data1.Recordset("codmedcmt")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("codmedcmt")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
                End If
                If IsNull(Data1.Recordset("telef")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("telef")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
                End If
                ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "CMT DESPACHO"
                ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nrolla")
                If IsNull(Data1.Recordset("repite")) = False Then
                   If Data1.Recordset("repite") = 1 Then
                      ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SI"
                   Else
                      ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO"
                   End If
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO"
                End If
                ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "19"
                If IsNull(Data1.Recordset("motmov")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("motmov")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "Sin Zona"
                End If
                Xcount = Xcount + 1
              End If
              Data1.Recordset.MoveNext
           Loop
        Else
           frm_pendcmtpol.MousePointer = 0
           MsgBox "No hay registros para CMT Despacho"
        End If
   End If
   frm_pendcmtpol.MousePointer = 0
   
   Label3.Caption = "Total registros:" & Trim(str(Xcount - 1))
End If
frm_pendcmtpol.MousePointer = 0

End Sub

Private Sub Command3_Click()

frm_pendcmtpol.MousePointer = 11

Data4.RecordSource = "infvtas"
Data4.Refresh
If Data4.Recordset.RecordCount > 0 Then
   Data4.Recordset.MoveFirst
   Do While Not Data4.Recordset.EOF
      Data4.Recordset.Delete
      Data4.Recordset.MoveNext
   Loop
End If

MsgBox "Se imprimen sólo CMT Polic y Polic. Med.Gral (NO Despacho)", vbInformation

'Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod =" & 10050
If Check2.Value = 1 Then
   Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod in (10001,10003,10005) order by base"
Else
   Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod in (10050,10001,10003,10005) order by base"
''   Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod =" & 10050 & " and base =" & frm_menu.data_parse.Recordset("base")
End If

Data1.Refresh

If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "# and cednum =" & Data1.Recordset("ced_socio")
      Data2.Refresh
      If Data2.Recordset.RecordCount > 0 Then
      Else
         Data4.Recordset.AddNew
         Data4.Recordset("cod_cli") = Data1.Recordset("cod_cli")
         Data4.Recordset("nom_cli") = Data1.Recordset("nom_cli")
         Data4.Recordset("fecha") = Data1.Recordset("fecha")
         Data4.Recordset("nro_med_a") = Data1.Recordset("nro_med_a")
         Data4.Recordset("nom_med_a") = Data1.Recordset("nom_med_a")
         Data4.Recordset("base") = Data1.Recordset("base")
         Data4.Recordset("convenio") = Data1.Recordset("convenio")
         If IsNull(Data1.Recordset("contact_tel")) = False Then
            Data4.Recordset("ced_socio") = Val(Data1.Recordset("contact_tel"))
         Else
            Data4.Recordset("ced_socio") = 0
         End If
         Data4.Recordset("cod_prod") = Data1.Recordset("cod_prod")
         Data4.Recordset("nom_prod") = Data1.Recordset("nom_prod")
         Data4.Recordset.Update
      End If
      Data1.Recordset.MoveNext
   Loop
   frm_pendcmtpol.MousePointer = 0
   cr1.ReportFileName = App.path & "\infvtasxser.rpt"
   cr1.ReportTitle = "INFORME DE CMT PENDIENTES: " & Format(Date, "dd/mm/yyyy") & " USUARIO:" & WElusuario
   cr1.Action = 1
Else
    frm_pendcmtpol.MousePointer = 0
    MsgBox "No hay registros de CMT Pendientes."
End If

frm_pendcmtpol.MousePointer = 0

End Sub



Private Sub Form_Load()
Dim Xcount As Long
Dim Xven As Date

Xcount = 1

Data3.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_llam.Connect = "odbc;dsn=" & Xconexrmt & ";"

Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"

Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

If md.Text = "__/__/____" Then
   Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod in (10050,10001,10003,10005) order by base"
Else
   Data1.RecordSource = "select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cod_prod in (10050,10001,10003,10005) order by base"
End If
''   Data1.RecordSource = "select * from linmmdd where fecha =#" & Format(Date, "yyyy/mm/dd") & "# and cod_prod =" & 10050 & " and base =" & frm_menu.data_parse.Recordset("base")
Data1.Refresh

Data4.DatabaseName = App.path & "\informes.mdb"

ListView1.ListItems.Clear
frm_pendcmtpol.MousePointer = 11
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      If md.Text = "__/__/____" Then
         Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "# and hora >='" & Data1.Recordset("hora") & "' and cednum =" & Data1.Recordset("ced_socio")
      Else
         Data2.RecordSource = "select * from cabezal_hcdig where fecha >=#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "# and cednum =" & Data1.Recordset("ced_socio")
      End If
      Data2.Refresh
      If Data2.Recordset.RecordCount > 0 Then
      Else
        ListView1.ListItems.Add Xcount, , Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
        ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("hora")
        If IsNull(Data1.Recordset("cod_cli")) = False Then
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("cod_cli")
        Else
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
        End If
        ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nom_cli")
        ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("convenio")
        If IsNull(Data1.Recordset("nom_med_a")) = False Then
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nom_med_a")
        Else
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
        End If
        If IsNull(Data1.Recordset("nro_med_a")) = False Then
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nro_med_a")
        Else
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
        End If
        If IsNull(Data1.Recordset("contact_tel")) = False Then
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("contact_tel")
        Else
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
        End If
        ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nom_prod")
        ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("factura")
        If IsNull(Data1.Recordset("repetir")) = False Then
           If Data1.Recordset("repetir") = "S" Then
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SI"
           Else
              ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO"
           End If
        Else
           ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO"
        End If
        ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("base")
        Consulta_zonas
        ListView1.ListItems.Item(Xcount).ListSubItems.Add , , labzon.Caption
        
        Xcount = Xcount + 1
     End If
     Data1.Recordset.MoveNext
   Loop
   Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/05/2022", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by fecha,hora"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         If md.Text = "__/__/____" Then
            Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "# and hora >='" & Data1.Recordset("hora") & "' and cednum =" & Data1.Recordset("ci") & " and tipo_consd in ('Orientación Telefónica')"
         Else
            Data2.RecordSource = "select * from cabezal_hcdig where fecha >=#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "# and hora >='" & Data1.Recordset("hora") & "' and cednum =" & Data1.Recordset("ci") & " and tipo_consd in ('Orientación Telefónica')"
         End If
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
         Else
         
            ListView1.ListItems.Add Xcount, , Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("hora")
            If IsNull(Data1.Recordset("matric")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("matric")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(Data1.Recordset("nombre")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nombre")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NN"
            End If
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("categ")
            If IsNull(Data1.Recordset("codmedcmt")) = False Then
               Data2.RecordSource = "select * from medicos_esp where id =" & Data1.Recordset("codmedcmt")
               Data2.Refresh
               If Data2.Recordset.RecordCount > 0 Then
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("nom_med")
               Else
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO REGISTRADO"
               End If
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO REGISTRADO"
            End If
            If IsNull(Data1.Recordset("codmedcmt")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("codmedcmt")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(Data1.Recordset("telef")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("telef")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "CMT DESPACHO"
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nrolla")
            If IsNull(Data1.Recordset("repite")) = False Then
               If Data1.Recordset("repite") = 1 Then
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SI"
               Else
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO"
               End If
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO"
            End If
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "19"
            If IsNull(Data1.Recordset("motmov")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("motmov")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "Sin Zona"
            End If
            Xcount = Xcount + 1
         End If
         
         Data1.Recordset.MoveNext
      Loop
   End If
   Label3.Caption = "Total registros:" & Trim(str(Xcount - 1))
Else
   frm_pendcmtpol.MousePointer = 0
   MsgBox "No hay registros de CMT Polic, Pendientes."
   frm_pendcmtpol.MousePointer = 11
      
   Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/05/2022", "yyyy/mm/dd") & "# and pend =" & 4 & " and codmot <>'" & "Z" & "' and (segui_covid not in (1) or segui_covid is null) order by fecha,hora"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         If md.Text = "__/__/____" Then
            Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "# and hora >='" & Data1.Recordset("hora") & "' and cednum =" & Data1.Recordset("ci") & " and tipo_consd in ('Orientación Telefónica')"
         Else
            Data2.RecordSource = "select * from cabezal_hcdig where fecha >=#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "# and hora >='" & Data1.Recordset("hora") & "' and cednum =" & Data1.Recordset("ci") & " and tipo_consd in ('Orientación Telefónica')"
         End If
         Data2.Refresh
         If Data2.Recordset.RecordCount > 0 Then
         Else
            ListView1.ListItems.Add Xcount, , Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("hora")
            If IsNull(Data1.Recordset("matric")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("matric")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(Data1.Recordset("nombre")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nombre")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NN"
            End If
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("categ")
            If IsNull(Data1.Recordset("codmedcmt")) = False Then
               Data2.RecordSource = "select * from medicos_esp where id =" & Data1.Recordset("codmedcmt")
               Data2.Refresh
               If Data2.Recordset.RecordCount > 0 Then
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data2.Recordset("nom_med")
               Else
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO REGISTRADO"
               End If
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO REGISTRADO"
            End If
            If IsNull(Data1.Recordset("codmedcmt")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("codmedcmt")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(Data1.Recordset("telef")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("telef")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "CMT DESPACHO"
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("nrolla")
            If IsNull(Data1.Recordset("repite")) = False Then
               If Data1.Recordset("repite") = 1 Then
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SI"
               Else
                  ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO"
               End If
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "NO"
            End If
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "19"
            If IsNull(Data1.Recordset("motmov")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Data1.Recordset("motmov")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "Sin Zona"
            End If
            Xcount = Xcount + 1
         End If
         Data1.Recordset.MoveNext
      Loop
   Else
      frm_pendcmtpol.MousePointer = 0
      MsgBox "No hay registros de CMT Despacho, Pendientes."
   End If
   Label3.Caption = "Total registros:" & Trim(str(Xcount - 1))


End If
Carga_medicos
frm_pendcmtpol.MousePointer = 0


End Sub
Public Sub Carga_medicos()
Dim Xsqlpromo As String
Dim Xsqlbusca As String
Dim Xrecclii As New ADODB.Recordset
Dim XrecBusca As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select medicos_cmt.cod_med,medicos_cmt.fecha,medicos_cmt.cedula,medicos_cmt.nom_usuario,medicos.med_cod,medicos.med_nombre from medicos_cmt inner join medicos on medicos_cmt.cod_med=medicos.med_cod where medicos_cmt.fecha ='" & Format(Date, "yyyy-mm-dd") & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      Combo1.AddItem Xrecclii("med_nombre")
      Xrecclii.MoveNext
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Buscar_medico()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from medicos where med_nombre ='" & Combo1.Text & "'"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   labcodmed.Caption = Xrecclii("med_cod")
   labnommed.Caption = Xrecclii("med_nombre")
Else
   labcodmed.Caption = ""
   labnommed.Caption = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub
Public Sub Buscar_medicoCmt()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset


If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

If Trim(labcodmed.Caption) <> "" Then

   ConectarBD
   ConbdSapp.Open
             
   Xsqlpromo = "Select * from medicos_esp where cod_sapp =" & Val(labcodmed.Caption)
   With Xrecclii
     .CursorLocation = adUseClient
     .CursorType = adOpenKeyset
     .LockType = adLockOptimistic
     .Open Xsqlpromo, ConbdSapp, , , adCmdText
   End With
   If Xrecclii.RecordCount > 0 Then
      labcodmedcmt.Caption = Xrecclii("id")
   Else
      MsgBox "No se encuentra médico en lista de CMT. Verifique en Agenda", vbCritical
      labcodmedcmt.Caption = ""
   End If

   Xrecclii.Close
   ConbdSapp.Close
Else
   MsgBox "No se encuentra médico en lista de CMT. Verifique en Agenda.", vbCritical
   labcodmedcmt.Caption = ""
End If
End Sub
Public Sub Consulta_zonas()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset


If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from clientes where cl_codigo =" & Data1.Recordset("cod_cli")

With Xrecclii
  .CursorLocation = adUseClient
  .CursorType = adOpenKeyset
  .LockType = adLockOptimistic
  .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   If IsNull(Xrecclii("cl_zona")) = False Then
      labzon.Caption = Xrecclii("cl_zona")
   Else
      labzon.Caption = "Sin zona"
   End If
Else
   labzon.Caption = "Sin Zona"
End If
Xrecclii.Close
ConbdSapp.Close


End Sub

