VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_seleccmt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Agendar CMT"
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frm_seleccmt.frx":0000
   ScaleHeight     =   5385
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_lista 
      Caption         =   "data_lista"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_agenda 
      Caption         =   "data_agenda"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10320
      Picture         =   "frm_seleccmt.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos para la agenda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   10695
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   3600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2640
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data data_hc 
         Caption         =   "data_hc"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2280
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Data data_lla2 
         Caption         =   "data_lla2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2280
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data data_llamod 
         Caption         =   "data_llamod"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2640
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data data_lla 
         Caption         =   "data_lla"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox t_codcons 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9240
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Anular"
         Enabled         =   0   'False
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
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Elimina el paciente de la agenda y vuelve el llamado a pendientes de despacho"
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Agendar"
         Enabled         =   0   'False
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
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2040
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2775
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4895
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Hora"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cédula"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Convenio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Matrícula"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Celular"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Teléfono"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consultar agenda"
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
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Seleccione Médico:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Label lablabase 
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label labcedhc 
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label labced 
      BackColor       =   &H00C00000&
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
      Height          =   255
      Left            =   9120
      TabIndex        =   12
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label labmat 
      BackColor       =   &H00C00000&
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
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label labnom 
      BackColor       =   &H00C00000&
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
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "Socio:"
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
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Seleccione agenda para anotar al paciente"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frm_seleccmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xcount As Integer

ListView1.ListItems.Clear
t_codcons.Text = 0

If Trim(Combo1.Text) <> "" Then
   data_agenda.RecordSource = "select * from t_fechas where nom_med ='" & Combo1.Text & "' and fecha ='" & Format(Date, "dd/mm/yyyy") & "' and base in (98,99) order by nro"
   data_agenda.Refresh
   If data_agenda.Recordset.RecordCount > 0 Then
      data_agenda.Recordset.MoveFirst
      Xcount = 1
      t_codcons.Text = data_agenda.Recordset("cod_cons")
      Do While Not data_agenda.Recordset.EOF
         ListView1.ListItems.Add Xcount, , data_agenda.Recordset("nro")
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("hora")
         If IsNull(data_agenda.Recordset("ced_pac")) = False Then
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("ced_pac")
         Else
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
         End If
         If IsNull(data_agenda.Recordset("nom_pac")) = False Then
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("nom_pac")
         Else
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
         End If
         If IsNull(data_agenda.Recordset("convenio")) = False Then
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("convenio")
         Else
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
         End If
         If IsNull(data_agenda.Recordset("mat_pac")) = False Then
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("mat_pac")
         Else
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
         End If
         If IsNull(data_agenda.Recordset("cel_pac")) = False Then
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("cel_pac")
         Else
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
         End If
         If IsNull(data_agenda.Recordset("tel_pac")) = False Then
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("tel_pac")
         Else
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
         End If
         data_agenda.Recordset.MoveNext
         Xcount = Xcount + 1
      Loop
      Command3.Enabled = True
      Command2.Enabled = True
      data_agenda.Recordset.MovePrevious
      If Format(Time, "HH:mm") > Trim(data_agenda.Recordset("hora")) Then
         MsgBox "El horario de esta consulta ya ha finalizado.", vbInformation
         ListView1.Enabled = False
      Else
         ListView1.Enabled = True
      End If
      
   Else
      Command3.Enabled = False
      Command2.Enabled = False
   End If
Else
   Command3.Enabled = False
   Command2.Enabled = False
End If


End Sub

Private Sub Command2_Click()

Dim Xind, Xcant, Xnro As Long
Dim Xfecdeuda, Xlafechacons As Date
Dim Xloslabos, Xlacedconsulta As String
Dim Xcantlibres, Xellugar, Xtienecmt As Integer
Dim Xlafv As Date
Dim Xelcodigoaut, Xlapersona, Mensajecmtsi As String
Xtienecmt = 0
Xlafv = Date
Mensajecmtsi = ""
Xcantlibres = 0

Xloslabos = ""
Xlafechacons = Date
Dim Xcountt As Long
Dim Xind22 As Integer
Dim Xrecconve As New ADODB.Recordset
Dim Xsqlstr As String

Xind = 0
Xnro = 0
Xcant = 0
Xcountt = 1
            
For Xind = 1 To ListView1.ListItems.count
    ListView1.ListItems(Xind).Selected = True
    If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
       Xcant = Xcant + 1
    End If
Next Xind
Xind = 0
        
If Trim(labmat.Caption) <> "" Then
   Xtienecmt = YatieneCMT()
End If
If Xtienecmt = 1 Then
   Mensajecmtsi = MsgBox("Ya figura registrado en CMT de policlínica " & lablabase.Caption & ". Desea registrar igual bajo su responsabilidad?", vbCritical + vbYesNo, "Despacho")
'   MsgBox "Ya figura registrado en CMT de policlínica " & lablabase.Caption & ". No se puede agendar.", vbCritical, "Despacho"
   If Mensajecmtsi = vbYes Then
      Xtienecmt = 0
   End If
End If

If Xtienecmt <> 1 Then
    If Trim(labced.Caption) <> "" Then
       If Xcant = 1 Then
          For Xind = 1 To ListView1.ListItems.count
              ListView1.ListItems(Xind).Selected = True
              If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
                 Xnro = ListView1.ListItems(Xind).Text
                 Xellugar = Xind
                 Xlacedconsulta = labced.Caption
                 data_lista.RecordSource = "select * from t_fechas where cdate(fecha) =#" & Format(Xlafechacons, "yyyy/mm/dd") & "# and ced_pac ='" & Trim(labced.Caption) & "' and base in (98,99)"
                 data_lista.Refresh
                 If data_lista.Recordset.RecordCount > 0 Then
                    MsgBox "Estaba anotado para " & data_lista.Recordset("nom_med") & ". Quedará el lugar disponible.", vbInformation
                    data_lista.Recordset.Edit
                    data_lista.Recordset("mat_pac") = Null
                    data_lista.Recordset("nom_pac") = Null
                    data_lista.Recordset("tipoconsulta") = Null
                    data_lista.Recordset("tipoconsultan") = Null
                    data_lista.Recordset("ced_pac") = Null
                    data_lista.Recordset("convenio") = Null
                    data_lista.Recordset("cel_pac") = Null
                    data_lista.Recordset("tel_pac") = Null
                    data_lista.Recordset("fec_anota") = Null
                    data_lista.Recordset("hora_anota") = Null
                    data_lista.Recordset("usua_anota") = Null
                    data_lista.Recordset("usua_web") = Null
                    data_lista.Recordset("obs") = Null
                    data_lista.Recordset.Update
                 End If
                 data_lista.RecordSource = "select * from t_fechas where cdate(fecha) >#" & Format(Xlafechacons, "yyyy/mm/dd") & "# and ced_pac ='" & Trim(labced.Caption) & "' and base in (98,99) and nom_med ='" & Combo1.Text & "'"
                 data_lista.Refresh
                 If data_lista.Recordset.RecordCount > 0 Then
                    MsgBox "Ya se encuentra anotado para una consulta telefónica. VERIFIQUE!", vbExclamation
                 Else
                    data_lista.RecordSource = "select * from t_fechas where fecha ='" & Format(Date, "dd/mm/yyyy") & "' and cod_cons =" & t_codcons.Text & " and nro =" & Xnro
                    data_lista.Refresh
                    If data_lista.Recordset.RecordCount > 0 Then
                       If IsNull(data_lista.Recordset("nom_pac")) = False And IsNull(data_lista.Recordset("ced_pac")) = False Then
                          MsgBox "Ya existe un paciente anotado, verifique!!", vbCritical
                       Else
                          data_lista.Recordset.Edit
                          If Trim(labmat.Caption) <> "" Then
                             data_lista.Recordset("mat_pac") = Val(labmat.Caption)
                          Else
                             data_lista.Recordset("mat_pac") = 0
                          End If
                          If Trim(labnom.Caption) <> "" Then
                             data_lista.Recordset("nom_pac") = labnom.Caption
                          End If
                          data_lista.Recordset("tipoconsulta") = "Telefónica"
                          data_lista.Recordset("tipoconsultan") = 1
                          If Trim(labced.Caption) <> "" Then
                             data_lista.Recordset("ced_pac") = labced.Caption
                          End If
                          If frm_largador.txt_cat.Text <> "" Then
                             data_lista.Recordset("convenio") = frm_largador.txt_cat.Text
                          End If
                          If frm_largador.txt_tel.Text <> "" Then
                             data_lista.Recordset("cel_pac") = Mid(Trim(frm_largador.txt_tel.Text), 1, 45)
                          End If
                          If frm_largador.txt_tel.Text <> "" Then
                             data_lista.Recordset("tel_pac") = Mid(Trim(frm_largador.txt_tel.Text), 1, 45)
                          End If
                          data_lista.Recordset("fec_anota") = Format(Date, "dd/mm/yyyy")
                          data_lista.Recordset("hora_anota") = Format(Time, "HH:mm")
                          data_lista.Recordset("usua_anota") = WElusuario
                          data_lista.Recordset("usua_web") = "SAPP"
                          If Trim(frm_largador.txt_mot.Text) <> "" Then
                             data_lista.Recordset("obs") = Mid(frm_largador.txt_mot.Text, 1, 190)
                          End If
                          data_lista.Recordset.Update
                          data_lla.RecordSource = "Select * from llamado where nrolla =" & frm_largador.txt_nro.Text
                          data_lla.Refresh
                          If data_lla.Recordset.RecordCount > 0 Then
                             If data_lla.Recordset("pend") = 4 Or data_lla.Recordset("pend") = 1 Or data_lla.Recordset("pend") = 2 Then
                                MsgBox "El llamado ya estaba pasado a CMT.", vbInformation
                                If IsNull(data_lla.Recordset("codmedcmt")) = False Then
                                   If data_lla.Recordset("codmedcmt") <> data_lista.Recordset("cod_med") Then
                                      data_lla.Recordset.Edit
                                      data_lla.Recordset("codmedcmt") = data_lista.Recordset("cod_med")
                                      data_lla.Recordset("editando") = 1
                                      data_lla.Recordset.Update
                                   End If
                                Else
                                   data_lla.Recordset.Edit
                                   data_lla.Recordset("codmedcmt") = data_lista.Recordset("cod_med")
                                   data_lla.Recordset("editando") = 1
                                   data_lla.Recordset.Update
                                End If
                             Else
                                data_lla.Recordset.Edit
                                data_lla.Recordset("pend") = 4
                                data_lla.Recordset("codmedcmt") = data_lista.Recordset("cod_med")
                                data_lla.Recordset("editando") = 1
                                data_lla.Recordset.Update
                                data_llamod.RecordSource = "Select * from resplla where nro =" & frm_largador.txt_nro.Text
                                data_llamod.Refresh
                                If data_llamod.Recordset.RecordCount > 0 Then
                                   data_llamod.Recordset.Edit
                                   data_llamod.Recordset("hzona") = Format(Time, "HH:mm")
                                   data_llamod.Recordset("pend") = 1
                                   data_llamod.Recordset.Update
                                Else
                                   data_llamod.Recordset.AddNew
                                   data_llamod.Recordset("nro") = txt_nro.Text
                                   data_llamod.Recordset("fecha") = frm_largador.mfecha.Text
                                   data_llamod.Recordset("hzona") = Format(Time, "HH:mm")
                                   data_llamod.Recordset("pend") = 1
                                   data_llamod.Recordset.Update
                                End If
                                frm_largador.labcmt.Visible = True
                                frm_largador.labcmt.Caption = "PASADO A CMT HORA:" & Format(Time, "HH:mm")
                                frm_largador.b_cmt.Enabled = False
                             End If
                          Else
                             MsgBox "Error al buscar el llamado"
                          End If
                          
                       End If
                    End If
                 End If
              End If
          Next Xind
       
          data_agenda.RecordSource = "select * from t_fechas where nom_med ='" & Combo1.Text & "' and fecha ='" & Format(Date, "dd/mm/yyyy") & "' and base in (98,99) order by nro"
          data_agenda.Refresh
          ListView1.ListItems.Clear
          If data_agenda.Recordset.RecordCount > 0 Then
             data_agenda.Recordset.MoveFirst
             Xcount = 1
             t_codcons.Text = data_agenda.Recordset("cod_cons")
             Do While Not data_agenda.Recordset.EOF
                ListView1.ListItems.Add Xcount, , data_agenda.Recordset("nro")
                ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("hora")
                If IsNull(data_agenda.Recordset("ced_pac")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("ced_pac")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
                End If
                If IsNull(data_agenda.Recordset("nom_pac")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("nom_pac")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
                End If
                If IsNull(data_agenda.Recordset("convenio")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("convenio")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
                End If
                If IsNull(data_agenda.Recordset("mat_pac")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("mat_pac")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
                End If
                If IsNull(data_agenda.Recordset("cel_pac")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("cel_pac")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
                End If
                If IsNull(data_agenda.Recordset("tel_pac")) = False Then
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("tel_pac")
                Else
                   ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
                End If
                data_agenda.Recordset.MoveNext
                Xcount = Xcount + 1
             Loop
          End If
          ListView1.ListItems.Item(Xellugar).Selected = True
          ListView1.ListItems.Item(Xellugar).EnsureVisible
          ListView1.SetFocus
          Command3.Enabled = False
          Command2.Enabled = False
    '      Unload Me
       Else
          MsgBox "Debe seleccionar solo un registro.", vbCritical
       End If
    Else
       MsgBox "Debe ingresar cédula del paciente para agendar.", vbCritical
    End If
Else
    MsgBox "No se agendó.", vbInformation
End If

'desde acá cerrar en despacho


End Sub

Private Sub Command3_Click()
Dim Xind, Xcant, Xnro As Long
Dim Xborralaconsulta As String
Dim Xcantlibres As Integer
Xcantlibres = 0

Xborralaconsulta = MsgBox("Desea anular la CMT agendada ?", vbInformation + vbYesNo)
If Xborralaconsulta = vbYes Then
    Xind = 0
    Xnro = 0
    Xcant = 0
    Dim Xcountt As Long
    Dim Xdeudasiono As Integer
    Xdeudasiono = 0
    Xcountt = 1

    For Xind = 1 To ListView1.ListItems.count
        ListView1.ListItems(Xind).Selected = True
        If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
           Xcant = Xcant + 1
        End If
    Next Xind
    Xind = 0
    
    If Xcant = 1 Then
       For Xind = 1 To ListView1.ListItems.count
           ListView1.ListItems(Xind).Selected = True
           If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
              Xnro = ListView1.ListItems(Xind).Text
              data_lista.RecordSource = "select * from t_fechas where fecha ='" & Format(Date, "dd/mm/yyyy") & "' and cod_cons =" & t_codcons.Text & " and nro =" & Xnro
              data_lista.Refresh
              If data_lista.Recordset.RecordCount > 0 Then
                 data_lista.Recordset.Edit
                 data_lista.Recordset("mat_pac") = Null
                 data_lista.Recordset("nom_pac") = Null
                 data_lista.Recordset("ced_pac") = Null
                 data_lista.Recordset("convenio") = Null
                 data_lista.Recordset("cel_pac") = Null
                 data_lista.Recordset("tel_pac") = Null
                 data_lista.Recordset("fec_nac") = Null
                 data_lista.Recordset("hcsiono") = Null
                 data_lista.Recordset("tipo_cons") = Null
                 data_lista.Recordset("tipo_consd") = Null
                 data_lista.Recordset("fec_anota") = Null
                 data_lista.Recordset("hora_anota") = Null
                 data_lista.Recordset("usua_anota") = Null
                 data_lista.Recordset("edad") = Null
                 data_lista.Recordset("usua_web") = Null
                 data_lista.Recordset("tipoconsulta") = Null
                 data_lista.Recordset("tipoconsultan") = Null
                 data_lista.Recordset("hora_realizacmt") = Null
                 data_lista.Recordset.Update
                 
                 data_lla2.RecordSource = "Select * from llamado where nrolla =" & frm_largador.txt_nro.Text
                 data_lla2.Refresh
                 If data_lla2.Recordset.RecordCount > 0 Then
                    If data_lla2.Recordset("pend") = 4 Then
                       data_lla2.Recordset.Edit
                       data_lla2.Recordset("pend") = 0
                       If IsNull(data_lla2.Recordset("obsmot")) = False Then
                          data_lla2.Recordset("obsmot") = data_lla2.Recordset("obsmot") & " " & "Se cancela la agenda a CMT por " & WElusuario
                       Else
                          data_lla2.Recordset("obsmot") = "Se cancela la agenda a CMT por " & WElusuario
                       End If
                       data_lla2.Recordset("hora_anterior") = data_lla2.Recordset("hora")
                       data_lla2.Recordset("hora") = Format(Time, "HH:mm")
                       data_lla2.Recordset("activo") = Format(Time, "HH:mm:ss")
                       data_lla2.Recordset("cmt_enproceso") = 2
                       data_lla2.Recordset("codmedcmt") = Null
                       data_lla2.Recordset.Update
                    End If
                 End If
                 
                 data_lla2.RecordSource = "Select * from resplla where nro =" & frm_largador.txt_nro.Text
                 data_lla2.Refresh
                 If data_lla2.Recordset.RecordCount > 0 Then
                    If IsNull(data_lla2.Recordset("hzona")) = False Then
                       data_lla2.Recordset.Edit
                       data_lla2.Recordset("hzona") = Null
                       data_lla2.Recordset.Update
                    End If
                    If IsNull(data_lla2.Recordset("mm")) = False Then
                       data_lla2.Recordset.Edit
                       data_lla2.Recordset("mm") = Null
                       data_lla2.Recordset.Update
                    End If
                    If IsNull(data_lla2.Recordset("hsald")) = False Then
                       data_lla2.Recordset.Edit
                       data_lla2.Recordset("hsald") = Null
                       data_lla2.Recordset.Update
                    End If
                    If IsNull(data_lla2.Recordset("totend")) = False Then
                       data_lla2.Recordset.Edit
                       data_lla2.Recordset("totend") = Null
                       data_lla2.Recordset.Update
                    End If
                 End If
                 MsgBox "El llamado pasó a PENDIENTES!", vbInformation
              
              End If
           End If
       Next Xind
    
      data_agenda.RecordSource = "select * from t_fechas where nom_med ='" & Combo1.Text & "' and fecha ='" & Format(Date, "dd/mm/yyyy") & "' and base in (98,99) order by nro"
      data_agenda.Refresh
      ListView1.ListItems.Clear
      If data_agenda.Recordset.RecordCount > 0 Then
         data_agenda.Recordset.MoveFirst
         Xcount = 1
         t_codcons.Text = data_agenda.Recordset("cod_cons")
         Do While Not data_agenda.Recordset.EOF
            ListView1.ListItems.Add Xcount, , data_agenda.Recordset("nro")
            ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("hora")
            If IsNull(data_agenda.Recordset("ced_pac")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("ced_pac")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(data_agenda.Recordset("nom_pac")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("nom_pac")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(data_agenda.Recordset("convenio")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("convenio")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(data_agenda.Recordset("mat_pac")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("mat_pac")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(data_agenda.Recordset("cel_pac")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("cel_pac")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            If IsNull(data_agenda.Recordset("tel_pac")) = False Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_agenda.Recordset("tel_pac")
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            End If
            data_agenda.Recordset.MoveNext
            Xcount = Xcount + 1
         Loop
      End If
        
    Else
       MsgBox "Debe seleccionar un solo registro"
    End If
End If

End Sub

Private Sub Command4_Click()
Unload Me

End Sub

Private Sub Form_Load()
If Trim(frm_largador.txt_nomb.Text) <> "" Then
   labnom.Caption = frm_largador.txt_nomb.Text
Else
   labnom.Caption = "NN"
End If
labmat.Caption = frm_largador.txt_mat.Text
labced.Caption = frm_largador.txt_ced.Text & frm_largador.t_codced.Text
data_lista.Connect = "odbc;dsn=" & Xconexrmt & ";"
labcedhc.Caption = frm_largador.txt_ced.Text

data_lla2.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_hc.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_lla.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_llamod.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_agenda.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_agenda.RecordSource = "select * from t_cabfechas where fecha ='" & Format(Date, "dd/mm/yyyy") & "' and base =" & 98
data_agenda.Refresh
If data_agenda.Recordset.RecordCount > 0 Then
   data_agenda.Recordset.MoveFirst
   Do While Not data_agenda.Recordset.EOF
      Combo1.AddItem data_agenda.Recordset("nom_med")
      data_agenda.Recordset.MoveNext
   Loop
End If

End Sub

Public Function YatieneCMT() As Integer

Dim XsqlpromoF As String
Dim XreccliiAvisoF As New ADODB.Recordset

ConectarAvisoF
ConbdSappAvisoF.Open

XsqlpromoF = "Select * from linmmdd where cod_cli =" & Val(labmat.Caption) & " and fecha ='" & Format(Date, "yyyy-mm-dd") & "' and cod_prod in (10050)"
With XreccliiAvisoF
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open XsqlpromoF, ConbdSappAvisoF, , , adCmdText
End With
If XreccliiAvisoF.RecordCount > 0 Then
   lablabase.Caption = XreccliiAvisoF("base")
   Data2.RecordSource = "select * from cabezal_hcdig where fecha =#" & Format(XreccliiAvisoF("fecha"), "yyyy/mm/dd") & "# and cednum =" & Val(labced.Caption)
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      lablabase.Caption = ""
      YatieneCMT = 0
   Else
      YatieneCMT = 1
   End If
Else
   lablabase.Caption = ""
   YatieneCMT = 0
End If

XreccliiAvisoF.Close
ConbdSappAvisoF.Close


End Function

