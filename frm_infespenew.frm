VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_infespenew 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de reservas especialistas"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   Icon            =   "frm_infespenew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboes 
      BackColor       =   &H00FF0000&
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
      Height          =   315
      ItemData        =   "frm_infespenew.frx":058A
      Left            =   2040
      List            =   "frm_infespenew.frx":05DF
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Data data_buscasapp 
      Caption         =   "data_buscasapp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2400
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_buscar 
      Caption         =   "data_buscar"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      MouseIcon       =   "frm_infespenew.frx":074A
      MousePointer    =   99  'Custom
      Picture         =   "frm_infespenew.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      MouseIcon       =   "frm_infespenew.frx":125E
      MousePointer    =   99  'Custom
      Picture         =   "frm_infespenew.frx":17E8
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Datos para el informe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.Data data_facturas 
         Caption         =   "data_facturas"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1560
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   720
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF0000&
         Caption         =   "Emitir detalle de pacientes anotados"
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
         TabIndex        =   14
         Top             =   3360
         Width           =   5055
      End
      Begin VB.TextBox t_b 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
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
         Left            =   1800
         TabIndex        =   12
         Text            =   "99"
         Top             =   2280
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF0000&
         Caption         =   "Resumen"
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
         Left            =   3600
         TabIndex        =   8
         Top             =   2880
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Detalle"
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
         Left            =   1920
         TabIndex        =   7
         Top             =   2880
         Width           =   1575
      End
      Begin VB.ComboBox cboop 
         BackColor       =   &H00FF0000&
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
         Height          =   315
         ItemData        =   "frm_infespenew.frx":1D72
         Left            =   1800
         List            =   "frm_infespenew.frx":1D8B
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   3375
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16711680
         ForeColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16711680
         ForeColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         Caption         =   "Especialidad:"
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
         TabIndex        =   15
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "(99=TODAS)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Número de Base:"
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
         TabIndex        =   11
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Formato:"
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
         TabIndex        =   6
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Opción de listado:"
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
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Rango de fechas:"
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
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   960
      Top             =   3240
      Width           =   2055
   End
End
Attribute VB_Name = "frm_infespenew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboop_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_b.SetFocus
End If

End Sub

Private Sub Command1_Click()
Dim Xmatesp As Long
Xmatesp = 0
'infespec.mdb
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\infespec.mdb")

MiBaseact.Execute "Delete * from lista"
MiBaseact.Execute "Delete * from novino"
MiBaseact.Execute "delete * from fechascons"

If md.Text <> "__/__/____" And mh.Text <> "__/__/____" Then
  If Not cboop.ListIndex = 5 Then 'ListIndex = 5 informe agregado sobre consultorios, mientras no sea la 5ta opcion, sigue todo igual'
   If cboop.ListIndex = 2 Or cboop.ListIndex = 3 Then
      data_inf.RecordSource = "novino"
      data_inf.Refresh
   Else
      data_inf.RecordSource = "fechascons"
      data_inf.Refresh
   End If
   If cboop.ListIndex = 0 Then
      If t_b.Text = 99 Then
         If Check1.Value = 1 Then
            data_inf.RecordSource = "lista"
            data_inf.Refresh
            If cboes.ListIndex <= 0 Then
               data_buscar.RecordSource = "Select * from t_cabfechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_med =" & frm_especialistas.Data1.Recordset("cod_med") & " and cancela not in ('SI') order by base,fecha"
            Else
               data_buscar.RecordSource = "Select * from t_cabfechas where especial ='" & cboes.Text & "' and cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_med =" & frm_especialistas.Data1.Recordset("cod_med") & " and cancela not in ('SI') order by base,fecha"
            End If
'            data_buscar.RecordSource = "select * from t_fechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_med =" & frm_especialistas.Data1.Recordset("cod_med") & " and cancela not in ('SI') and fecha is not null order by base,fecha,nro"
            data_buscar.Refresh
         Else
            If cboes.ListIndex <= 0 Then
               data_buscar.RecordSource = "Select * from t_cabfechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_med =" & frm_especialistas.Data1.Recordset("cod_med") & " and cancela not in ('SI') order by base,fecha"
            Else
               data_buscar.RecordSource = "Select * from t_cabfechas where especial ='" & cboes.Text & "' and cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_med =" & frm_especialistas.Data1.Recordset("cod_med") & " and cancela not in ('SI') order by base,fecha"
            End If
            data_buscar.Refresh
         End If
      Else
         If Check1.Value = 1 Then
            data_inf.RecordSource = "lista"
            data_inf.Refresh
            If cboes.ListIndex <= 0 Then
               data_buscar.RecordSource = "Select * from t_cabfechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_med =" & frm_especialistas.Data1.Recordset("cod_med") & " and base =" & t_b.Text & " and cancela not in ('SI') order by fecha"
            Else
               data_buscar.RecordSource = "Select * from t_cabfechas where especial ='" & cboes.Text & "' and cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_med =" & frm_especialistas.Data1.Recordset("cod_med") & " and base =" & t_b.Text & " and cancela not in ('SI') order by fecha"
            End If
'            data_buscar.RecordSource = "select * from t_fechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_med =" & frm_especialistas.Data1.Recordset("cod_med") & " and base =" & t_b.Text & " and cancela not in ('SI') order by base,fecha,nro"
            data_buscar.Refresh
         Else
            If cboes.ListIndex <= 0 Then
               data_buscar.RecordSource = "Select * from t_cabfechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_med =" & frm_especialistas.Data1.Recordset("cod_med") & " and base =" & t_b.Text & " and cancela not in ('SI') order by fecha"
            Else
               data_buscar.RecordSource = "Select * from t_cabfechas where especial ='" & cboes.Text & "' and cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_med =" & frm_especialistas.Data1.Recordset("cod_med") & " and base =" & t_b.Text & " and cancela not in ('SI') order by fecha"
            End If
            data_buscar.Refresh
         End If
      End If
   Else
      If cboop.ListIndex = 1 Then
         If t_b.Text = 99 Then
            If cboes.ListIndex <= 0 Then
               data_buscar.RecordSource = "Select * from t_cabfechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cancela not in ('SI') order by base,fecha"
            Else
               data_buscar.RecordSource = "Select * from t_cabfechas where especial ='" & cboes.Text & "' and cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cancela not in ('SI') order by base,fecha"
            End If
            data_buscar.Refresh
         Else
            If cboes.ListIndex <= 0 Then
               data_buscar.RecordSource = "Select * from t_cabfechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and base =" & t_b.Text & " and cancela not in ('SI') order by fecha"
            Else
               data_buscar.RecordSource = "Select * from t_cabfechas where especial ='" & cboes.Text & "' and cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and base =" & t_b.Text & " and cancela not in ('SI') order by fecha"
            End If
            data_buscar.Refresh
         End If
      Else
         If cboop.ListIndex = 2 Then
            data_buscar.RecordSource = "Select * from borrados where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and obs not in ('CANCELA SIN AVISO')"
'            data_buscar.RecordSource = "Select * from t_fechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and nom_pac is not null order by fecha"
            data_buscar.Refresh
         Else
            If cboop.ListIndex = 3 Then
               data_buscar.RecordSource = "Select * from borrados where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and obs in ('CANCELA SIN AVISO') order by local,fecha"
               data_buscar.Refresh
            Else
               If cboop.ListIndex = 4 Then
                  If cboes.ListIndex <= 0 Then
                     data_buscar.RecordSource = "Select * from t_cabfechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cancela ='" & "SI" & "'  order by base,fecha"
                  Else
                     data_buscar.RecordSource = "Select * from t_cabfechas where especial ='" & cboes.Text & "' and cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cancela ='" & "SI" & "'  order by base,fecha"
                  End If
                  data_buscar.Refresh
               Else
                  If cboop.ListIndex = 6 Then
                     data_inf.RecordSource = "lista"
                     data_inf.Refresh
                     data_buscar.RecordSource = "SELECT * FROM t_fechas where usua_anota='" & "WhatsApp" & "' and fecha_cons >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha_cons <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by fecha"
                     data_buscar.Refresh
                  Else
                     If cboes.ListIndex <= 0 Then
                        data_buscar.RecordSource = "Select * from t_cabfechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by base,fecha"
                     Else
                        data_buscar.RecordSource = "Select * from t_cabfechas where especial ='" & cboes.Text & "' and cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by base,fecha"
                     End If
                     data_buscar.Refresh
                  End If
               End If
            End If
         End If
      End If
   End If
   If data_buscar.Recordset.RecordCount > 0 Then
      data_buscar.Recordset.MoveLast
      data_buscar.Recordset.MoveFirst
      Do While Not data_buscar.Recordset.EOF
            If cboop.ListIndex = 2 Or cboop.ListIndex = 3 Then
               data_inf.Recordset.AddNew
               data_inf.Recordset("nom") = data_buscar.Recordset("cedula")
               data_inf.Recordset("fecha") = Format(data_buscar.Recordset("fecha"), "dd/mm/yyyy")
               data_inf.Recordset("nom_med") = data_buscar.Recordset("medico")
               data_inf.Recordset("espec") = Mid(data_buscar.Recordset("obs"), 1, 40)
               data_inf.Recordset("base") = data_buscar.Recordset("local")
               data_inf.Recordset("telef") = data_buscar.Recordset("hora_cons")
               data_inf.Recordset("fecons") = data_buscar.Recordset("fecha_cons")
               data_inf.Recordset("hora_borra") = Format(Time, "HH:mm")
               data_inf.Recordset.Update
               data_buscar.Recordset.MoveNext
            Else
               If cboop.ListIndex = 0 And Check1.Value = 1 Then
                  Data1.RecordSource = "Select * from t_fechas where cod_cons =" & data_buscar.Recordset("cod_cons")
                  Data1.Refresh
                  If Data1.Recordset.RecordCount > 0 Then
                     Data1.Recordset.MoveFirst
                     Do While Not Data1.Recordset.EOF
                        If IsNull(Data1.Recordset("nom_pac")) = False Then
                           data_inf.Recordset.AddNew
                           data_inf.Recordset("fecha") = Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
                           data_inf.Recordset("medico") = Data1.Recordset("nom_med")
                           data_inf.Recordset("espec") = Data1.Recordset("especial")
                           data_inf.Recordset("base") = Data1.Recordset("base")
                           data_inf.Recordset("nro") = Data1.Recordset("nro")
                           data_inf.Recordset("hora") = Data1.Recordset("hora")
                           If IsNull(Data1.Recordset("nom_pac")) = False Then
                              data_inf.Recordset("nom_pac") = Data1.Recordset("nom_pac")
                           End If
                           If IsNull(Data1.Recordset("ced_pac")) = False Then
                              data_inf.Recordset("cedula") = Data1.Recordset("ced_pac")
                           End If
                           If IsNull(Data1.Recordset("mat_pac")) = False Then
                              data_inf.Recordset("mat") = Data1.Recordset("mat_pac")
                           End If
                           If IsNull(Data1.Recordset("convenio")) = False Then
                              data_inf.Recordset("convenio") = Data1.Recordset("convenio")
                           End If
                           If IsNull(Data1.Recordset("cel_pac")) = False Then
                              data_inf.Recordset("celular") = Data1.Recordset("cel_pac")
                           End If
                           If IsNull(Data1.Recordset("tel_pac")) = False Then
                              data_inf.Recordset("telef") = Data1.Recordset("tel_pac")
                           End If
                           If IsNull(Data1.Recordset("tipo_consd")) = False Then
                              data_inf.Recordset("tipocons") = Data1.Recordset("tipo_consd")
                           End If
                           If IsNull(Data1.Recordset("hcsiono")) = False Then
                              If Data1.Recordset("hcsiono") = 0 Then
                                 data_inf.Recordset("hc") = "SI"
                              Else
                                 data_inf.Recordset("hc") = "NO"
                              End If
                           End If
                           If IsNull(Data1.Recordset("edad")) = False Then
                              data_inf.Recordset("edad") = Data1.Recordset("edad")
                           End If
                           If IsNull(Data1.Recordset("fec_nac")) = False Then
                              data_inf.Recordset("fnac") = Data1.Recordset("fec_nac")
                           End If
                           If IsNull(Data1.Recordset("cod_cons")) = False Then
                              data_inf.Recordset("codcons") = Data1.Recordset("cod_cons")
                           End If
                           data_inf.Recordset("via") = Mid(Data1.Recordset("usua_anota"), 1, 15)
                           data_inf.Recordset.Update
                        End If
                        Data1.Recordset.MoveNext
                     Loop
                  End If
                  data_buscar.Recordset.MoveNext
               Else
                  If cboop.ListIndex = 6 Then
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("fecha") = data_buscar.Recordset("fecha_cons")
                     data_inf.Recordset("hora") = data_buscar.Recordset("hora")
                     data_inf.Recordset("medico") = data_buscar.Recordset("nom_med")
                     data_inf.Recordset("espec") = data_buscar.Recordset("especial")
                     data_inf.Recordset("base") = data_buscar.Recordset("base")
                     data_inf.Recordset("nom_pac") = data_buscar.Recordset("nom_pac")
                     data_inf.Recordset("cedula") = data_buscar.Recordset("ced_pac")
                     data_inf.Recordset("mat") = data_buscar.Recordset("mat_pac")
                     data_inf.Recordset("convenio") = data_buscar.Recordset("convenio")
                     data_inf.Recordset("celular") = data_buscar.Recordset("cel_pac")
                     data_inf.Recordset("via") = Mid(data_buscar.Recordset("usua_anota"), 1, 15)
                     If IsNull(data_buscar.Recordset("mat_pac")) = False Then
                        data_facturas.RecordSource = "select * from linmmdd where fecha =#" & Format(data_buscar.Recordset("fecha_cons"), "yyyy/mm/dd") & "# and cod_cli =" & data_buscar.Recordset("mat_pac") & " and cod_prod in (2,3)"
                        data_facturas.Refresh
                        If data_facturas.Recordset.RecordCount > 0 Then
                           data_inf.Recordset("edad") = "ASISTIÓ"
                        Else
                           data_inf.Recordset("edad") = "NO ASISTIÓ"
                        End If
                     End If
                     data_inf.Recordset.Update
                     data_buscar.Recordset.MoveNext
                  Else
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("fecha") = Format(data_buscar.Recordset("fecha"), "dd/mm/yyyy")
                     data_inf.Recordset("hora") = data_buscar.Recordset("hora")
                     data_inf.Recordset("cod_med") = data_buscar.Recordset("cod_med")
                     data_inf.Recordset("descfec") = data_buscar.Recordset("des_fecha")
                     data_inf.Recordset("base") = data_buscar.Recordset("base")
                     data_inf.Recordset("especi") = data_buscar.Recordset("especial")
                     data_inf.Recordset("nom_med") = data_buscar.Recordset("nom_med")
                     If IsNull(data_buscar.Recordset("hora_fin")) = False Then
                        data_inf.Recordset("hora_fin") = data_buscar.Recordset("hora_fin")
                     End If
                     If IsNull(data_buscar.Recordset("cant_pac")) = False Then
                        data_inf.Recordset("cant_pac") = data_buscar.Recordset("cant_pac")
                     End If
                     If IsNull(data_buscar.Recordset("cancela")) = False Then
                        If data_buscar.Recordset("cancela") = "SI" Then
                           data_inf.Recordset("obs") = data_buscar.Recordset("motivo")
                        End If
                     End If
                     data_inf.Recordset.Update
                     data_buscar.Recordset.MoveNext
                  End If
               End If
            End If
         'End If
      Loop
      MsgBox "Proceso terminado"
      If cboop.ListIndex = 2 Or cboop.ListIndex = 3 Then
         data_inf.RecordSource = "select * from novino"
         data_inf.Refresh
      Else
         data_inf.RecordSource = "select * from fechascons"
         data_inf.Refresh
      End If
      If cboop.ListIndex = 0 Then
         If Check1.Value = 1 Then
            data_inf.RecordSource = "Select * from lista"
            data_inf.Refresh
            cr1.ReportFileName = App.path & "\inflistaesptot.rpt"
            cr1.Action = 1
         Else
            cr1.ReportFileName = App.path & "\inffechasnew.rpt"
            cr1.ReportTitle = "Informe de Especialistas por fecha: " & md.Text & " HASTA:" & mh.Text
            cr1.Action = 1
         End If
      Else
         If cboop.ListIndex = 2 Then
            cr1.ReportFileName = App.path & "\inffechasnew3.rpt"
            cr1.ReportTitle = "Informe de Pacientes que se anotaron y no concurrieron fecha: " & md.Text & " HASTA:" & mh.Text
            cr1.Action = 1
         Else
            If cboop.ListIndex = 3 Then
               cr1.ReportFileName = App.path & "\inffechasnew4.rpt"
               cr1.ReportTitle = "Informe de Pacientes que se BORRARON de consulta vía WEB fecha: " & md.Text & " HASTA:" & mh.Text
               cr1.Action = 1
            Else
               If cboop.ListIndex = 4 Then
                  cr1.ReportFileName = App.path & "\inffechasnew5.rpt"
                  cr1.ReportTitle = "Informe de Consultas Canceladas por fecha: " & md.Text & " HASTA:" & mh.Text
                  cr1.Action = 1
               Else
                  If cboop.ListIndex = 6 Then
                     data_inf.RecordSource = "Select * from lista"
                     data_inf.Refresh
                     cr1.ReportFileName = App.path & "\inflistaespchat.rpt"
                     cr1.Action = 1
                  Else
                     cr1.ReportFileName = App.path & "\inffechasnew2.rpt"
                     cr1.ReportTitle = "Informe de Especialistas por fecha: " & md.Text & " HASTA:" & mh.Text
                     cr1.Action = 1
                  End If
               End If
            End If
         End If
      End If
   Else
      MsgBox "No existen registros", vbInformation
   End If
    Else 'ListIndex = 5 , informe de reservas de consultorios
        base = t_b.Text
        If base = "99" Then
            baseId = 0
        Else
            baseId = nroBase_idBase.Item(Val(base))
        End If
        
        If IsDate(md.Text) Then
            desde = Format(md.Text, "yyyy-MM-dd")
            hasta = Format(mh.Text, "yyyy-MM-dd")
            pathArchivo = App.path & "/reservas consultorios " & desde & "_" & base & ".xls"
            
            FileModule.DownloadFile GetParameters.getValor(1) & "/bases/" & baseId & "/consultorios/0/disponibilidades/informe/" & desde & "?hasta=" & hasta, pathArchivo
            
            FileModule.LoadUserFile pathArchivo
        Else
            MsgBox "Ingrese una fecha valida en el formato: dd/mm/yyyy"
        End If
  End If
End If
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_buscar.Connect = "ODBC;DSN=" & Xconexrmt & ";"
data_inf.DatabaseName = App.path & "\infespec.mdb"
data_buscasapp.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.Connect = "ODBC;DSN=" & Xconexrmt & ";"
data_facturas.Connect = "ODBC;DSN=" & Xconexrmt & ";"


End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboop.SetFocus
End If

End Sub

Private Sub t_b_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Option1.SetFocus
End If

End Sub
