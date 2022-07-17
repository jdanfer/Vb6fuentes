VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_infconssoc 
   BackColor       =   &H00FF8080&
   Caption         =   "Consultas por socios"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6345
   Icon            =   "frm_infconssoc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   6345
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_lin 
      Caption         =   "data_lin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   2655
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   120
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   5640
      Picture         =   "frm_infconssoc.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4560
      Picture         =   "frm_infconssoc.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Informes de Consultas por Socio"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.Data Data1 
         Caption         =   "Data1"
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
         Top             =   720
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox checkenf 
         BackColor       =   &H00FF0000&
         Caption         =   "Todos los actos de enfermería"
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
         TabIndex        =   16
         Top             =   2520
         Width           =   2535
      End
      Begin VB.CheckBox checkped 
         BackColor       =   &H00FF0000&
         Caption         =   "Todas las consultas pediatría"
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
         Left            =   2880
         TabIndex        =   15
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CheckBox checkesp 
         BackColor       =   &H00FF0000&
         Caption         =   "Todas las consultas especialistas"
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
         TabIndex        =   14
         Top             =   2040
         Width           =   2535
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Data data_temp 
         Caption         =   "data_temp"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   840
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FF0000&
         Caption         =   "Multiconsultas domicilio"
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
         Left            =   2880
         TabIndex        =   12
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox t_can 
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   960
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
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   3000
         Width           =   2175
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
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FF0000&
         Caption         =   "Todas las consultas en policlínica"
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
         Left            =   2880
         TabIndex        =   5
         Top             =   1560
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF0000&
         Caption         =   "Todas las consultas en domicilio"
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
         Top             =   1560
         Width           =   2535
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Consultas > a:"
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
         TabIndex        =   10
         Top             =   960
         Width           =   2175
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
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4080
      Picture         =   "frm_infconssoc.frx":109E
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1695
   End
End
Attribute VB_Name = "frm_infconssoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xlamat As Long
Dim Xca As Integer
Xca = 0

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"
MiBaseact.Execute "Delete * from inflla"

'If Check3.Value = 1 Then
'   data_inf.RecordSource = "inflla"
'   data_inf.Refresh
'Else
   data_inf.RecordSource = "infvtas"
   data_inf.Refresh
'End If

If md.Text = "__/__/____" And mh.Text = "__/__/____" Then
   MsgBox "Ingrese rango de fechas"
Else
   frm_infconssoc.MousePointer = 11
   If Check1.Value = 1 Then
      data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_prod in (10002,10004,10006) order by cod_cli"
      data_lin.Refresh
   Else
      If Check2.Value = 1 Then
         data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_prod in (10001,10003,10005) order by cod_cli"
         data_lin.Refresh
      Else
         If Check3.Value = 1 Then
            data_lin.RecordSource = "Select * from llamado where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and matric not in (0) and codmot not in ('C') order by matric"
            data_lin.Refresh
         Else
            If checkesp.Value = 1 Then
               data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_prod in (2) order by cod_cli"
               data_lin.Refresh
            Else
               If checkped.Value = 1 Then
                  data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_prod in (14001,14002,14003) order by cod_cli"
                  data_lin.Refresh
               Else
                  If checkenf.Value = 1 Then
                     data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and nro_flia =" & 2 & " order by cod_cli"
                     data_lin.Refresh
                  Else
                     data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_prod in (10001,10003,10005,10002,10004,10006) order by cod_cli"
                     data_lin.Refresh
                  End If
               End If
            End If
         End If
      End If
   End If
   If data_lin.Recordset.RecordCount > 0 Then
      data_lin.Recordset.MoveLast
      data_lin.Recordset.MoveFirst
      pb1.Max = data_lin.Recordset.RecordCount
      pb1.Value = 0
      
      If Check3.Value = 1 Then
         Xlamat = data_lin.Recordset("matric")
      Else
         Xlamat = data_lin.Recordset("cod_cli")
      End If
      Do While Not data_lin.Recordset.EOF
         If Check3.Value = 1 Then
            If Xlamat = data_lin.Recordset("matric") Then
               Xca = Xca + 1
            Else
               If Xca > t_can.Text Then
                  data_lin.Recordset.MovePrevious
                  Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and matric not in (0) and codmot not in ('C') and matric =" & data_lin.Recordset("matric") & " order by fecha"
                  Data1.Refresh
                  If Data1.Recordset.RecordCount > 0 Then
                     Data1.Recordset.MoveFirst
                     Do While Not Data1.Recordset.EOF
                        data_inf.Recordset.AddNew
                        data_inf.Recordset("fecha") = Data1.Recordset("fecha")
                        If Data1.Recordset("codmot") = "R" Then
                           data_inf.Recordset("cod_prod") = 1
                        Else
                           If Data1.Recordset("codmot") = "A" Then
                              data_inf.Recordset("cod_prod") = 2
                           Else
                              data_inf.Recordset("cod_prod") = 3
                           End If
                        End If
                        data_inf.Recordset("nom_prod") = "LLAMADO CLAVE " & Data1.Recordset("codmot")
                        data_inf.Recordset("base") = Data1.Recordset("movilpas")
                        data_inf.Recordset("cod_cli") = Data1.Recordset("matric")
                        data_inf.Recordset("nom_cli") = Mid(Data1.Recordset("nombre"), 1, 30)
                        data_inf.Recordset("nro_med_a") = Data1.Recordset("codmed")
                        data_inf.Recordset("nom_med_a") = Data1.Recordset("nommed")
                        data_inf.Recordset("cantidad") = Xca
                        data_inf.Recordset("convenio") = Data1.Recordset("categ")
                        data_inf.Recordset.Update
                        Data1.Recordset.MoveNext
                     Loop
                  End If
                  data_lin.Recordset.MoveNext
               End If
               Xca = 1
            End If
            Xlamat = data_lin.Recordset("matric")
         
         Else
            If Xlamat = data_lin.Recordset("cod_cli") Then
               Xca = Xca + 1
            Else
               If Xca > t_can.Text Then
                  data_lin.Recordset.MovePrevious
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                  data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                  data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                  data_inf.Recordset("base") = data_lin.Recordset("base")
                  data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                  data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                  data_inf.Recordset("nro_med_a") = data_lin.Recordset("nro_med_a")
                  data_inf.Recordset("nom_med_a") = data_lin.Recordset("nom_med_a")
                  data_inf.Recordset("cantidad") = Xca
                  data_inf.Recordset("convenio") = data_lin.Recordset("convenio")
                  data_inf.Recordset.Update
                  data_lin.Recordset.MoveNext
               End If
               Xca = 1
            End If
            Xlamat = data_lin.Recordset("cod_cli")
         End If
         data_lin.Recordset.MoveNext
         pb1.Value = pb1.Value + 1
      Loop
      frm_infconssoc.MousePointer = 0
      MsgBox "Proceso terminado"
      If Option2.Value = True Then
         cr1.ReportFileName = App.path & "\infconssocn.rpt"
      Else
         cr1.ReportFileName = App.path & "\infconssoc.rpt"
      End If
      If Check1.Value = 1 Then
         cr1.ReportTitle = "Informe de Socios con mas de " & t_can.Text & " Consultas en domicilio. Período:" & md.Text & "--" & mh.Text
      Else
         If Check2.Value = 1 Then
            cr1.ReportTitle = "Informe de Socios con mas de " & t_can.Text & " Consultas en policlínica. Período:" & md.Text & "--" & mh.Text
         Else
            If checkesp.Value = 1 Then
               cr1.ReportTitle = "Informe de Socios con mas de " & t_can.Text & " Consultas ESPECIALISTAS. Período:" & md.Text & "--" & mh.Text
            Else
               If checkped.Value = 1 Then
                  cr1.ReportTitle = "Informe de Socios con mas de " & t_can.Text & " Consultas PEDIATRIA. Período:" & md.Text & "--" & mh.Text
               Else
                  If checkenf.Value = 1 Then
                     cr1.ReportTitle = "Informe de Socios con mas de " & t_can.Text & " Actos de Enfermería. Período:" & md.Text & "--" & mh.Text
                  Else
                     cr1.ReportTitle = "Informe de Socios con mas de " & t_can.Text & " Consultas. Período:" & md.Text & "--" & mh.Text
                  End If
               End If
            End If
         End If
      End If
      cr1.Action = 1
   End If
   frm_infconssoc.MousePointer = 0
End If


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.RecordSource = "infvtas"
'data_inf.Refresh
data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_temp.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_can.SetFocus
End If

End Sub
