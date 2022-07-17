VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_infconsultas 
   BackColor       =   &H00C000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de consultas por socios"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6660
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infconsultas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6660
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_lin 
      Caption         =   "data_lin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2640
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5880
      Picture         =   "frm_infconsultas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Salir"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_infconsultas.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Procesar"
      Top             =   3240
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Datos del informe"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Desde respaldos"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "Resumen"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   2160
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_infconsultas.frx":0F56
         Left            =   2280
         List            =   "frm_infconsultas.frx":0F63
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox t_fam 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2280
         TabIndex        =   5
         Text            =   "99"
         Top             =   960
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "Socios:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Familia (99=Todas)"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "FECHAS:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   4680
      Picture         =   "frm_infconsultas.frx":0F7E
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "frm_infconsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xlaemi As String
If Month(Date) > 9 Then
   Xlaemi = "emi" & Trim(str(Month(Date))) & Mid(Trim(str(Year(Date))), 3, 2)
Else
   Xlaemi = "emi" & "0" & Trim(str(Month(Date))) & Mid(Trim(str(Year(Date))), 3, 2)
End If

If md.Text <> "__/__/____" Then
   frm_infconsultas.MousePointer = 11
   If Combo1.ListIndex = 1 Then
      data_cli.DatabaseName = ""
      data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
      data_cli.RecordSource = "Select * from " & Trim(Xlaemi) & " order by nro_cobr"
      data_cli.Refresh
      data_inf.RecordSource = "infcli"
      data_inf.Refresh
   Else
      data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
      data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','UCM','CASH','MSP','CEMSA','SJ01','SJ02','CCASMU','911','911B','1727')"
      data_cli.Refresh
      data_inf.RecordSource = "infcli"
      data_inf.Refresh
   End If
   If Check1.Value = 1 Then
      data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and nroflia in(1,2,3,5,6,7,9,10,14,16)"
      data_lin.Refresh
   Else
'      data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and nro_flia in(1,2,3,5,6,7,9,10,14,16)"
      data_lin.Refresh
   End If
   If data_inf.Recordset.RecordCount > 0 Then
      data_inf.Recordset.MoveFirst
      Do While Not data_inf.Recordset.EOF
         data_inf.Recordset.Delete
         data_inf.Recordset.MoveNext
      Loop
   End If
   If data_cli.Recordset.RecordCount > 0 Then
      data_cli.Recordset.MoveFirst
      Do While Not data_cli.Recordset.EOF
         If Combo1.ListIndex = 1 Then
            data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_cli =" & data_cli.Recordset("cliente")
            data_lin.Refresh
'            data_lin.Recordset.FindFirst "cod_cli =" & data_cli.Recordset("cliente")
         Else
            data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_cli =" & data_cli.Recordset("cl_codigo")
'            data_lin.Recordset.FindFirst "cod_cli =" & data_cli.Recordset("cl_codigo")
            data_lin.Refresh
         End If
         If Not data_lin.Recordset.RecordCount > 0 Then
         Else
            If Combo1.ListIndex = 1 Then
                data_inf.Recordset.AddNew
                data_inf.Recordset("cl_codigo") = data_cli.Recordset("cliente")
                data_inf.Recordset("cl_apellid") = data_cli.Recordset("apellidos")
                data_inf.Recordset("cl_cedula") = data_cli.Recordset("cedula")
                data_inf.Recordset("cl_telefon") = data_cli.Recordset("tel_cli")
                data_inf.Recordset("cl_grupo") = data_cli.Recordset("grupo")
                data_inf.Recordset("cl_zona") = data_cli.Recordset("zona")
                data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("nro_cobr")
                data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("nom_cobr")
                data_inf.Recordset("cl_fecing") = data_cli.Recordset("fecha_ing")
                data_inf.Recordset("cl_codconv") = data_cli.Recordset("cod_cnv")
                Data1.RecordSource = "Select * from clientes where cl_codigo =" & data_cli.Recordset("cliente")
                Data1.Refresh
                If Data1.Recordset.RecordCount > 0 Then
                   data_inf.Recordset("cl_ultmesp") = Data1.Recordset("cl_ultmesp")
                   data_inf.Recordset("cl_ultanop") = Data1.Recordset("cl_ultanop")
                End If
                data_inf.Recordset.Update
            Else
                data_inf.Recordset.AddNew
                data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                data_inf.Recordset.Update
            End If
         End If
         data_cli.Recordset.MoveNext
      Loop
      frm_infconsultas.MousePointer = 0
      MsgBox "Proceso terminado"
      data_inf.RecordSource = "Select * from infcli order by cl_nrocobr"
      data_inf.Refresh
      If Option1.Value = True Then
         cr1.ReportFileName = App.path & "\infsincons.rpt"
         cr1.ReportTitle = "Informe de socios sin consultas desde: " & md.Text & " HASTA: " & mh.Text
         cr1.Action = 1
      Else
         cr1.ReportFileName = App.path & "\infsinconsn.rpt"
         cr1.ReportTitle = "Informe de socios sin consultas desde: " & md.Text & " HASTA: " & mh.Text
         cr1.Action = 1
      End If
   End If
End If

      
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_inf.DatabaseName = App.path & "\informes.mdb"
data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

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
   t_fam.SetFocus
End If

End Sub
