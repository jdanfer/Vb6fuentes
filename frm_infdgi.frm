VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_infdgi 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes para DGI"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infdgi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6330
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr1 
      Left            =   2400
      Top             =   2760
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
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
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
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5640
      Picture         =   "frm_infdgi.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_infdgi.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Procesar"
      Top             =   5400
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Formato de informe"
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   4560
      Width           =   5895
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FF8080&
         Caption         =   "Resumen"
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FF8080&
         Caption         =   "Detalle"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos de informe"
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      Begin VB.Data data_busca 
         Caption         =   "data_busca"
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
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3840
         TabIndex        =   12
         Top             =   480
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
         Left            =   2040
         TabIndex        =   11
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FF8080&
         Caption         =   "Facturación con RUT"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3000
         Width           =   3855
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF8080&
         Caption         =   "Ventas emisión con RUT"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2400
         Width           =   3855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "Total de compras por comercio"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   3855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Total empresas registradas"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Value           =   -1  'True
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "FECHAS:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frm_infdgi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xemi As String
frm_infdgi.MousePointer = 11
Command1.Enabled = False

data_inf.RecordSource = "infcli"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If
If md.Text = "__/__/____" Then
   md.Text = Date
End If
If mh.Text = "__/__/____" Then
   mh.Text = Date
End If
If Month(md.Text) > 9 Then
   Xemi = "EMI" & Trim(Str(Month(md.Text)))
Else
   Xemi = "EMI0" & Trim(Str(Month(md.Text)))
End If
Xemi = Xemi & Mid(Trim(Str(Year(md.Text))), 3, 2)

If Option1.value = True Then
   Data1.RecordSource = "Select * from abmdesp where base <>" & 99 & " and base <>" & 97
   Data1.Refresh
Else
   If Option2.value = True Then
      Data1.RecordSource = "Select * from tesorero where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by fecha"
      Data1.Refresh
   Else
      If Option3.value = True Then
         Data1.DatabaseName = ""
         Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
         Data1.RecordSource = Trim(Xemi)
         Data1.Refresh
      Else
         If Option4.value = True Then
            Data1.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by fecha"
            Data1.Refresh
         Else
            Data1.RecordSource = "Select * from abmdesp where base <>" & 99 & " and base <>" & 97
            Data1.Refresh
         End If
      End If
   End If
End If
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   If Option1.value = True Then
      Do While Not Data1.Recordset.EOF
         data_inf.Recordset.AddNew
         data_inf.Recordset("cl_codigo") = Data1.Recordset("nro")
         data_inf.Recordset("cl_direcci") = Mid(Data1.Recordset("obsmot"), 1, 80)
         data_inf.Recordset("cl_apellid") = Mid(Data1.Recordset("obs"), 1, 50)
         data_inf.Recordset("cl_cedula") = Data1.Recordset("mat")
         data_inf.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
      Command1.Enabled = True
      frm_infdgi.MousePointer = 0
      MsgBox "Proceso terminado"
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         data_inf.Refresh
         cr1.ReportFileName = App.Path & "\infdgi1.rpt"
         cr1.ReportTitle = "INFORME DE COMERCIOS REGISTRADOS"
         cr1.Action = 1
      End If
   End If
   If Option2.value = True Then
      If Data1.Recordset.RecordCount > 0 Then
         Data1.Recordset.MoveFirst
         Do While Not Data1.Recordset.EOF
            If IsNull(Data1.Recordset("bandera")) = False Then
               If Data1.Recordset("bandera") > 0 Then
                  data_busca.RecordSource = "Select * from abmdesp where nro =" & Data1.Recordset("bandera") & " and base not in (99,97)"
                  data_busca.Refresh
                  If data_busca.Recordset.RecordCount > 0 Then
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("cl_codigo") = data_busca.Recordset("nro")
                     data_inf.Recordset("cl_direcci") = Mid(data_busca.Recordset("obsmot"), 1, 80)
                     data_inf.Recordset("cl_cedula") = data_busca.Recordset("mat")
                     data_inf.Recordset("cl_fnac") = Data1.Recordset("fecha")
                     data_inf.Recordset("saldo_cc") = Data1.Recordset("monto")
                     data_inf.Recordset("cl_apellid") = Mid(Data1.Recordset("obs"), 1, 50)
                     If Data1.Recordset("moneda") = 2 Then
                        data_inf.Recordset("cl_codconv") = "U$s."
                     Else
                        data_inf.Recordset("cl_codconv") = "$."
                     End If
                     data_inf.Recordset.Update
                  End If
               End If
            End If
            Data1.Recordset.MoveNext
         Loop
      End If
      Command1.Enabled = True
      frm_infdgi.MousePointer = 0
      MsgBox "Proceso terminado"
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         data_inf.Refresh
         cr1.ReportFileName = App.Path & "\infdgi2.rpt"
         cr1.ReportTitle = "INFORME DE COMPRAS REGISTRADAS POR COMERCIO FECHA: " & md.Text & " HASTA: " & mh.Text
         cr1.Action = 1
      End If
   
   End If
   If Option3.value = True Then
      If Data1.Recordset.RecordCount > 0 Then
         Data1.Recordset.MoveFirst
         Do While Not Data1.Recordset.EOF
            If IsNull(Data1.Recordset("ruc")) = False Then
               If Val(Data1.Recordset("ruc")) > 0 Then
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_codigo") = Data1.Recordset("cliente")
                  data_inf.Recordset("cl_direcci") = Mid(Data1.Recordset("apellidos"), 1, 60)
                  data_inf.Recordset("cl_nombre") = Data1.Recordset("ruc")
                  data_inf.Recordset("saldo_cc") = Data1.Recordset("total")
                  data_inf.Recordset("cl_nomconv") = Mid(Data1.Recordset("nom_cnv"), 1, 30)
                  data_inf.Recordset("cl_ultmesp") = Data1.Recordset("mes")
                  data_inf.Recordset("cl_ultanop") = Data1.Recordset("ano")
                  data_inf.Recordset.Update
               End If
            End If
            Data1.Recordset.MoveNext
         Loop
      End If
      Command1.Enabled = True
      frm_infdgi.MousePointer = 0
      MsgBox "Proceso terminado"
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         data_inf.Refresh
         cr1.ReportFileName = App.Path & "\infdgi3.rpt"
         cr1.ReportTitle = "INFORME DE RECIBOS EMISION CON RUT FECHA: " & md.Text & " HASTA: " & mh.Text
         cr1.Action = 1
      End If
   
   End If

End If


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_inf.DatabaseName = App.Path & "\informes.mdb"
data_busca.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Option1.SetFocus
End If

End Sub
