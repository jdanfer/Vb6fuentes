VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infctrolcons 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de control tiempos de consulta en policlínica."
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infctrolcons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7230
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_est 
      Caption         =   "data_est"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_infctrolcons.frx":0442
      Height          =   2175
      Left            =   120
      OleObjectBlob   =   "frm_infctrolcons.frx":0459
      TabIndex        =   18
      Top             =   5400
      Width           =   6975
   End
   Begin VB.Data data_res 
      Caption         =   "data_res"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   3600
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6240
      Picture         =   "frm_infctrolcons.frx":0E30
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_infctrolcons.frx":13BA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Procesar"
      Top             =   4320
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos para informe..."
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   375
         Left            =   840
         Top             =   2640
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "data_lin"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   4440
         TabIndex        =   17
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Emitir para indicadores"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   3000
         Width           =   3975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Resumen"
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   3600
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Detalle"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3600
         Value           =   -1  'True
         Width           =   2775
      End
      Begin MSMask.MaskEdBox mhhor 
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mdhor 
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox t_base 
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox t_serv 
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   1680
         Width           =   1575
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   4680
         TabIndex        =   2
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "En orden ascendente"
         Height          =   495
         Left            =   5040
         TabIndex        =   15
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Rango de Horas:"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "BASE: (99=TODAS)"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Servicio: (999=TODOS)"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Rango de fechas:"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   4440
      Picture         =   "frm_infctrolcons.frx":1944
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   2055
   End
End
Attribute VB_Name = "frm_infctrolcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xh1, Xm1, Xh2, Xm2 As Long
Dim Xresh, Xresm As Long
Dim Xesp, Xmgral, Xmgralenf As String
Dim Xtotcon, Xporconm, Xporcone, Xporconenf As Long
Xesp = ""
Xmgral = ""
Xmgralenf = ""
Xtotcon = 0
Xporconm = 0
Xporconenf = 0
Xporcone = 0
frm_infctrolcons.MousePointer = 11
Command1.Enabled = False
Command2.Enabled = False

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"
data_inf.RecordSource = "infvtas"
data_inf.Refresh

MiBaseact.Execute "Delete * from infarqc"
data_res.RecordSource = "infarqc"
data_res.Refresh

If mfd.Text = "__/__/____" Then
   MsgBox "Sin Fecha desde"
Else
   If mfh.Text = "__/__/____" Then
      MsgBox "Sin Fecha hasta"
   Else
      If mhhor.Text = "__:__" Then
         If t_serv.Text = 999 Then
            If t_base.Text = 99 Then
               If Check1.Value = 1 Then
                  data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And cod_prod in (20001,20017,20003,30048,20043,20042,20053,20065,20070,20085,20091,20099,20106,20074,20048,20051,20063,20083,20097)"
                  data_lin.Refresh
               Else
                  data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And nro_flia in (1,2,14)"
                  data_lin.Refresh
               End If
            Else
               data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And nro_flia in (1,2,14) And base =" & t_base.Text
               data_lin.Refresh
            End If
         Else
            If t_base.Text = 99 Then
               data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And servicio =" & 1 & " And cod_prod =" & t_serv.Text
               data_lin.Refresh
            Else
               data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And servicio =" & 1 & " And base =" & t_base.Text & " And cod_prod =" & t_serv.Text
               data_lin.Refresh
            End If
         End If
      Else
         If t_serv.Text = 999 Then
            If t_base.Text = 99 Then
               data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And servicio =" & 1 & " And hora >='" & mdhor.Text & "' And hora <='" & mhhor.Text & "'"
               data_lin.Refresh
            Else
               data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And servicio =" & 1 & " And base =" & t_base.Text & " And hora >='" & mdhor.Text & "' And hora <='" & mhhor.Text & "'"
               data_lin.Refresh
            End If
         Else
            If t_base.Text = 99 Then
               data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And servicio =" & 1 & " And cod_prod =" & t_serv.Text & " And hora >='" & mdhor.Text & "' And hora <='" & mhhor.Text & "'"
               data_lin.Refresh
            Else
               data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And servicio =" & 1 & " And base =" & t_base.Text & " And cod_prod =" & t_serv.Text & " And hora >='" & mdhor.Text & "' And hora <='" & mhhor.Text & "'"
               data_lin.Refresh
            End If
         End If
      End If
      Dim Xtotmas15 As Double
      If data_lin.Recordset.RecordCount > 0 Then
         data_lin.Recordset.MoveFirst
         Do While Not data_lin.Recordset.EOF
            data_inf.Recordset.AddNew
            data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
            data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
            data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
            data_inf.Recordset("hora") = data_lin.Recordset("hora")
            data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
            data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
            data_inf.Recordset("convenio") = data_lin.Recordset("convenio")
            data_inf.Recordset("base") = data_lin.Recordset("base")
            data_inf.Recordset("nom_superv") = Mid(data_lin.Recordset("nom_superv"), 1, 5)
            data_inf.Recordset("nro_med_s") = data_lin.Recordset("nro_med_s")
            data_inf.Recordset("nom_med_s") = data_lin.Recordset("nom_med_s")
            data_inf.Recordset("nro_flia") = data_lin.Recordset("nro_flia")
            If IsNull(data_lin.Recordset("servicio")) = False Then
                If IsNull(data_lin.Recordset("hora")) = False Then
                   Xh1 = Val(Mid(Trim(data_lin.Recordset("hora")), 1, 2))
                   Xm1 = Val(Mid(Trim(data_lin.Recordset("hora")), 4, 2))
                Else
                   Xh1 = 0
                   Xm1 = 0
                End If
                If IsNull(data_lin.Recordset("nom_superv")) = False Then
                   Xh2 = Val(Mid(Trim(data_lin.Recordset("nom_superv")), 1, 2))
                   Xm2 = Val(Mid(Trim(data_lin.Recordset("nom_superv")), 4, 2))
                Else
                   Xh2 = 0
                   Xm2 = 0
                End If
                Xresh = Xh2 - Xh1
                If Xresh < 0 Then
                   Xresh = 0
                End If
                Xresm = Xm2 - Xm1
                If Xresm < 0 Then
                   Xresm = Xresm + 60
                   Xresh = Xresh - 1
                End If
            Else
                Xresm = 0
                Xresh = 0
            End If
            If Xresh < 0 Then
               If Xresm < 0 Then
                  data_inf.Recordset("zona") = "00:00"
               Else
                  data_inf.Recordset("zona") = "00:" + Trim(str(Xresm))
               End If
            Else
               If Xresm < 0 Then
                  data_inf.Recordset("zona") = Trim(str(Xresh)) + ":00"
               Else
                  data_inf.Recordset("zona") = Trim(str(Xresh)) + ":" + Trim(str(Xresm))
               End If
               If Xresh < 10 Then
                  If Xresm < 10 Then
                     data_inf.Recordset("zona") = "0" + Trim(str(Xresh)) + ":" + "0" + Trim(str(Xresm))
                  Else
                     data_inf.Recordset("zona") = "0" + Trim(str(Xresh)) + ":" + Trim(str(Xresm))
                  End If
               Else
                  If Xresm < 10 Then
                     data_inf.Recordset("zona") = Trim(str(Xresh)) + ":" + "0" + Trim(str(Xresm))
                  Else
                     data_inf.Recordset("zona") = Trim(str(Xresh)) + ":" + Trim(str(Xresm))
                  End If
               End If
            End If
            If data_lin.Recordset("cod_prod") = 2 Then
               If Xresh >= 1 Then
                  Xesp = "NO"
               Else
                  Xporcone = Xporcone + 1
                  If Xesp <> "NO" Then
                     Xesp = "SI"
                  End If
               End If
            Else
               If data_lin.Recordset("nro_flia") = 2 Then
                  If Xresh >= 1 Then
                     Xmgralenf = "NO"
                     Xtotmas15 = Xtotmas15 + 1
                  Else
                     If Xresm > 15 Then
                        Xmgralenf = "NO"
                        Xtotmas15 = Xtotmas15 + 1
                     Else
                        Xporconenf = Xporconenf + 1
                        If Xmgralenf <> "NO" Then
                           Xmgralenf = "SI"
                        End If
                     End If
                  End If
               Else
                  If Xresh >= 1 Then
                     Xmgral = "NO"
                  Else
                     If Xresm > 30 Then
                        Xmgral = "NO"
                     Else
                        Xporconm = Xporconm + 1
                        If Xmgral <> "NO" Then
                           Xmgral = "SI"
                        End If
                     End If
                  End If
               End If
            End If
            data_inf.Recordset.Update
            data_lin.Recordset.MoveNext
            Xtotcon = Xtotcon + 1
         Loop
         MsgBox "Proceso terminado"
         data_inf.RecordSource = "Select * from infvtas"
         data_inf.Refresh
         data_res.Recordset.AddNew
         data_res.Recordset("desc1") = Xmgral
         data_res.Recordset("desc2") = Xesp
         data_res.Recordset("cob") = Xtotmas15
         If Xtotcon = 0 Then
            data_res.Recordset("mesarq") = 0
            data_res.Recordset("anoarq") = 0
            data_res.Recordset("totimp") = 0
         Else
            data_res.Recordset("mesarq") = Xporconm / Xtotcon * 100
            data_res.Recordset("anoarq") = Xporcone / Xtotcon * 100
            data_res.Recordset("totimp") = Xporconenf / Xtotcon * 100
         End If
         data_res.Recordset.Update
         data_res.RecordSource = "Select * from infarqc"
         data_res.Refresh
         If data_res.Recordset.RecordCount > 0 Then
            data_res.Recordset.MoveFirst
         End If
'         If Check1.value = 1 Then
            cr1.ReportTitle = "Informe demoras de ACTOS ENFERMERIA DESDE:" & mfd.Text & " HASTA:" & mfh.Text
            If Option2.Value = True Then
               cr1.DiscardSavedData = True
               cr1.ReportFileName = App.path & "\infctrenfen.rpt"
            Else
               cr1.DiscardSavedData = True
               cr1.ReportFileName = App.path & "\infctrenfe.rpt"
            End If
'         Else
'            cr1.ReportTitle = "Informe demoras de consultas en policlínica DESDE:" & mfd.Text & " HASTA:" & mfh.Text
'            If Option2.value = True Then
'               cr1.DiscardSavedData = True
'               cr1.ReportFileName = App.Path & "\infctrconsn.rpt"
'            Else
'               cr1.DiscardSavedData = True
'               cr1.ReportFileName = App.Path & "\infctrcons.rpt"
'            End If
'         End If

         cr1.Action = 1
      Else
         MsgBox "No existen registros"
      End If
   End If
End If
       
frm_infctrolcons.MousePointer = 0
Command1.Enabled = True
Command2.Enabled = True
   
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
frm_infctrolcons.Height = 8145
DBGrid1.SetFocus

End Sub

Private Sub Form_Load()
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.RecordSource = "infvtas"
'data_inf.Refresh
data_lin.ConnectionString = "dsn=" & Xconexrmt
data_res.DatabaseName = App.path & "\informes.mdb"
'data_res.RecordSource = "infarqc"
'data_res.Refresh
data_est.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_est.RecordSource = "Select * from estudios order by descrip"
data_est.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

Private Sub mfh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_serv.SetFocus
End If

End Sub

Private Sub t_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub

Private Sub t_serv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_base.SetFocus
End If

End Sub
