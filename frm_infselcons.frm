VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infselcons 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe selección de consultas"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infselcons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   5430
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
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
      Top             =   4800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2880
      Top             =   5520
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
      Left            =   4560
      Picture         =   "frm_infselcons.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_infselcons.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Procesar"
      Top             =   5520
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Opciones de informe"
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   4800
      Width           =   4815
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Detalle"
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Resumen"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos de informes"
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin MSAdodcLib.Adodc data1 
         Height          =   330
         Left            =   2400
         Top             =   3000
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
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
         Caption         =   "data1"
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
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   375
         Left            =   360
         Top             =   2520
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Desde clave final"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   4200
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "frm_infselcons.frx":0F56
         Left            =   2400
         List            =   "frm_infselcons.frx":0F66
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox txt_cant 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Consultas en domicilio"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Si marca ésta opción, se emitirán SOLO consultas en domicilio"
         Top             =   3240
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Consultas en policlínica"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Si marca ésta opción, se emitirán solo consultas en policlínica"
         Top             =   2760
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_infselcons.frx":0F8C
         Left            =   2040
         List            =   "frm_infselcons.frx":0FA2
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2040
         Width           =   2655
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   2160
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
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "CLAVES:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Cant. Registros:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Familia:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Rango de fechas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1560
      Picture         =   "frm_infselcons.frx":0FE6
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   2295
   End
End
Attribute VB_Name = "frm_infselcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xtotregcon As Long
Dim Xvalor, Xtotvan, Xbande As Long
Xtotvan = 1
Xbande = 0
frm_infselcons.MousePointer = 11
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.ConnectionString = "dsn=" & Xconexrmt
Command1.Enabled = False
Command2.Enabled = False
'data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lin.ConnectionString = "dsn=" & Xconexrmt
data_inf.DatabaseName = App.path & "\informes.mdb"
data_inf.RecordSource = "infvtas"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If

Xtotregcon = 0
If mfd.Text <> "__/__/____" And mfh.Text <> "__/__/____" And txt_cant.Text <> "" Then
'   If Combo2.ListIndex < 0 Then
'      Combo2.ListIndex = 0
'   End If
   If Check1.Value = 1 Then
      If Combo2.ListIndex = 0 Then
         data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (10001,10003,10005,10009,10010,10011,14001,14002,14003,2) order by cod_prod"
         data_lin.Refresh
      Else
         If Combo2.ListIndex = 1 Then
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (10005,10009) order by cod_prod"
            data_lin.Refresh
         Else
            If Combo2.ListIndex = 2 Then
               data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (10003,10010) order by cod_prod"
               data_lin.Refresh
            Else
               If Combo2.ListIndex = 3 Then
                  data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (10001,10009) order by cod_prod"
                  data_lin.Refresh
               Else
                  If Combo1.ListIndex = 0 Then
                     data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (10001,10003,10005) order by cod_prod"
                     data_lin.Refresh
                  Else
                     If Combo1.ListIndex = 1 Then 'laboratorio
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia =" & 3 & " order by cod_prod"
                        data_lin.Refresh
                     Else
                        If Combo1.ListIndex = 2 Then 'traslados
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia =" & 9 & " order by cod_prod"
                           data_lin.Refresh
                        Else
                           If Combo1.ListIndex = 3 Then 'pediatria
                              data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (14001,14002,14003) order by cod_prod"
                              data_lin.Refresh
                           Else
                              If Combo1.ListIndex = 4 Then 'especialistas
                                 data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 2 & " order by cod_prod"
                                 data_lin.Refresh
                              Else
                                 If Combo1.ListIndex = 5 Then 'tasas
                                    data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia =" & 13 & " order by cod_prod"
                                    data_lin.Refresh
                                 Else
                                    data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (10001,10003,10005,10009,10010,10011,14001,14002,14003,2) order by cod_prod"
                                    data_lin.Refresh
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
   Else
      If Check2.Value = 1 Then
         If Combo2.ListIndex = 0 Then
            data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and trasla in (1,2,14,15) order by codmot"
            data_lin.Refresh
         Else
            If Combo2.ListIndex = 1 Then
               If Check3.Value = 1 Then
                  data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and colormot ='" & "R" & "' and base =" & 0 & " order by colormot"
               Else
                  data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codmot ='" & "R" & "' and base =" & 0 & " order by codmot"
               End If
               data_lin.Refresh
            Else
               If Combo2.ListIndex = 2 Then
                  If Check3.Value = 1 Then
                     data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and colormot ='" & "A" & "' and base =" & 0 & " order by colormot"
                  Else
                     data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codmot ='" & "A" & "' and base =" & 0 & " and trasla in (1,2,14,15) order by codmot"
                  End If
                  data_lin.Refresh
               Else
                  If Combo2.ListIndex = 3 Then
                     If Check3.Value = 1 Then
                        data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and colormot ='" & "V" & "' and base =" & 0 & " order by colormot"
                     Else
                        data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codmot ='" & "V" & "' and base =" & 0 & " order by codmot"
                     End If
                     data_lin.Refresh
                  Else
                     data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " order by codmot"
                     data_lin.Refresh
                  End If
               End If
            End If
         End If
      Else
         data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia =" & 1 & " order by cod_prod"
         data_lin.Refresh
      End If
   End If
   If txt_cant.Text > 0 Then
        If data_lin.Recordset.RecordCount > 0 Then
           data_lin.Recordset.MoveLast
           Xtotregcon = data_lin.Recordset.RecordCount
           If Xtotregcon > txt_cant.Text Then
              data_lin.Recordset.MoveFirst
              Randomize
              For Xtotvan = 1 To Val(txt_cant.Text) Step 1
                  Xvalor = CInt(Int((Xtotregcon * Rnd()) + 1))
                  For Xbande = 1 To Xvalor
                      If Xbande = Xvalor Then
                         If Check1.Value = 1 Then
                            If data_inf.Recordset.RecordCount > 0 Then
                               data_inf.RecordSource = "Select * from infvtas where cod_cli =" & data_lin.Recordset("cod_cli")
                               data_inf.Refresh
                               If data_inf.Recordset.RecordCount > 0 Then
                                  txt_cant.Text = txt_cant.Text + 1
                                  If Xtotregcon > txt_cant.Text Then
                                     data_inf.Recordset.Delete
                                     data_inf.RecordSource = "infvtas"
                                     data_inf.Refresh
                                  Else
                                     MsgBox "ATENCION:!! se excedió el límite de búsqueda con socio diferente", vbInformation, "Informes"
                                  End If
                               End If
                            End If
                         End If
                         If Check1.Value = 1 Then
                            data_inf.Recordset.AddNew
                            data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                            data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                            data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                            data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                            data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                            data_inf.Recordset("nro_med_a") = data_lin.Recordset("nro_med_a")
                            data_inf.Recordset("nom_med_a") = data_lin.Recordset("nom_med_a")
                            data_inf.Recordset("base") = data_lin.Recordset("base")
                            data_inf.Recordset("convenio") = data_lin.Recordset("convenio")
                            data_inf.Recordset("ced_socio") = data_lin.Recordset("ced_socio")
                            Data1.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                            Data1.Refresh
                            If Data1.Recordset.RecordCount > 0 Then
                               data_inf.Recordset("zona") = Data1.Recordset("cl_telefon")
                            End If
                            data_inf.Recordset.Update
                         Else
                            data_inf.Recordset.AddNew
                            If Check3.Value = 1 Then
                               If data_lin.Recordset("colormot") = "R" Then
                                  data_inf.Recordset("cod_prod") = 11111
                               End If
                               If data_lin.Recordset("colormot") = "A" Then
                                  data_inf.Recordset("cod_prod") = 11112
                               End If
                               If data_lin.Recordset("colormot") = "V" Then
                                  data_inf.Recordset("cod_prod") = 11113
                               End If
                               data_inf.Recordset("nom_prod") = "CONSULTA DOM. " & data_lin.Recordset("colormot")
                            Else
                               If data_lin.Recordset("codmot") = "R" Then
                                  data_inf.Recordset("cod_prod") = 11111
                               End If
                               If data_lin.Recordset("codmot") = "A" Then
                                  data_inf.Recordset("cod_prod") = 11112
                               End If
                               If data_lin.Recordset("codmot") = "V" Then
                                  data_inf.Recordset("cod_prod") = 11113
                               End If
                               data_inf.Recordset("nom_prod") = "CONSULTA DOM. " & data_lin.Recordset("codmot")
                            End If
                            data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                            data_inf.Recordset("cod_cli") = data_lin.Recordset("matric")
                            data_inf.Recordset("nom_cli") = Mid(data_lin.Recordset("nombre"), 1, 30)
                            data_inf.Recordset("nro_med_a") = data_lin.Recordset("codmed")
                            data_inf.Recordset("nom_med_a") = data_lin.Recordset("nommed")
                            data_inf.Recordset("base") = data_lin.Recordset("movilpas")
                            data_inf.Recordset("convenio") = data_lin.Recordset("categ")
                            data_inf.Recordset("ced_socio") = data_lin.Recordset("ci")
                            Data1.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric")
                            Data1.Refresh
                            If Data1.Recordset.RecordCount > 0 Then
                               data_inf.Recordset("zona") = Data1.Recordset("cl_telefon")
                            End If
                            data_inf.Recordset.Update
                         End If
                      End If
                      If data_lin.Recordset.EOF = True Then
                      Else
                         data_lin.Recordset.MoveNext
                      End If
                  Next
                  data_lin.Recordset.MoveFirst
              Next
              frm_infselcons.MousePointer = 0
              Command1.Enabled = True
              Command2.Enabled = True
              MsgBox "Proceso terminado"
              data_inf.RecordSource = "infvtas"
              data_inf.Refresh
              If Option1.Value = True Then
                 cr1.ReportFileName = App.path & "\infselconsn.rpt"
              Else
                 If Check1.Value = 1 Then
                    cr1.ReportFileName = App.path & "\infselconsd.rpt"
                 Else
                    cr1.ReportFileName = App.path & "\infselconsd2.rpt"
                 End If
              End If
              cr1.ReportTitle = "INFORME DE CONSULTAS SELECCIONADAS EN FORMA ALEATORIA: " & mfd.Text & " -- " & mfh.Text
              cr1.Action = 1
           Else
              MsgBox "La cantidad de registros son menores a la cantidad a procesar " & str(data_lin.Recordset.RecordCount)
           End If
        End If
   End If
End If

frm_infselcons.MousePointer = 0
Command1.Enabled = True
Command2.Enabled = True

End Sub

Private Sub Command2_Click()
Unload Me

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
   txt_cant.SetFocus
End If

End Sub

Private Sub txt_cant_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub
