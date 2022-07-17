VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_vtasservjv 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas por Servicio con datos de clientes"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7515
   Icon            =   "frm_vtasservjv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7515
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3840
      Top             =   4320
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar barra 
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin Crystal.CrystalReport crs2 
      Left            =   6960
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crs 
      Left            =   2280
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
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
      RecordSource    =   "infvtas"
      Top             =   480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton b_canc 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      MouseIcon       =   "frm_vtasservjv.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "frm_vtasservjv.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton b_acep 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      MouseIcon       =   "frm_vtasservjv.frx":0CD6
      MousePointer    =   99  'Custom
      Picture         =   "frm_vtasservjv.frx":0FE0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Procesar"
      Top             =   4200
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos para informe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      Begin MSAdodcLib.Adodc data_emi 
         Height          =   330
         Left            =   4440
         Top             =   1320
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         Caption         =   "data_emi"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc data2 
         Height          =   375
         Left            =   360
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         Caption         =   "data2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc data1 
         Height          =   375
         Left            =   2280
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
         Caption         =   "data1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   375
         Left            =   4560
         Top             =   2400
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Informe sin detalle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox tm 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Data data_med 
         Caption         =   "data_med"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "estudios"
         Top             =   1800
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txt_b 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1680
         Width           =   735
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "frm_vtasservjv.frx":156A
         Height          =   360
         Left            =   1560
         TabIndex        =   5
         Top             =   960
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648384
         ForeColor       =   0
         ListField       =   "DESCRIP"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Left            =   2040
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "BASE: (99=TODAS)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Servicio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Rango de fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   1440
      Picture         =   "frm_vtasservjv.frx":1581
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1815
   End
End
Attribute VB_Name = "frm_vtasservjv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b_acep_Click()
b_acep.Enabled = False
b_canc.Enabled = False
If DBCombo1.Text = "" Then
   tm.Text = 999999
End If
Dim Xemitim As String
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"

data_inf.RecordSource = "infvtas"
data_inf.Refresh

If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      If tm.Text <> "" Then
         If txt_b.Text <> "" Then
            If txt_b.Text = 99 Then
               If tm.Text = 999999 Then
                  data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 10001
                  data_lin.Refresh
               Else
                  data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod =" & tm.Text
                  data_lin.Refresh
               End If
            Else
               If tm.Text = 999999 Then
                  data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & txt_b.Text
                  data_lin.Refresh
               Else
                  data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod =" & tm.Text & " And base =" & txt_b.Text
                  data_lin.Refresh
               End If
            End If
            If data_lin.Recordset.RecordCount > 0 Then
'               data_lin.Recordset.MoveLast
               barra.Visible = True
               barra.Max = data_lin.Recordset.RecordCount
               barra.value = 0
               data_lin.Recordset.MoveFirst
               frm_vtasserv.MousePointer = 11
               Do While Not data_lin.Recordset.EOF
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                  data_inf.Recordset("factura") = data_lin.Recordset("factura")
                  data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                  data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                  data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                  data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                  data_inf.Recordset("nro_flia") = data_lin.Recordset("nro_flia")
                  data_inf.Recordset("nom_flia") = data_lin.Recordset("nom_flia")
                  Data1.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                  Data1.Refresh
                  If Data1.Recordset.RecordCount > 0 Then
                     If IsNull(Data1.Recordset("cl_dpto")) = False Then
                        data_inf.Recordset("nom_med_a") = Mid(Data1.Recordset("cl_dpto"), 1, 40) 'celular
                     End If
                     If IsNull(Mid(Data1.Recordset("cl_telefon"), 1, 40)) = False Then
                        data_inf.Recordset("nom_med_s") = Mid(Data1.Recordset("cl_telefon"), 1, 40)
                     End If
                     If IsNull(Data1.Recordset("cl_direcci")) = False Then
                        data_inf.Recordset("nom_medic") = Mid(Data1.Recordset("cl_direcci"), 1, 50)
                     End If
                     If IsNull(Data1.Recordset("cl_zona")) = False Then
                        data_inf.Recordset("rub_nomb") = Mid(Data1.Recordset("cl_zona"), 1, 40)
                     End If
                     If IsNull(Data1.Recordset("cl_fnac")) = False Then
                        data_inf.Recordset("vto") = Data1.Recordset("cl_fnac")
                     End If
                     If IsNull(data_lin.Recordset("convenio")) = False Then
                        data_inf.Recordset("convenio") = data_lin.Recordset("convenio")
                        Adodc1.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
                        Adodc1.Refresh
                        If Adodc1.Recordset.RecordCount > 0 Then
                           If IsNull(Adodc1.Recordset("cnv_grupo")) = False Then
                              If Adodc1.Recordset("cnv_grupo") <> "" Then
                                 data_inf.Recordset("nom_superv") = "MUTUAL/CONV"
                              Else
                                 If Adodc1.Recordset("cnv_emite") = "S" Then
                                    If IsNull(Adodc1.Recordset("cnv_colrec")) = False Then
                                       If Adodc1.Recordset("cnv_colrec") = "R" Or Adodc1.Recordset("cnv_colrec") = "M" Or _
                                          Adodc1.Recordset("cnv_colrec") = "A" Or Adodc1.Recordset("cnv_colrec") = "V" Then
                                          data_inf.Recordset("nom_superv") = "SAPP"
                                       Else
                                          data_inf.Recordset("nom_superv") = "OTRO"
                                       End If
                                    Else
                                       data_inf.Recordset("nom_superv") = "OTRO"
                                    End If
                                 Else
                                    data_inf.Recordset("nom_superv") = "OTRO"
                                 End If
                              End If
                           Else
                              data_inf.Recordset("nom_superv") = "OTRO"
                           End If
                        Else
                           data_inf.Recordset("nom_superv") = "OTRO"
                        End If
                     End If
                  End If
                  data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                  data_inf.Recordset("nro_med_a") = data_lin.Recordset("nro_med_a")
                  data_inf.Recordset("mes_paga") = data_lin.Recordset("mes_paga")
                  data_inf.Recordset("ano_paga") = data_lin.Recordset("ano_paga")
                  data_inf.Recordset("base") = data_lin.Recordset("base")
                  data_inf.Recordset("hora") = data_lin.Recordset("hora")
                  If IsNull(data_lin.Recordset("grupo")) = False Then
                     data_inf.Recordset("nro_superv") = data_lin.Recordset("grupo")
                  Else
                     data_inf.Recordset("nro_superv") = 0
                  End If
                  If IsNull(data_lin.Recordset("ced_socio")) = True Then
                     data_inf.Recordset("zona") = "0"
                  Else
                     If IsNull(data_lin.Recordset("fact")) = True Then
                        data_inf.Recordset("zona") = Trim(Str(data_lin.Recordset("ced_socio"))) & "-0"
                     Else
                        data_inf.Recordset("zona") = Trim(Str(data_lin.Recordset("ced_socio"))) & "-" & Trim(Str(data_lin.Recordset("fact")))
                     End If
                  End If
                  data_inf.Recordset.Update
                  data_lin.Recordset.MoveNext
                  barra.Max = data_lin.Recordset.RecordCount
                  barra.value = barra.value + 1
               Loop
               data_inf.RecordSource = "Select * from infvtas order by cod_prod"
               data_inf.Refresh
               frm_vtasserv.MousePointer = 0
               If Check1.value = 1 Then
                  crs.ReportFileName = App.Path & "\infvtasxsern.rpt"
                  If txt_b.Text = 99 Then
                     crs.ReportTitle = "INFORME VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                  Else
                     crs.ReportTitle = "INFORME DE VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " BASE: " & txt_b.Text
                  End If
                  crs.Action = 1
               Else
                  crs.ReportFileName = App.Path & "\infvtasxserus.rpt"
                  If txt_b.Text = 99 Then
                     crs.ReportTitle = "INFORME DE VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                  Else
                     crs.ReportTitle = "INFORME DE VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & "  ==BASE: " & txt_b.Text
                  End If
                  crs.Action = 1
               End If
               barra.Visible = False
            Else
               MsgBox "No existen registros con esta selección", vbInformation, "Mensaje"
            End If
         Else
            MsgBox "Ingrese Base", vbInformation, "Mensaje"
            txt_b.SetFocus
         End If
      Else
         MsgBox "Número de SERVICIO incorrecto", vbInformation, "Mensaje"
         DBCombo1.SetFocus
      End If
   Else
      MsgBox "Ingrese Fecha", vbInformation, "Mensaje"
      mh.SetFocus
   End If
Else
   MsgBox "Ingrese fecha", vbInformation, "Mensaje"
   md.SetFocus
End If
b_acep.Enabled = True
b_canc.Enabled = True
               
End Sub

Private Sub b_canc_Click()
Unload Me

End Sub

Private Sub DBCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_b.SetFocus
End If

End Sub

Private Sub DBCombo1_LostFocus()

If DBCombo1.Text <> "" Then
    If IsNumeric(DBCombo1.Text) = True Then
       data_med.Recordset.FindFirst "codest =" & DBCombo1.Text
       If Not data_med.Recordset.NoMatch Then
          tm.Text = data_med.Recordset("codest")
          DBCombo1.Text = data_med.Recordset("descrip")
       Else
          MsgBox "No encontrado, consulte por nombre", vbInformation, "Mensaje"
          DBCombo1.SetFocus
       End If
    Else
       data_med.Recordset.FindFirst "descrip ='" & DBCombo1.Text & "'"
       If Not data_med.Recordset.NoMatch Then
          tm.Text = data_med.Recordset("codest")
          DBCombo1.Text = data_med.Recordset("descrip")
       Else
          MsgBox "No encontrado, consulte por nombre", vbInformation, "Mensaje"
          DBCombo1.SetFocus
       End If
    End If
Else
    tm.Text = 999999
    
End If

End Sub

Private Sub Form_Load()
data_med.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_med.RecordSource = "Select * from estudios order by descrip"
data_med.Refresh
data_inf.DatabaseName = App.Path & "\informes.mdb"
data_lin.ConnectionString = "dsn=" & Xconexrmt
Data1.ConnectionString = "dsn=" & Xconexrmt
'Data1.RecordSource = "clientes"
'Data1.Refresh
'Data2.DatabaseName = ""
Data2.ConnectionString = "dsn=" & Xconexrmt
data_emi.ConnectionString = "dsn=" & Xconexrmt
'Data2.RecordSource = "Select * from cabezal"
'Data2.Refresh
Adodc1.ConnectionString = "dsn=" & Xconexrmt 'convenios


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
   DBCombo1.SetFocus
End If

End Sub

Private Sub txt_b_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_acep.SetFocus
End If

End Sub
