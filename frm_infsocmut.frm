VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infsocmut 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de socios por mutualistas"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frm_infsocmut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6855
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   3480
      Top             =   3600
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
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Picture         =   "frm_infsocmut.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Picture         =   "frm_infsocmut.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Procesar"
      Top             =   3960
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Datos para  informe"
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
      Width           =   6375
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2160
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00800000&
         Caption         =   "Incluir consultas"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00800000&
         Caption         =   "Solo categorías NOSAPP"
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
         TabIndex        =   13
         Top             =   2760
         Width           =   2895
      End
      Begin MSAdodcLib.Adodc data_conv 
         Height          =   375
         Left            =   240
         Top             =   1200
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
         Caption         =   "data_conv"
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
      Begin MSAdodcLib.Adodc data_cli 
         Height          =   375
         Left            =   3360
         Top             =   240
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
         Caption         =   "data_cli"
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
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_infsocmut.frx":0F56
         Left            =   2160
         List            =   "frm_infsocmut.frx":0F81
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2160
         Width           =   3975
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_infsocmut.frx":0FFB
         Left            =   2160
         List            =   "frm_infsocmut.frx":1005
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1560
         Width           =   3975
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_infsocmut.frx":102A
         Left            =   2160
         List            =   "frm_infsocmut.frx":1040
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   3975
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   255
         Left            =   4680
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
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
      Begin MSMask.MaskEdBox mfd 
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
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
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "ZONA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
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
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Supervisor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
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
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Mutualista:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
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
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Rango de Fechas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1920
      Picture         =   "frm_infsocmut.frx":1081
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   2295
   End
End
Attribute VB_Name = "frm_infsocmut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo2.SetFocus
End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub

Private Sub Command1_Click()
Dim Xdesde, Xhasta As String

frm_infsocmut.MousePointer = 11
If mfd.Text <> "__/__/____" Then
   If mfh.Text <> "__/__/____" Then
      Dim MiBaseact As Database
      Dim Unasesact As Workspace
      Set Unasesact = Workspaces(0)
      Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
      MiBaseact.Execute "Delete * from infcli"
      data_inf.RecordSource = "infcli"
      data_inf.Refresh
      If Combo2.Text = "SUPERVISOR GENERAL" Then
         If Combo3.Text = "TODAS" Then
            If Check1.Value = 1 Then
               data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv in ('CCNOS','SMIN','UNIVS','CASANO','GANOS','HEVANO') and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
               data_cli.Refresh
            Else
               data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','SA','EMERN','SAP','UCM','911','911B','CASH') and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
               data_cli.Refresh
            End If
         Else
            If Combo3.Text = "PARQUE DEL PLATA" Then
               data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','SA','EMERN','SAP','UCM') and cl_grupo in (101,102,103,104) and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
               data_cli.Refresh
            Else
               If Combo3.Text = "FLORESTA" Then
                  data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','SA','EMERN','SAP','UCM') and cl_grupo in (201,202,203,204,205,206,207,208,209) and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
                  data_cli.Refresh
               Else
                  If Combo3.Text = "SALINAS" Then
                     data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','SA','EMERN','SAP','UCM') and cl_grupo in (301,302,303,304,305,306,307,308,309,310,312,321) and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
                     data_cli.Refresh
                  Else
                     If Combo3.Text = "ATLANTIDA" Then
                        data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','SA','EMERN','SAP','UCM') and cl_grupo in (400,401,402,403,404,405,406,419,311) and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
                        data_cli.Refresh
                     Else
                        If Combo3.Text = "SOCA" Then
                           data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','SA','EMERN','SAP','UCM') and cl_grupo in (500,501) and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
                           data_cli.Refresh
                        Else
                           If Combo3.Text = "PANDO" Then
                              data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','SA','EMERN','SAP','UCM') and cl_grupo in (600,601,602,603,604,605,606,610,611,612,613) and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
                              data_cli.Refresh
                           Else
                              If Combo3.Text = "BARROS BLS" Then
                                 data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','SA','EMERN','SAP','UCM') and cl_grupo in (620,621,622,623,624,630,631) and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
                                 data_cli.Refresh
                              Else
                                 If Combo3.Text = "SUAREZ" Then
                                    data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','SA','EMERN','SAP','UCM') and cl_grupo in (632,633,634,635,650,800,801,802,803) and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
                                    data_cli.Refresh
                                 Else
                                    If Combo3.Text = "TOLEDO" Then
                                       data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','SA','EMERN','SAP','UCM') and cl_grupo in (810,811) and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
                                       data_cli.Refresh
                                    Else
                                       If Combo3.Text = "SAUCE" Then
                                          data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','SA','EMERN','SAP','UCM') and cl_grupo in (815,816) and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
                                          data_cli.Refresh
                                       Else
                                          If Combo3.Text = "TALA" Then
                                             data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','SA','EMERN','SAP','UCM') and cl_grupo in (670,672,673,674) and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
                                             data_cli.Refresh
                                          Else
                                             If Combo3.Text = "LA TUNA" Then
                                                data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','SA','EMERN','SAP','UCM') and cl_grupo in (700,701,702,703,704,705,706,707,708,709,710,711,712,713,714,715,716,717,718,719,720) and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
                                                data_cli.Refresh
                                             Else
                                                data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_codconv not in ('PART','SA','EMERN','SAP','UCM','911','911B','CASH') and cl_fecing >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fecing <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
                                                data_cli.Refresh
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
               End If
            End If
         End If
         If data_cli.Recordset.RecordCount > 0 Then
            pb1.Max = data_cli.Recordset.RecordCount
            pb1.Value = 0
            data_cli.Recordset.MoveFirst
            DoEvents
            Do While Not data_cli.Recordset.EOF
               If data_cli.Recordset("cl_grupo") <> 670 Then '670 Es san Jacinto
                    data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                    data_conv.Refresh
                    If data_conv.Recordset.RecordCount > 0 Then
                       If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                          data_inf.Recordset.AddNew
                          data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                          data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                          data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                          data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                          data_inf.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
                          data_inf.Recordset("cl_nro_sup") = data_cli.Recordset("cl_nro_sup")
                          data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                          data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                          data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                          data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                          data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                          data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                          data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                          data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                          data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                          data_inf.Recordset("CL_GRUPO") = data_cli.Recordset("CL_GRUPO")
                          If IsNull(data_cli.Recordset("cl_decuota")) = False Then
                             If data_cli.Recordset("cl_decuota") = 1 Then
                                data_inf.Recordset("cl_email") = "Aviso de Carta"
                             Else
                                If data_cli.Recordset("cl_decuota") = 2 Then
                                   data_inf.Recordset("cl_email") = "Se recibe Carta"
                                End If
                             End If
                          End If
                          data_inf.Recordset.Update
                       End If
                    End If
               End If
               data_cli.Recordset.MoveNext
               pb1.Max = data_cli.Recordset.RecordCount
               pb1.Value = pb1.Value + 1
            Loop
            DoEvents
            If Check2.Value = 1 Then
               Xdesde = InputBox("Ingrese fecha desde:")
               Xhasta = InputBox("Ingrese fecha hasta:")
               If Xdesde <> "" And Xhasta <> "" Then
                  If data_inf.Recordset.RecordCount > 0 Then
                     data_inf.Recordset.MoveFirst
                     pb1.Max = pb1.Max + data_inf.Recordset.RecordCount + 1
                     DoEvents
                     Do While Not data_inf.Recordset.EOF
                        Data1.RecordSource = "select * from linmmdd where fecha >=#" & Format(Xdesde, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xhasta, "yyyy/mm/dd") & "# and nro_flia in (1,10,14) and cod_cli =" & data_inf.Recordset("cl_codigo")
                        Data1.Refresh
                        If Data1.Recordset.RecordCount > 0 Then
                           Data1.Recordset.MoveLast
                           data_inf.Recordset.Edit
                           data_inf.Recordset("cl_atrasoa") = Data1.Recordset.RecordCount
                           data_inf.Recordset.Update
                        End If
                        data_inf.Recordset.MoveNext
                        pb1.Value = pb1.Value + 1
                     Loop
                  End If
               End If
               cr1.ReportFileName = App.path & "\infsocmutcons.rpt"
               cr1.ReportTitle = " INFORME AL : " & mfh.Text
               data_inf.RecordSource = "select * from infcli"
               data_inf.Refresh
               cr1.Action = 1
            Else
                If Check1.Value = 1 Then
                   cr1.ReportFileName = App.path & "\infsocmut2.rpt"
                Else
                   cr1.ReportFileName = App.path & "\infsocmut.rpt"
                End If
                cr1.ReportTitle = " INFORME AL : " & mfh.Text
                data_inf.RecordSource = "select * from infcli"
                data_inf.Refresh
                cr1.Action = 1
            End If
         End If
      End If
      If Combo2.Text = "SAN JACINTO" Then
         data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " or estado =" & 0 & " and cl_nrocobr in (5,11,6)"
         data_cli.Refresh
         If data_cli.Recordset.RecordCount > 0 Then
            data_cli.Recordset.MoveFirst
            Do While Not data_cli.Recordset.EOF
               If data_cli.Recordset("cl_nrocobr") = 5 Or _
                  data_cli.Recordset("cl_nrocobr") = 11 Or _
                  data_cli.Recordset("cl_nrocobr") = 6 Then
                    data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                    data_conv.Refresh
                    If data_conv.Recordset.RecordCount > 0 Then
                       If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                          data_inf.Recordset.AddNew
                          data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                          data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                          data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                          data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                          data_inf.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
                          data_inf.Recordset("cl_nro_sup") = data_cli.Recordset("cl_nro_sup")
                          data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                          data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                          data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                          data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                          data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                          data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                          data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                          data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                          data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                          data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                          data_inf.Recordset("CL_GRUPO") = data_cli.Recordset("CL_GRUPO")
                          data_inf.Recordset.Update
                       End If
                    End If
               End If
               data_cli.Recordset.MoveNext
            Loop
            If Check1.Value = 1 Then
               cr1.ReportFileName = App.path & "\infsocmut2.rpt"
            Else
               cr1.ReportFileName = App.path & "\infsocmut.rpt"
            End If
            cr1.ReportTitle = " INFORME AL : " & mfh.Text
            data_inf.RecordSource = "select * from infcli"
            data_inf.Refresh
            cr1.Action = 1
         End If
      End If
   End If
End If

frm_infsocmut.MousePointer = 0
         
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.ConnectionString = "dsn=" & Xconexrmt
'data_cli.RecordSource = "clientes"
'data_cli.Refresh
'data_conv.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_conv.ConnectionString = "dsn=" & Xconexrmt
'data_conv.RecordSource = "convenio"
'data_conv.Refresh
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.RecordSource = "infcli"
'data_inf.Refresh
cr1.ReportFileName = App.path & "\infsocmut.rpt"
Data1.Connect = "odbc;dsn=sappnew;"

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
   Combo1.SetFocus
End If

End Sub
