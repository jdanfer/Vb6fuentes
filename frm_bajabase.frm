VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_bajabase 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proceso de bajas cobrador BASE"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frm_bajabase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6780
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr3 
      Left            =   6000
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport cr2 
      Left            =   6240
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_motiv 
      Caption         =   "data_motiv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   2760
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   6000
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Data data_conv 
      Caption         =   "data_conv"
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
      Top             =   3840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_cob 
      Caption         =   "data_cob"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data data_lin 
      Caption         =   "data_lin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
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
      Top             =   3120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Terminar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   7
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   6
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Datos para dar Baja"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6255
      Begin VB.Data data_clib 
         Caption         =   "data_clib"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   4320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Data data_inflin 
         Caption         =   "data_inflin"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1920
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox txt_cob 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   3360
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin MSMask.MaskEdBox mfec 
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
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
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label2 
         Caption         =   "Nro.Cobrador (0=Todos):"
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
         TabIndex        =   3
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Baja:"
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
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frm_bajabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim xmm, Xaa, Xatra As Long
Dim Xnomem As String
xmm = Month(mfec.Text)
Xaa = Year(mfec.Text)
Xnomem = "EMI" + Mid(mfec.Text, 4, 2) + Mid(mfec.Text, 9, 2)
frm_bajabase.MousePointer = 11
Dim Xdesdef As Date
Xdesdef = Date - 92
If mfec.Text <> "__/__/____" Then
   If txt_cob.Text <> "" Then
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         Do While Not data_inf.Recordset.EOF
            data_inf.Recordset.Delete
            data_inf.Recordset.MoveNext
         Loop
      End If
      If data_inflin.Recordset.RecordCount > 0 Then
         data_inflin.Recordset.MoveFirst
         Do While Not data_inflin.Recordset.EOF
            data_inflin.Recordset.Delete
            data_inflin.Recordset.MoveNext
         Loop
      End If
      data_lin.RecordSource = "Select * from linmmdd where cod_prod =" & 998 & " or cod_prod =" & 999 & " order by cod_cli,ano_paga, mes_paga"
      data_lin.Refresh
      If data_lin.Recordset.RecordCount > 0 Then
         data_lin.Recordset.MoveLast
         pb.Max = data_lin.Recordset.RecordCount + data_lin.Recordset.RecordCount
         pb.Value = 0
         data_lin.Recordset.MoveFirst
         Do While Not data_lin.Recordset.EOF
            data_inflin.Recordset.AddNew
            data_inflin.Recordset("fecha") = data_lin.Recordset("fecha")
            data_inflin.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
            data_inflin.Recordset("mes_paga") = data_lin.Recordset("mes_paga")
            data_inflin.Recordset("ano_paga") = data_lin.Recordset("ano_paga")
            data_inflin.Recordset.Update
            data_lin.Recordset.MoveNext
            pb.Value = pb.Value + 1
         Loop
         DoEvents
         data_inflin.RecordSource = "Select * from infvtas order by cod_cli,ano_paga,mes_paga"
         data_inflin.Refresh
         data_cli.DatabaseName = ""
         data_cli.Connect = "ODBC;DSN=sapp;"
         If txt_cob.Text = 0 Then
            data_cli.RecordSource = "Select * from " & Xnomem & " where nro_cobr =" & 616 & " or nro_cobr =" & 636 & _
            " or nro_cobr =" & 615 & " or nro_cobr =" & 635 & " or nro_cobr =" & 602 & " or nro_cobr =" & 653 & _
            " or nro_cobr =" & 672 & " or nro_cobr =" & 113 & " or nro_cobr =" & 685 & " or nro_cobr =" & 1 & _
            " or nro_cobr =" & 10
            data_cli.Refresh
         Else
            data_cli.RecordSource = "Select * from " & Xnomem & " where nro_cobr =" & txt_cob.Text
            data_cli.Refresh
         End If
         If data_cli.Recordset.RecordCount > 0 Then
            data_cli.Recordset.MoveLast
            pb.Max = pb.Max + data_cli.Recordset.RecordCount + data_cli.Recordset.RecordCount + data_cli.Recordset.RecordCount
            data_cli.Recordset.MoveFirst
            Do While Not data_cli.Recordset.EOF
'                If IsNull(data_cli.Recordset("fecha_baja")) = False Then
'                   data_cli.Recordset.MoveNext
'                Else
                    data_inf.Recordset.AddNew
                    data_inf.Recordset("cl_codigo") = data_cli.Recordset("cliente")
                    data_inf.Recordset("cl_apellid") = data_cli.Recordset("apellidos")
                    data_inf.Recordset("cl_direcci") = data_cli.Recordset("dir_cli")
                    data_inf.Recordset("cl_telefon") = data_cli.Recordset("tel_cli")
                    data_inf.Recordset("cl_fecing") = data_cli.Recordset("fecha_ing")
                    data_inf.Recordset("cl_fnac") = data_cli.Recordset("fecha_nac")
                    data_inf.Recordset("cl_grupo") = data_cli.Recordset("grupo")
                    data_inf.Recordset("cl_localid") = Mid(data_cli.Recordset("loc_cli"), 1, 35)
                    data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("nro_cobr")
                    data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("nom_cobr")
'                    data_conv.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cod_cnv") & "'"
                    data_clib.RecordSource = "Select * from clientes where cl_codigo =" & data_cli.Recordset("cliente")
                    data_clib.Refresh
                    If data_cli.Recordset.RecordCount > 0 Then
                       data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_clib.Recordset("cl_codconv") & "'"
                       data_conv.Refresh
                       If data_conv.Recordset.RecordCount > 0 Then
                          If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
                             If data_conv.Recordset("cnv_grupo") <> "" Then
                                data_inf.Recordset("cl_descpag") = "MUTUAL"
                             Else
                                If IsNull(data_conv.Recordset("cnv_emite")) = False Then
                                   If data_conv.Recordset("cnv_emite") = "NO" Then
                                      data_inf.Recordset("cl_descpag") = "MUTUAL"
                                   End If
                                Else
                                   data_inf.Recordset("cl_descpag") = "MUTUAL"
                                End If
                             End If
                          Else
                             If IsNull(data_clib.Recordset("estado")) = False Then
                                If data_clib.Recordset("estado") > 1 Then
                                   data_inf.Recordset("cl_descpag") = "MUTUAL"
                                End If
                             End If
                          End If
                       End If
                    End If
                    data_inf.Recordset("cl_codconv") = data_cli.Recordset("cod_cnv")
                    data_inf.Recordset("cl_nomconv") = data_cli.Recordset("nom_cnv")
                    data_inf.Recordset("cl_nrovend") = data_cli.Recordset("nro_vende")
                    data_inf.Recordset("cl_nomvend") = data_cli.Recordset("nom_vende")
                    data_inf.Recordset.Update
                    data_cli.Recordset.MoveNext
'                End If
                 pb.Value = pb.Value + 1
            Loop
            '60014
' Oscar Pereira

         End If
         DoEvents
         data_inf.Recordset.MoveFirst
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_descpag")) = False Then
               If data_inf.Recordset("cl_descpag") = "MUTUAL" Then
                  data_inf.Recordset.Delete
               End If
            End If
            data_inf.Recordset.MoveNext
         Loop
         Dim xxx As Integer
         data_inflin.Recordset.MoveFirst
         Do While Not data_inflin.Recordset.EOF
            If IsNull(data_inflin.Recordset("cod_cli")) = False Then
               data_inf.Recordset.FindFirst "cl_codigo =" & data_inflin.Recordset("cod_cli")
               If Not data_inf.Recordset.NoMatch Then
                  If IsNull(data_inf.Recordset("cl_ultmesp")) = False Then
                     If data_inflin.Recordset("ano_paga") > data_inf.Recordset("cl_ultanop") Then
                        data_inf.Recordset.Edit
                        data_inf.Recordset("cl_ultmesp") = data_inflin.Recordset("mes_paga")
                        data_inf.Recordset("cl_ultanop") = data_inflin.Recordset("ano_paga")
                        data_inf.Recordset.Update
                     Else
                        If data_inflin.Recordset("ano_paga") = data_inf.Recordset("cl_ultanop") Then
                           If data_inflin.Recordset("mes_paga") > data_inf.Recordset("cl_ultmesp") Then
                              data_inf.Recordset.Edit
                              data_inf.Recordset("cl_ultmesp") = data_inflin.Recordset("mes_paga")
                              data_inf.Recordset("cl_ultanop") = data_inflin.Recordset("ano_paga")
                              data_inf.Recordset.Update
                           End If
                        End If
                     End If
                  Else
                     data_inf.Recordset.Edit
                     data_inf.Recordset("cl_ultmesp") = data_inflin.Recordset("mes_paga")
                     data_inf.Recordset("cl_ultanop") = data_inflin.Recordset("ano_paga")
                     data_inf.Recordset.Update
                  End If
                  data_inflin.Recordset.MoveNext
               Else
                  data_inflin.Recordset.MoveNext
               End If
            Else
               data_inflin.Recordset.MoveNext
            End If
            pb.Value = pb.Value + 1
         Loop
         data_inf.Recordset.MoveFirst
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_fecing")) = True Then
               data_inf.Recordset.Edit
               data_inf.Recordset("cl_fecing") = CDate("01/01/2000")
               data_inf.Recordset.Update
            End If
            If IsNull(data_inf.Recordset("cl_ultmesp")) = True Then
               data_inf.Recordset.Edit
               data_inf.Recordset("cl_ultmesp") = Month(data_inf.Recordset("cl_fecing"))
               data_inf.Recordset("cl_ultanop") = Year(data_inf.Recordset("cl_fecing"))
               data_inf.Recordset.Update
            End If
            data_inf.Recordset.MoveNext
            pb.Value = pb.Value + 1
         Loop
         If data_inf.Recordset.RecordCount > 0 Then
            data_inf.Recordset.MoveFirst
         End If
         Do While Not data_inf.Recordset.EOF
            If data_inf.Recordset("cl_ultanop") = Xaa Then
               Xatra = xmm - data_inf.Recordset("cl_ultmesp")
               If Xatra < 0 Then
                  Xatra = 0
               End If
            Else
               Xatra = xmm - data_inf.Recordset("cl_ultmesp") + 12
            End If
            data_inf.Recordset.Edit
            data_inf.Recordset("cl_atrasoa") = Xatra
            data_inf.Recordset.Update
            If Xatra <= 4 Then
               data_inf.Recordset.Delete
            End If
            
            data_inf.Recordset.MoveNext
            pb.Value = pb.Value + 1
         Loop
         
         If data_inf.Recordset.RecordCount > 0 Then
            data_inf.Recordset.MoveFirst
            pb.Max = pb.Max + data_inf.Recordset.RecordCount + data_inf.Recordset.RecordCount
         End If
         data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
'         data_cli.RecordSource = "clientes"
'         data_cli.Refresh
         data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xdesdef, "yyyy/mm/dd") & "# And cod_prod <>" & 998 & " or cod_prod <>" & 999 & " order by cod_cli"
         data_lin.Refresh
         If data_inf.Recordset.RecordCount > 0 Then
            data_inf.Recordset.MoveFirst
            Do While Not data_inf.Recordset.EOF
               data_lin.Recordset.FindFirst "cod_cli =" & data_inf.Recordset("cl_codigo")
               If Not data_lin.Recordset.NoMatch Then
                  data_inf.Recordset.Edit
                  data_inf.Recordset("cl_nombre") = "SERVICIO el "
                  data_inf.Recordset("fecha_modi") = data_lin.Recordset("fecha")
                  data_inf.Recordset.Update
               End If
               data_inf.Recordset.MoveNext
               pb.Value = pb.Value + 1
            Loop
         End If
                  
         frm_bajabase.MousePointer = 0

         cr1.ReportTitle = "Informe de socios cobrador de BASE pasados a BAJA con fecha " & mfec.Text
         cr1.Action = 1
      
         cr2.ReportTitle = "Informe de socios cobrador de BASE pasados a BAJA con Servicios " & mfec.Text
         cr2.Action = 1
      
         cr3.ReportTitle = "Informe de socios cobrador de BASE pasados a BAJA con COMPLEMENTO " & mfec.Text
         cr3.Action = 1
                  
         Dim Xquevaahacer As String
         Xquevaahacer = ""
         Xquevaahacer = MsgBox("DESEA PROCESAR LAS BAJAS?", vbExclamation + vbYesNo, "BAJAS")
         If Xquevaahacer = vbYes Then
            If data_inf.Recordset.RecordCount > 0 Then
               data_inf.Recordset.MoveFirst
               Do While Not data_inf.Recordset.EOF
                  data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_inf.Recordset("cl_codigo")
                  data_cli.Refresh
                  If data_cli.Recordset.RecordCount > 0 Then
                     If IsNull(data_cli.Recordset("estado")) = False Then
                        If data_cli.Recordset("estado") = 2 Then
                        Else
                           data_cli.Recordset.Edit
                           data_cli.Recordset("fecha_baja") = mfec.Text
                           data_cli.Recordset("estado") = 2
                           data_cli.Recordset("cl_dircobr") = "FALTA de PAGO BAJA AUT"
                           data_cli.Recordset.Update
                           data_motiv.Recordset.AddNew
                           data_motiv.Recordset("usuario") = WElusuario
                           data_motiv.Recordset("fecha") = mfec.Text
                           data_motiv.Recordset("hora") = Format(Time, "HH:mm")
                           data_motiv.Recordset("cl_codigo") = data_inf.Recordset("cl_codigo")
                           data_motiv.Recordset("desc") = "BAJA"
                           data_motiv.Recordset("cl_motivo") = "DESINTERES"
                           data_motiv.Recordset("convenio") = data_inf.Recordset("cl_codconv")
                           data_motiv.Recordset.Update
                         End If
                     Else
                         data_cli.Recordset.Edit
                         data_cli.Recordset("fecha_baja") = mfec.Text
                         data_cli.Recordset("estado") = 2
                         data_cli.Recordset("cl_dircobr") = "FALTA de PAGO BAJA AUT"
                         data_cli.Recordset.Update
                         data_motiv.Recordset.AddNew
                         data_motiv.Recordset("usuario") = WElusuario
                         data_motiv.Recordset("fecha") = mfec.Text
                         data_motiv.Recordset("hora") = Format(Time, "HH:mm")
                         data_motiv.Recordset("cl_codigo") = data_inf.Recordset("cl_codigo")
                         data_motiv.Recordset("desc") = "BAJA"
                         data_motiv.Recordset("cl_motivo") = "DESINTERES"
                         data_motiv.Recordset("convenio") = data_inf.Recordset("cl_codconv")
                         data_motiv.Recordset.Update
                     End If
                  End If
                  data_inf.Recordset.MoveNext
                  pb.Value = pb.Value + 1
               Loop
            End If
         End If
         frm_bajabase.MousePointer = 0
         If Xquevaahacer = vbYes Then
            MsgBox "Proceso terminado, favor enviar copia de informes por correo"
         Else
            MsgBox "Proceso terminao, NO SE GRABARON LAS BAJAS"
         End If
      End If
   
   End If
End If

frm_bajabase.MousePointer = 0

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
MsgBox "Terminado"

End Sub

Private Sub Form_Load()
data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_cli.RecordSource = "clientes"
'data_cli.Refresh
data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_lin.RecordSource = "linmmdd"
'data_lin.Refresh
data_inf.DatabaseName = App.path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh
data_cob.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cob.RecordSource = "cobrador"
data_cob.Refresh
data_conv.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_conv.RecordSource = "convenio"
data_conv.Refresh
data_inflin.DatabaseName = App.path & "\informes.mdb"
data_inflin.RecordSource = "infvtas"
data_inflin.Refresh
data_motiv.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_motiv.RecordSource = "abmsocio"
'data_motiv.Refresh
cr1.ReportFileName = App.path & "\infbajabase.rpt"
cr2.ReportFileName = App.path & "\infbajacons.rpt"
cr3.ReportFileName = App.path & "\infbajabasem.rpt"
data_clib.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub mfec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_cob.SetFocus
End If

End Sub

Private Sub txt_cob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub

Private Sub txt_cob_LostFocus()
If txt_cob.Text <> "" Then
   If txt_cob.Text = 999 Then
      Label3.Caption = "TODOS"
   Else
      data_cob.Recordset.FindFirst "cb_numero =" & txt_cob.Text
      If Not data_cob.Recordset.NoMatch Then
         Label3.Caption = data_cob.Recordset("cb_nombre")
      Else
         MsgBox "No se encuentra cobrador"
         Label3.Caption = "No Encontrado"
      End If
   End If
End If


End Sub
