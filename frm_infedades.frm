VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infedades 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes por edad"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6300
   Icon            =   "frm_infedades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6300
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc data_cli 
      Height          =   375
      Left            =   2880
      Top             =   3240
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
   Begin Crystal.CrystalReport cr1 
      Left            =   2400
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   495
      Left            =   1920
      Picture         =   "frm_infedades.frx":0742
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Procesar"
      Height          =   495
      Left            =   360
      Picture         =   "frm_infedades.frx":0CCC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      Caption         =   "Datos para el informe"
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      Begin VB.TextBox t_mesh 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox t_mesd 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3600
         TabIndex        =   12
         Top             =   1080
         Width           =   495
      End
      Begin MSAdodcLib.Adodc data_conv 
         Height          =   375
         Left            =   2040
         Top             =   2160
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
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
      Begin VB.ComboBox cbocat 
         Height          =   315
         ItemData        =   "frm_infedades.frx":1256
         Left            =   1920
         List            =   "frm_infedades.frx":126F
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox t_h 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox t_d 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1080
         Width           =   495
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "MESES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "AÑOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Categoría:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rango de edad:"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rango de fechas:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   120
      Picture         =   "frm_infedades.frx":12B9
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   2415
   End
End
Attribute VB_Name = "frm_infedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xfecnacdesde, Xfecnachasta As Date
Dim Xdiasdesde, Xdiashasta As Long
Dim Xmes1, Xmes2, Xmes3 As Integer
Xmes1 = Month(md.Text)

If Xmes1 = 1 Then
   Xmes2 = 2
End If
If Xmes1 = 2 Then
   Xmes2 = 3
End If
If Xmes1 = 3 Then
   Xmes2 = 4
End If
If Xmes1 = 4 Then
   Xmes2 = 5
End If
If Xmes1 = 5 Then
   Xmes2 = 6
End If
If Xmes1 = 6 Then
   Xmes2 = 7
End If
If Xmes1 = 7 Then
   Xmes2 = 8
End If
If Xmes1 = 8 Then
   Xmes2 = 9
End If
If Xmes1 = 9 Then
   Xmes2 = 10
End If
If Xmes1 = 10 Then
   Xmes2 = 11
End If
If Xmes1 = 11 Then
   Xmes2 = 12
End If
If Xmes1 = 12 Then
   Xmes2 = 1
End If

Xmes3 = Month(mh.Text)
frm_infedades.MousePointer = 11

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"
data_inf.RecordSource = "infcli"
data_inf.Refresh

If md.Text <> "__/__/____" And mh.Text <> "__/__/____" Then
   If t_d.Text <> "" And t_h.Text <> "" Then
      If t_d.Text = 0 Then
         Xfecnacdesde = "01/01/" & Trim(str(Year(Date)))
         Xdiasdesde = DateDiff("d", CDate(Xfecnacdesde), Date)
         Xdiasdesde = t_d.Text * 365
         Xfecnacdesde = CDate(mh.Text) - Xdiasdesde
         
'         Xfecnacdesde = Date - Xdiasdesde
         Xdiashasta = t_h.Text * 365
         Xfecnachasta = CDate(md.Text) - Xdiashasta
'         Xfecnachasta = Date - Xdiashasta
      Else
         Xdiasdesde = t_d.Text * 365
         Xfecnacdesde = CDate(mh.Text) - Xdiasdesde
         Xdiashasta = t_h.Text * 365
         Xfecnachasta = CDate(md.Text) - Xdiashasta
      End If
'      data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
      data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_fnac is not null and " & _
      "cl_codconv not in ('PART','UDEMM','CERSEM','911','911B','CASH','CCASMU','SEMM1','SEMM2','UCM','SEMM')"
'      data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_fnac >=#" & Format(Xfecnachasta, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(Xfecnacdesde, "yyyy/mm/dd") & "#"
      data_cli.Refresh
   Else
      If t_mesd.Text <> "" Then
         Xdiasdesde = t_mesd.Text * 30
         Xfecnacdesde = CDate(mh.Text) - Xdiasdesde
         Xdiashasta = t_mesh.Text * 30
         Xfecnachasta = CDate(md.Text) - Xdiashasta
      End If
'      data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
      data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_fnac is not null and " & _
      "cl_codconv not in ('PART','UDEMM','CERSEM','911','911B','CASH','CCASMU','SEMM1','SEMM2','UCM','SEMM')"
'      data_cli.RecordSource = "Select * from clientes where estado =" & 1 & " and cl_fnac >=#" & Format(Xfecnachasta, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(Xfecnacdesde, "yyyy/mm/dd") & "#"
      data_cli.Refresh
   
   End If
   If data_cli.Recordset.RecordCount > 0 Then
      data_cli.Recordset.MoveFirst
      Do While Not data_cli.Recordset.EOF
         If t_mesd.Text <> "" Then
            If Format(data_cli.Recordset("cl_fnac"), "yyyy/mm/dd") >= Format(Xfecnachasta, "yyyy/mm/dd") And Format(data_cli.Recordset("cl_fnac"), "yyyy/mm/dd") <= Format(Xfecnacdesde, "yyyy/mm/dd") Then
               If t_mesd.Text > 11 Then
                  'If Month(data_cli.Recordset("cl_fnac")) >= Month(md.Text) Then
                  '   If Month(data_cli.Recordset("cl_fnac")) <= Month(mh.Text) Then
                        data_inf.Recordset.AddNew
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                        data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                        data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                        data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                        data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                        data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                        data_inf.Recordset.Update
                   '  End If
                  'End If
               Else
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                  data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                  data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                  data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                  data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                  data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                  data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                  data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                  data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                  data_inf.Recordset.Update
               End If
            End If
         Else
            If Format(data_cli.Recordset("cl_fnac"), "yyyy/mm/dd") >= Format(Xfecnachasta, "yyyy/mm/dd") And Format(data_cli.Recordset("cl_fnac"), "yyyy/mm/dd") <= Format(Xfecnacdesde, "yyyy/mm/dd") Then
               If Month(data_cli.Recordset("cl_fnac")) >= Month(md.Text) Then
                  If Month(data_cli.Recordset("cl_fnac")) <= Month(mh.Text) Then
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                     data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                     data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_inf.Recordset.Update
                   End If
                End If
            End If
         End If
         data_cli.Recordset.MoveNext
      Loop
      If cbocat.ListIndex > 0 Then
         If cbocat.ListIndex = 1 Then
            data_inf.Refresh
            If data_inf.Recordset.RecordCount > 0 Then
               data_inf.Recordset.MoveFirst
               Do While Not data_inf.Recordset.EOF
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_inf.Recordset("cl_codconv") & "'"
                  data_conv.Refresh
                  If data_conv.Recordset.RecordCount > 0 Then
                     If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
                        If Trim(data_conv.Recordset("cnv_grupo")) = "CCOU" Or _
                           data_conv.Recordset("cnv_grupo") = "SMI" Or _
                           data_conv.Recordset("cnv_grupo") = "IMPASA" Or _
                           data_conv.Recordset("cnv_grupo") = "H.EVANGELICO" Or _
                           data_conv.Recordset("cnv_grupo") = "UNIVERSAL" Or _
                           data_conv.Recordset("cnv_grupo") = "CASA DE GALICIA" Then
                        Else
                           data_inf.Recordset.Delete
                        End If
                     Else
                        data_inf.Recordset.Delete
                     End If
                  End If
                  data_inf.Recordset.MoveNext
               Loop
            End If
         End If
         If cbocat.ListIndex > 1 Then
            data_inf.Refresh
            If data_inf.Recordset.RecordCount > 0 Then
               data_inf.Recordset.MoveFirst
               Do While Not data_inf.Recordset.EOF
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_inf.Recordset("cl_codconv") & "'"
                  data_conv.Refresh
                  If data_conv.Recordset.RecordCount > 0 Then
                     If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
                        If Trim(data_conv.Recordset("cnv_grupo")) = cbocat.Text Then
                        Else
                           data_inf.Recordset.Delete
                        End If
                     Else
                        data_inf.Recordset.Delete
                     End If
                  Else
                     data_inf.Recordset.Delete
                  End If
                  data_inf.Recordset.MoveNext
               Loop
            End If
         End If
      End If
      frm_infedades.MousePointer = 0
      MsgBox "Proceso terminado"
      data_inf.RecordSource = "Select * from infcli order by cl_fnac"
      data_inf.Refresh
      cr1.ReportFileName = App.path & "\infedades.rpt"
      cr1.ReportTitle = "Informe de socios por rango de edad período desde " & md.Text & " hasta " & mh.Text
      cr1.Action = 1
      
   End If
   frm_infedades.MousePointer = 0
End If
frm_infedades.MousePointer = 0

   
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.ConnectionString = "dsn=" & Xconexrmt
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
'data_inf.RecordSource = "infcli"
'data_inf.Refresh
'data_conv.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_conv.ConnectionString = "dsn=" & Xconexrmt


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_d.SetFocus
End If

End Sub

Private Sub t_d_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_h.SetFocus
End If

End Sub

Private Sub t_h_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_mesd.SetFocus
End If

End Sub

Private Sub t_mesd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_mesh.SetFocus
End If

End Sub
