VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infatsoc 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes atención al socio"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6480
   Icon            =   "frm_infatsoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   1800
      Top             =   4200
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
   Begin Crystal.CrystalReport cr1 
      Left            =   4080
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
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
      Left            =   5400
      Picture         =   "frm_infatsoc.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command1 
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
      Left            =   360
      Picture         =   "frm_infatsoc.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Procesar"
      Top             =   4080
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tipo de informe"
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   5895
      Begin VB.OptionButton Option4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Resumen"
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
         Left            =   3240
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Detalle"
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
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Datos para informe"
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_infatsoc.frx":0F56
         Left            =   2160
         List            =   "frm_infatsoc.frx":0F66
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
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
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000000FF&
         Caption         =   "Ordenar por conformidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   3495
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
         ItemData        =   "frm_infatsoc.frx":0F92
         Left            =   2160
         List            =   "frm_infatsoc.frx":0FA8
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   3375
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
         BackColor       =   &H0080C0FF&
         Caption         =   "Opción:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Selección:"
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
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
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
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   120
      Picture         =   "frm_infatsoc.frx":0FE8
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "frm_infatsoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"
data_inf.RecordSource = "infcli"
data_inf.Refresh

If mfd.Text <> "__/__/____" And mfh.Text <> "__/__/____" Then
   
'   Data1.DatabaseName = ""
'   Data1.Connect = "ODBC;DSN=sappat;"
'    data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
   If Combo1.Text = "TODO" Then
      If Combo2.Text = "Terminado" Then
         Data1.RecordSource = "Select * from ingresosat where at_fecfin >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and at_fecfin <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by at_fecfin"
      Else
         Data1.RecordSource = "Select * from ingresosat where at_fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and at_fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by at_fecha"
      End If
      Data1.Refresh
   Else
      If Combo2.Text = "Terminado" Then
         Data1.RecordSource = "Select * from ingresosat where at_fecfin >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and at_fecfin <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and at_descat ='" & Combo1.Text & "' order by at_fecfin"
      Else
         Data1.RecordSource = "Select * from ingresosat where at_fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and at_fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and at_descat ='" & Combo1.Text & "' order by at_fecha"
      End If
      Data1.Refresh
   End If
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         If Check1.Value <> 1 Then
            data_inf.Recordset.AddNew
            data_inf.Recordset("cl_codigo") = Data1.Recordset("at_cliente")
            If IsNull(Data1.Recordset("at_nomb")) = False Then
               data_inf.Recordset("cl_apellid") = Mid(Data1.Recordset("at_nomb"), 1, 60)
            End If
            data_inf.Recordset("cl_codconv") = Data1.Recordset("at_codconv")
            If IsNull(Data1.Recordset("at_nomconv")) = False Then
               data_inf.Recordset("cl_nomconv") = Mid(Data1.Recordset("at_nomconv"), 1, 30)
            End If
            data_inf.Recordset("cl_cedula") = Data1.Recordset("at_ced")
            data_inf.Recordset("cl_codced") = Data1.Recordset("at_codced")
            data_inf.Recordset("cl_fultpag") = Data1.Recordset("at_fecha")
            data_inf.Recordset("cl_dpto") = Data1.Recordset("at_hora")
            data_inf.Recordset("cl_nomvend") = Data1.Recordset("at_usuario")
            data_inf.Recordset("cl_nrovend") = Data1.Recordset("at_categ")
            data_inf.Recordset("cl_descpag") = Data1.Recordset("at_descat")
            data_inf.Recordset("cl_nombre") = Data1.Recordset("at_moti")
            If IsNull(Data1.Recordset("at_detal")) = False Then
               data_inf.Recordset("info_debit") = Mid(Data1.Recordset("at_detal"), 1, 130)
            End If
            data_inf.Recordset("cl_forpago") = Data1.Recordset("at_estado")
            If Data1.Recordset("at_estado") = 0 Then
               data_inf.Recordset("cl_localid") = "EN PROCESO"
            Else
               If Data1.Recordset("at_estado") = 1 Then
                  data_inf.Recordset("cl_localid") = "TERMINADO"
               Else
                  If Data1.Recordset("at_estado") = 2 Then
                     data_inf.Recordset("cl_localid") = "CANCELADO"
                  Else
                     data_inf.Recordset("cl_localid") = "EN PROCESO"
                  End If
               End If
            End If
            data_inf.Recordset("cl_fultvta") = Data1.Recordset("at_fecfin")
            data_inf.Recordset("cl_atrasoa") = Data1.Recordset("at_confor")
            data_inf.Recordset.Update
         Else
            If IsNull(Data1.Recordset("at_estado")) = False Then
               'If Data1.Recordset("at_estado") = 1 Then
                    data_inf.Recordset.AddNew
                    data_inf.Recordset("cl_codigo") = Data1.Recordset("at_cliente")
                    If IsNull(Data1.Recordset("at_nomb")) = False Then
                       data_inf.Recordset("cl_apellid") = Mid(Data1.Recordset("at_nomb"), 1, 60)
                    End If
                    data_inf.Recordset("cl_codconv") = Data1.Recordset("at_codconv")
                    If IsNull(Data1.Recordset("at_nomconv")) = False Then
                       data_inf.Recordset("cl_nomconv") = Mid(Data1.Recordset("at_nomconv"), 1, 30)
                    End If
                    data_inf.Recordset("cl_cedula") = Data1.Recordset("at_ced")
                    data_inf.Recordset("cl_codced") = Data1.Recordset("at_codced")
                    data_inf.Recordset("cl_fultpag") = Data1.Recordset("at_fecha")
                    data_inf.Recordset("cl_dpto") = Data1.Recordset("at_hora")
                    data_inf.Recordset("cl_nomvend") = Data1.Recordset("at_usuario")
                    data_inf.Recordset("cl_nrovend") = Data1.Recordset("at_categ")
                    data_inf.Recordset("cl_descpag") = Data1.Recordset("at_descat")
                    If IsNull(Data1.Recordset("at_detal")) = False Then
                       data_inf.Recordset("info_debit") = Mid(Data1.Recordset("at_detal"), 1, 130)
                    End If
                    data_inf.Recordset("cl_forpago") = Data1.Recordset("at_estado")
                    If Data1.Recordset("at_estado") = 0 Then
                       data_inf.Recordset("cl_localid") = "EN PROCESO"
                    Else
                       If Data1.Recordset("at_estado") = 1 Then
                          data_inf.Recordset("cl_localid") = "TERMINADO"
                       Else
                          If Data1.Recordset("at_estado") = 2 Then
                             data_inf.Recordset("cl_localid") = "CANCELADO"
                          Else
                             data_inf.Recordset("cl_localid") = "EN PROCESO"
                          End If
                       End If
                    End If
                    data_inf.Recordset("cl_fultvta") = Data1.Recordset("at_fecfin")
                    data_inf.Recordset("cl_atrasoa") = Data1.Recordset("at_confor")
                    If IsNull(Data1.Recordset("at_confor")) = False Then
                       If Data1.Recordset("at_confor") = 0 Then
                          data_inf.Recordset("cl_nom_sup") = "CONFORME"
                       Else
                          If Data1.Recordset("at_confor") = 1 Then
                             data_inf.Recordset("cl_nom_sup") = "NO CONFORME"
                          Else
                             data_inf.Recordset("cl_nom_sup") = "SIN DATOS"
                          End If
                       End If
                    Else
                       data_inf.Recordset("cl_nom_sup") = "SIN DATOS"
                    End If
                    data_inf.Recordset("cl_nombre") = Data1.Recordset("at_moti")
                    
                    data_inf.Recordset.Update
               'End If
            End If
         End If
         
         Data1.Recordset.MoveNext
      Loop
      MsgBox "Proceso terminado"
      If Combo2.Text = "En proceso" Then 'En proceso
         MiBaseact.Execute "Delete * from infcli where cl_forpago not in (0)"
      End If
      If Combo2.Text = "Terminado" Then
         MiBaseact.Execute "Delete * from infcli where cl_localid not in ('TERMINADO')"
      End If
      If Combo2.Text = "Cancelado" Then
         MiBaseact.Execute "Delete * from infcli where cl_localid not in ('CANCELADO')"
      End If
      
      data_inf.RecordSource = "Select * from infcli"
      data_inf.Refresh
      If Check1.Value = 1 Then
         If Option4.Value = True Then
            cr1.ReportFileName = App.path & "\infatsoc2n.rpt"
         Else
            cr1.ReportFileName = App.path & "\infatsoc2.rpt"
         End If
         If Combo2.Text = "En proceso" Then
            cr1.ReportTitle = "INFORME DE REGISTROS EN PROCESO POR CONFORMIDAD DESDE: " & mfd.Text & " HASTA: " & mfh.Text
         Else
            If Combo2.Text = "Terminado" Then
               cr1.ReportTitle = "INFORME DE REGISTROS CERRADOS POR CONFORMIDAD DESDE: " & mfd.Text & " HASTA: " & mfh.Text
            Else
               If Combo2.Text = "Cancelado" Then
                  cr1.ReportTitle = "INFORME DE REGISTROS CANCELADOS POR CONFORMIDAD DESDE: " & mfd.Text & " HASTA: " & mfh.Text
               Else
                  cr1.ReportTitle = "INFORME DE TODOS los REGISTROS POR CONFORMIDAD DESDE: " & mfd.Text & " HASTA: " & mfh.Text
               End If
            End If
         End If
      Else
         cr1.ReportFileName = App.path & "\infatsoc.rpt"
         If Combo2.Text = "En proceso" Then
            cr1.ReportTitle = "INFORME DE REGISTROS EN PROCESO INGRESADOS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
         Else
            If Combo2.Text = "Terminado" Then
               cr1.ReportTitle = "INFORME DE REGISTROS TERMINADOS INGRESADOS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
            Else
               If Combo2.Text = "Cancelado" Then
                  cr1.ReportTitle = "INFORME DE REGISTROS CANCELADOS INGRESADOS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
               Else
                  cr1.ReportTitle = "INFORME DE TODOS los REGISTROS INGRESADOS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
               End If
            End If
         End If
      End If
      cr1.Action = 1
   Else
      MsgBox "No existen registros"
      
   End If
End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh
Data1.ConnectionString = "dsn=" & Xconexrmt

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
