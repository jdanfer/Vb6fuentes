VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infaran 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de Aranceles"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   Icon            =   "frm_infaran.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6915
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_aran 
      Height          =   330
      Left            =   1200
      Top             =   2760
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "data_aran"
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
      Left            =   5040
      Top             =   2280
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
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_gpos 
      Caption         =   "data_gpos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton b_proc 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      Picture         =   "frm_infaran.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Datos para el informe"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      Begin VB.TextBox t_gpo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         ToolTipText     =   "Unicamente para listado de grupos"
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "frm_infaran.frx":0B14
         Left            =   2040
         List            =   "frm_infaran.frx":0B21
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Grupo (Opcional):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Opción de informe:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   2520
      Picture         =   "frm_infaran.frx":0B6A
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1335
   End
End
Attribute VB_Name = "frm_infaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_proc_Click()
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"

data_inf.RecordSource = "infcli"
data_inf.Refresh

If Combo1.ListIndex = 0 Then
'   data_aran.RecordSource = "Select * from Aran_servicios order by id_gpo,id_serv"
   If t_gpo.Text = "" Then
      data_aran.RecordSource = "select Aran_servicios.id_gpo,Aran_servicios.id_serv,Aran_servicios.desc_serv," & _
      "Aran_servicios.prec_serv,Aran_servicios.por_serv,Aran_grupos.desc_gpo from Aran_servicios " & _
      "inner join Aran_grupos on Aran_servicios.id_gpo=Aran_grupos.id order by Aran_servicios.id_gpo,Aran_servicios.id_serv"
   Else
      data_aran.RecordSource = "select Aran_servicios.id_gpo,Aran_servicios.id_serv,Aran_servicios.desc_serv," & _
      "Aran_servicios.prec_serv,Aran_servicios.por_serv,Aran_grupos.desc_gpo from Aran_servicios " & _
      "inner join Aran_grupos on Aran_servicios.id_gpo=Aran_grupos.id where Aran_servicios.id_gpo=" & t_gpo.Text & " order by Aran_servicios.id_gpo,Aran_servicios.id_serv"
   End If
   data_aran.Refresh
   If data_aran.Recordset.RecordCount > 0 Then
      frm_infaran.MousePointer = 11
      data_aran.Recordset.MoveFirst
      Do While Not data_aran.Recordset.EOF
         data_inf.Recordset.AddNew
         data_inf.Recordset("cl_codigo") = Val(data_aran.Recordset("id_gpo"))
         data_inf.Recordset("cl_cedula") = Val(data_aran.Recordset("id_serv"))
         data_inf.Recordset("cl_direcci") = Mid(data_aran.Recordset("desc_serv"), 1, 80)
         data_inf.Recordset("saldo_cc") = CDbl(data_aran.Recordset("prec_serv"))
         data_inf.Recordset("saldo_cc2") = CDbl(data_aran.Recordset("por_serv"))
         data_inf.Recordset("cl_entre") = Mid(data_aran.Recordset("desc_gpo"), 1, 80)
         data_inf.Recordset.Update
         data_aran.Recordset.MoveNext
      Loop
      frm_infaran.MousePointer = 0
      MsgBox "Proceso terminado"
      cr1.ReportFileName = App.path & "\infgpos.rpt"
      cr1.ReportTitle = "Informe de Grupos de Aranceles con Servicios"
      cr1.Action = 1
      
   End If
End If
If Combo1.ListIndex = 1 Then
   data_aran.RecordSource = "Select * from convenio where cnv_fbaja is null and cnv_aran is not null order by cnv_codigo"
   data_aran.Refresh
   If data_aran.Recordset.RecordCount > 0 Then
      frm_infaran.MousePointer = 11
      data_aran.Recordset.MoveFirst
      Do While Not data_aran.Recordset.EOF
         If Val(data_aran.Recordset("cnv_aran")) <> 0 Then
            data_inf.Recordset.AddNew
            data_inf.Recordset("cl_codigo") = Val(data_aran.Recordset("cnv_aran"))
            data_inf.Recordset("cl_codconv") = data_aran.Recordset("cnv_codigo")
            data_inf.Recordset("cl_nomconv") = Mid(data_aran.Recordset("cnv_desc"), 1, 30)
            data_inf.Recordset("cl_tipcli") = data_aran.Recordset("cnv_colrec")
            data_gpos.RecordSource = "Select * from Aran_grupos where id =" & data_aran.Recordset("cnv_aran")
            data_gpos.Refresh
            If data_gpos.Recordset.RecordCount > 0 Then
               data_inf.Recordset("cl_entre") = Mid(data_gpos.Recordset("desc_gpo"), 1, 80)
            Else
               data_inf.Recordset("cl_entre") = "No encontrado"
            End If
            data_inf.Recordset.Update
         End If
         data_aran.Recordset.MoveNext
      Loop
      frm_infaran.MousePointer = 0
      MsgBox "Proceso terminado"
      cr1.ReportFileName = App.path & "\infgposcnv.rpt"
      cr1.ReportTitle = "Informe de Convenios con Grupo de arancel registrado."
      cr1.Action = 1
      
   End If
End If
If Combo1.ListIndex = 2 Then
   data_aran.RecordSource = "Select * from convenio where cnv_fbaja is null order by cnv_codigo"
   data_aran.Refresh
   If data_aran.Recordset.RecordCount > 0 Then
      frm_infaran.MousePointer = 11
      data_aran.Recordset.MoveFirst
      Do While Not data_aran.Recordset.EOF
         If IsNull(data_aran.Recordset("cnv_aran")) = False Then
            If Val(data_aran.Recordset("cnv_aran")) = 0 Then
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_codigo") = 0
               data_inf.Recordset("cl_codconv") = data_aran.Recordset("cnv_codigo")
               data_inf.Recordset("cl_nomconv") = Mid(data_aran.Recordset("cnv_desc"), 1, 30)
               data_inf.Recordset("cl_tipcli") = data_aran.Recordset("cnv_colrec")
               data_inf.Recordset("cl_entre") = "Sin registrar"
               data_inf.Recordset.Update
            End If
         Else
            data_inf.Recordset.AddNew
            data_inf.Recordset("cl_codigo") = 0
            data_inf.Recordset("cl_codconv") = data_aran.Recordset("cnv_codigo")
            data_inf.Recordset("cl_nomconv") = Mid(data_aran.Recordset("cnv_desc"), 1, 30)
            data_inf.Recordset("cl_tipcli") = data_aran.Recordset("cnv_colrec")
            data_inf.Recordset("cl_entre") = "Sin registrar"
            data_inf.Recordset.Update
         End If
         data_aran.Recordset.MoveNext
      Loop
      frm_infaran.MousePointer = 0
      MsgBox "Proceso terminado"
      cr1.ReportFileName = App.path & "\infgposcnv.rpt"
      cr1.ReportTitle = "Informe de Convenios sin Grupo de arancel registrado."
      cr1.Action = 1
   
   End If
End If

   
End Sub

Private Sub Form_Load()
data_inf.DatabaseName = App.path & "\informes.mdb"
data_gpos.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_aran.ConnectionString = "dsn=" & Xconexrmt


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub
