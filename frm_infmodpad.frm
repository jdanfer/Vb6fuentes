VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infmodpad 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes solicitud de modificaciones al padrón"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6285
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infmodpad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6285
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_lla 
      Height          =   375
      Left            =   240
      Top             =   2400
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
      Caption         =   "data_lla"
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
   Begin MSAdodcLib.Adodc data_cli 
      Height          =   375
      Left            =   3720
      Top             =   2280
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
   Begin VB.OptionButton Option3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Modificaciones desde despacho"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   5295
   End
   Begin Crystal.CrystalReport cr11 
      Left            =   3120
      Top             =   3360
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
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4560
      Picture         =   "frm_infmodpad.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      Picture         =   "frm_infmodpad.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Procesar"
      Top             =   3120
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Socios en otros convenios"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   5295
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Afiliaciones SAPP registradas en base."
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Value           =   -1  'True
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   6240
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   1680
      Picture         =   "frm_infmodpad.frx":0F56
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1935
   End
End
Attribute VB_Name = "frm_infmodpad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = False

frm_infmodpad.MousePointer = 11
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"

data_inf.RecordSource = "infcli"
data_inf.Refresh

If Option1.Value = True Then
   data_cli.RecordSource = "Select * from clientes where cl_codconv ='" & "SOLAF" & "' or cl_codconv ='" & "SOLAMB" & "' or cl_codconv ='" & "SOLEME" & "' or cl_codconv ='" & "SOLPAR" & "' and cl_tipocli =" & 1 & " and estado =" & 1
   data_cli.Refresh
End If
If Option2.Value = True Then
   data_cli.RecordSource = "Select * from clientes where cl_fultvta >='" & Format("01/01/2017", "yyyy-mm-dd") & "' and cl_tipocli =" & 1
   data_cli.Refresh
End If
If Option3.Value = True Then
   data_cli.RecordSource = "Select * from clientes where cl_fultvta >='" & Format("01/01/2018", "yyyy-mm-dd") & "' and cl_tipocli =" & 2
   data_cli.Refresh
End If

If data_cli.Recordset.RecordCount > 0 Then
   data_cli.Recordset.MoveFirst
   Do While Not data_cli.Recordset.EOF
      If Option2.Value = True Then
         If IsNull(data_cli.Recordset("cl_fultvta")) = False Then
            If data_cli.Recordset("estado") = 2 Then
               data_lla.RecordSource = "select * from linmmdd where cod_cli =" & data_cli.Recordset("cl_codigo") & " and fecha >='" & Format(data_cli.Recordset("cl_fultvta"), "yyyy-mm-dd") & "'"
               data_lla.Refresh
               If data_lla.Recordset.RecordCount > 0 Then
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_fultvta") = data_cli.Recordset("cl_fultvta")
                  data_inf.Recordset("cl_tipocli") = data_cli.Recordset("cl_tipocli")
                  data_inf.Recordset("cl_celular") = data_cli.Recordset("cl_celular")
                  data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                  data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                  data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                  data_inf.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
                  data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                  data_inf.Recordset.Update
               End If
            Else
                data_inf.Recordset.AddNew
                data_inf.Recordset("cl_fultvta") = data_cli.Recordset("cl_fultvta")
                data_inf.Recordset("cl_tipocli") = data_cli.Recordset("cl_tipocli")
                data_inf.Recordset("cl_celular") = data_cli.Recordset("cl_celular")
                data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                data_inf.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
                data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                data_inf.Recordset.Update
            End If
            
         End If
      Else
         If Option3.Value = True Then
            If IsNull(data_cli.Recordset("cl_fultvta")) = False Then
               data_lla.RecordSource = "Select * from llamado where fecha ='" & Format(data_cli.Recordset("cl_fultvta"), "yyyy-mm-dd") & "' and matric =" & data_cli.Recordset("cl_codigo")
               data_lla.Refresh
               If data_lla.Recordset.RecordCount > 0 Then
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_fultvta") = data_cli.Recordset("cl_fultvta")
                  data_inf.Recordset("cl_tipocli") = data_cli.Recordset("cl_tipocli")
                  data_inf.Recordset("cl_celular") = data_cli.Recordset("cl_celular")
                  data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                  data_inf.Recordset("cl_apellid") = Mid(data_lla.Recordset("nombre"), 1, 60)
                  data_inf.Recordset("cl_codconv") = data_lla.Recordset("categ")
                  data_inf.Recordset("cl_direcci") = Mid(data_lla.Recordset("direcc"), 1, 80)
                  data_inf.Recordset("cl_telefon") = Mid(data_lla.Recordset("telef"), 1, 20)
                  data_inf.Recordset("cl_cedula") = data_lla.Recordset("ci")
                  data_inf.Recordset("cl_nomconv") = Mid(data_lla.Recordset("nomcat"), 1, 30)
                  data_inf.Recordset.Update
               End If
            End If
         Else
            data_inf.Recordset.AddNew
            data_inf.Recordset("cl_fultvta") = data_cli.Recordset("cl_fultvta")
            data_inf.Recordset("cl_tipocli") = data_cli.Recordset("cl_tipocli")
            data_inf.Recordset("cl_celular") = data_cli.Recordset("cl_celular")
            data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
            data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
            data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
            data_inf.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
            data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
            data_inf.Recordset.Update
         End If
      End If
      data_cli.Recordset.MoveNext
   Loop
   data_inf.RecordSource = "Select * from infcli"
   data_inf.Refresh
   frm_infmodpad.MousePointer = 0
   MsgBox "Proceso terminado"
   cr11.ReportFileName = App.path & "\infmodpad.rpt"
   cr11.ReportTitle = "INFORME MODIFICACIONES REALIZADAS EN BASE "
   cr11.Action = 1
End If
frm_infmodpad.MousePointer = 0
Command1.Enabled = True
Command2.Enabled = True


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.ConnectionString = "dsn=" & Xconexrmt
'data_cli.RecordSource = "clientes"
'data_cli.Refresh
data_inf.DatabaseName = App.path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh

'data_lla.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lla.ConnectionString = "dsn=" & Xconexrmt
End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
