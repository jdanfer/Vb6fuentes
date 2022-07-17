VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infazul 
   BackColor       =   &H0080FF80&
   Caption         =   "Informes CODIGO AZUL"
   ClientHeight    =   3210
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infazul.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   5445
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_lla 
      Height          =   330
      Left            =   960
      Top             =   2880
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "data_lla"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2520
      Top             =   1920
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
      Picture         =   "frm_infazul.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_infazul.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Procesar"
      Top             =   2400
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos para informe"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_infazul.frx":109E
         Left            =   1680
         List            =   "frm_infazul.frx":10AE
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   3015
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "SELECCION:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "FECHAS:"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1560
      Picture         =   "frm_infazul.frx":10DB
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "frm_infazul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from inflla"
data_inf.RecordSource = "inflla"
data_inf.Refresh

If md.Text <> "__/__/____" And mh.Text <> "__/__/____" Then
   If Combo1.ListIndex = 0 Then 'todos
'''   Data1.RecordSource = "Select * from llamado where pend <>" & 2 & " And pend <>" & 1 & " and codmot ='" & "Z" & "' order by mm,nrolla"
      data_lla.RecordSource = "Select * from llamado where codmot ='" & "Z" & "' and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' order by fecha,mm"
   Else
      If Combo1.ListIndex = 1 Then 'terminados
         data_lla.RecordSource = "Select * from llamado where codmot ='" & "Z" & "' and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and pend =" & 2 & " and cancela is null order by fecha,mm"
      Else
         If Combo1.ListIndex = 2 Then 'pendientes
            data_lla.RecordSource = "Select * from llamado where codmot ='" & "Z" & "' and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and pend <>" & 2 & " and cancela is null order by fecha,mm"
         Else
            If Combo1.ListIndex = 3 Then
               data_lla.RecordSource = "Select * from llamado where codmot ='" & "Z" & "' and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and pend =" & 1 & " and cancela is null order by fecha,mm"
            Else
               data_lla.RecordSource = "Select * from llamado where codmot ='" & "Z" & "' and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cancela is null order by fecha,mm"
            End If
         End If
      End If
   End If
   data_lla.Refresh
   If data_lla.Recordset.RecordCount > 0 Then
      data_lla.Recordset.MoveFirst
      Do While Not data_lla.Recordset.EOF
         data_inf.Recordset.AddNew
         data_inf.Recordset("fecha") = data_lla.Recordset("fecha")
         data_inf.Recordset("hora") = data_lla.Recordset("hora")
         data_inf.Recordset("matric") = data_lla.Recordset("matric")
         data_inf.Recordset("categ") = data_lla.Recordset("categ")
         data_inf.Recordset("nomcat") = data_lla.Recordset("nomcat")
         data_inf.Recordset("edad") = data_lla.Recordset("edad")
         If data_lla.Recordset("unied") = 1 Then
            data_inf.Recordset("timdes") = "DIAS"
         Else
            If data_lla.Recordset("unied") = 2 Then
               data_inf.Recordset("timdes") = "MESES"
            Else
               data_inf.Recordset("timdes") = "AÑOS"
            End If
         End If
         data_inf.Recordset("referen") = data_lla.Recordset("referen")
         data_inf.Recordset("telef") = data_lla.Recordset("telef")
         data_inf.Recordset("motcon") = data_lla.Recordset("motcon")
         data_inf.Recordset("fecpas") = data_lla.Recordset("fecpas")
         data_inf.Recordset("horpas") = data_lla.Recordset("horpas")
         data_inf.Recordset("fec_rea") = data_lla.Recordset("fec_rea")
         data_inf.Recordset("hor_rea") = data_lla.Recordset("hor_rea")
         data_inf.Recordset("motmov") = data_lla.Recordset("motmov") 'zona
         data_inf.Recordset("ci") = data_lla.Recordset("ci")
         data_inf.Recordset("nombre") = data_lla.Recordset("nombre")
         data_inf.Recordset("mm") = data_lla.Recordset("mm")
         data_inf.Recordset.Update
         data_lla.Recordset.MoveNext
      Loop
      MsgBox "Terminado"
      data_inf.Refresh
      cr1.ReportFileName = App.path & "\infazul1.rpt"
      cr1.ReportTitle = "Informe de LLAMADOS CODIGO AZUL desde:" & md.Text & " hasta:" & mh.Text
      cr1.Action = 1
      
   End If
End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.RecordSource = "inflla"
'data_inf.Refresh
data_lla.ConnectionString = "dsn=" & Xconexrmt
'data_lla.Connect = "odbc;dsn=" & Xconexrmt & ";"

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
   Combo1.SetFocus
End If

End Sub
