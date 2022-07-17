VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_prodmed 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productividad de medicos"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_prodmed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   6915
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar bp1 
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   6120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command4 
      Caption         =   "HC95% y Trasl"
      Height          =   495
      Left            =   2400
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data data_inflla 
      Caption         =   "data_inflla"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc data_lla 
      Height          =   375
      Left            =   4680
      Top             =   1920
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
   Begin Crystal.CrystalReport cr3 
      Left            =   5520
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport cr2 
      Left            =   6240
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C00000&
      Caption         =   "Detalle"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   5760
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "Resumen"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   5760
      Value           =   -1  'True
      Width           =   1935
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   6360
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_hs 
      Caption         =   "data_hs"
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
      Top             =   2160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_med 
      Caption         =   "Data_med"
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
      Top             =   2160
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton b_fin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Terminar"
      Height          =   735
      Left            =   4320
      Picture         =   "frm_prodmed.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton b_proc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Procesar..."
      Height          =   735
      Left            =   600
      Picture         =   "frm_prodmed.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Datos requeridos"
      ForeColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      Begin VB.CommandButton Command6 
         Caption         =   "Solo findes"
         Height          =   495
         Left            =   2160
         TabIndex        =   27
         Top             =   4920
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00800000&
         Caption         =   "Fines de semana"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   4200
         Width           =   3975
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00800000&
         Caption         =   "Policlínicas Fin de semana"
         ForeColor       =   &H00FF80FF&
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   4800
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Data data_llaresp 
         Caption         =   "data_llaresp"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1920
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data data_llamsp 
         Caption         =   "data_llamsp"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1800
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sólo Médicos Incentivo"
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   4440
         TabIndex        =   24
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Data data_medhce 
         Caption         =   "data_medhce"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2040
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Data data_lla2 
         Caption         =   "data_lla2"
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
         Top             =   2520
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Data data_hce 
         Caption         =   "data_hce"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   3840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2280
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00800000&
         Caption         =   "Indicador Médicos"
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   3840
         Width           =   3975
      End
      Begin MSAdodcLib.Adodc data1 
         Height          =   495
         Left            =   4200
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
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
      Begin VB.CheckBox Check4 
         BackColor       =   &H00800000&
         Caption         =   "Incluir llamados de base"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   3480
         Width           =   3975
      End
      Begin VB.TextBox t_mov 
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00800000&
         Caption         =   "Desde respaldos"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4440
         TabIndex        =   17
         Top             =   4440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   5160
         TabIndex        =   14
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   5160
         TabIndex        =   13
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00800000&
         Caption         =   "Llamados y traslados 06a22hs"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3120
         Width           =   3975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00800000&
         Caption         =   "Informar únicamente TRASLADOS"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   3975
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_prodmed.frx":0F56
         Left            =   2400
         List            =   "frm_prodmed.frx":0F66
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Consultar médico"
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txt_med 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Text            =   "99"
         Top             =   1080
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "NUMERO MOVIL:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Código de llamado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Médico (99=Todos)"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Rango de Fechas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   3120
      Picture         =   "frm_prodmed.frx":0F88
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   2415
   End
End
Attribute VB_Name = "frm_prodmed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_fin_Click()
Unload Me

End Sub

Private Sub b_proc_Click()
Dim Xtotgral, Xtottras, Xtotllamed, Xtottramed, Xtothordom As Long
Dim xhh, xmm, Xhhh, Xmmh, xdemh, xdemm As Integer
frm_prodmed.MousePointer = 11

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from inflla"
data_inflla.RecordSource = "inflla"
data_inflla.Refresh


If Check7.Value = 1 Or Check8.Value = 1 Then
   If Check8.Value = 1 Then
      Command6_Click
   Else
      Command5_Click
   End If
Else
    If Check3.Value = 1 Then
    '   data_lla.DatabaseName = App.Path & "\llamado.mdb"
       data_lla.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.path & "\llamado.mdb"
    Else
       data_lla.ConnectionString = "dsn=" & Xconexrmt
    End If
    
    If md.Text <> "__/__/____" Then
       If mh.Text <> "__/__/____" Then
          If Check1.Value = 1 Or Check2.Value = 1 Or Check5.Value = 1 Then
             If Check5.Value = 1 Then
                Command4_Click
             Else
                Command2_Click
             End If
          Else
             If txt_med.Text = 99 Then
                If Combo1.ListIndex = 0 Then
                   If Check4.Value = 0 Then
                      data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codzon in(1,2,3,5) and cancela is null order by codmed"
                      data_lla.Refresh
                   Else
                      data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base >=" & 0 & " and codzon in(1,2,3,5) and cancela is null order by codmed"
                      data_lla.Refresh
                   End If
                Else
                   If Combo1.ListIndex = 1 Then
                      If Check4.Value = 0 Then
                         data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmot ='" & "V" & "' and codzon in(1,2,3,5) and cancela is null order by codmed"
                         data_lla.Refresh
                      Else
                         data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base >=" & 0 & " and codmot ='" & "V" & "' and codzon in(1,2,3,5) and cancela is null order by codmed"
                         data_lla.Refresh
                      End If
                   Else
                      If Combo1.ListIndex = 2 Then
                         If Check4.Value = 0 Then
                            data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmot ='" & "A" & "' and codzon in(1,2,3,5) and cancela is null order by codmed"
                            data_lla.Refresh
                         Else
                            data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base >=" & 0 & " and codmot ='" & "A" & "' and codzon in(1,2,3,5) and cancela is null order by codmed"
                            data_lla.Refresh
                         End If
                      Else
                         If Combo1.ListIndex = 3 Then
                            If Check4.Value = 0 Then
                               data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmot ='" & "R" & "' and codzon in(1,2,3,5) and cancela is null order by codmed"
                               data_lla.Refresh
                            Else
                               data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base >=" & 0 & " and codmot ='" & "R" & "' and codzon in(1,2,3,5) and cancela is null order by codmed"
                               data_lla.Refresh
                            End If
                         Else
                            If Check4.Value = 0 Then
                               data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codzon in(1,2,3,5) and cancela is null order by codmed"
                               data_lla.Refresh
                            Else
                               data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base >=" & 0 & " and codzon in(1,2,3,5) and cancela is null order by codmed"
                               data_lla.Refresh
                            End If
                         End If
                      End If
                   End If
                End If
             Else
                If Combo1.ListIndex = 0 Then
                   If Check4.Value = 0 Then
                      data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmed =" & txt_med.Text & " and codzon in(1,2,3,5) and cancela is null order by codmed"
                      data_lla.Refresh
                   Else
                      data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base >=" & 0 & " and codmed =" & txt_med.Text & " and codzon in(1,2,3,5) and cancela is null order by codmed"
                      data_lla.Refresh
                   End If
                Else
                   If Combo1.ListIndex = 1 Then
                      If Check4.Value = 0 Then
                         data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmot ='" & "V" & "' and codmed =" & txt_med.Text & " and codzon in(1,2,3,5) and cancela is null order by codmed"
                         data_lla.Refresh
                      Else
                         data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base >=" & 0 & " and codmot ='" & "V" & "' and codmed =" & txt_med.Text & " and codzon in(1,2,3,5) and cancela is null order by codmed"
                         data_lla.Refresh
                      End If
                   Else
                      If Combo1.ListIndex = 2 Then
                         If Check4.Value = 0 Then
                            data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmot ='" & "A" & "' and codmed =" & txt_med.Text & " and codzon in(1,2,3,5) and cancela is null order by codmed"
                            data_lla.Refresh
                         Else
                            data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base >=" & 0 & " and codmot ='" & "A" & "' and codmed =" & txt_med.Text & " and codzon in(1,2,3,5) and cancela is null order by codmed"
                            data_lla.Refresh
                         End If
                      Else
                         If Combo1.ListIndex = 3 Then
                            If Check4.Value = 0 Then
                               data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmot ='" & "R" & "' and codmed =" & txt_med.Text & " and codzon in(1,2,3,5) and cancela is null order by codmed"
                               data_lla.Refresh
                            Else
                               data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base >=" & 0 & " and codmot ='" & "R" & "' and codmed =" & txt_med.Text & " and codzon in(1,2,3,5) and cancela is null order by codmed"
                               data_lla.Refresh
                            End If
                         Else
                            If Check4.Value = 0 Then
                               data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmed =" & txt_med.Text & " and codzon in(1,2,3,5) and cancela is null order by codmed"
                               data_lla.Refresh
                            Else
                               data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base >=" & 0 & " and codmed =" & txt_med.Text & " and codzon in(1,2,3,5) and cancela is null order by codmed"
                               data_lla.Refresh
                            End If
                         End If
                      End If
                   End If
                End If
             End If
             If Check1.Value = 1 Then
                If Check4.Value = 0 Then
                   data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And trasla in (1,2,3,4,5,6,7,8,9,10,11,13) and base =" & 0 & " and cancela is null order by codmed"
                   data_lla.Refresh
                Else
                   data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And trasla in (1,2,3,4,5,6,7,8,9,10,11,13) and cancela is null order by codmed"
                   data_lla.Refresh
                End If
             End If
             If data_lla.Recordset.RecordCount > 0 Then
                data_lla.Recordset.MoveFirst
                Do While Not data_lla.Recordset.EOF
                   data_inflla.Recordset.AddNew
                   data_inflla.Recordset("nro") = data_lla.Recordset("nro")
                   data_inflla.Recordset("fecha") = data_lla.Recordset("fecha")
                   data_inflla.Recordset("hora") = data_lla.Recordset("hora")
                   data_inflla.Recordset("usuario") = data_lla.Recordset("usuario")
                   If data_lla.Recordset("matric") >= 999999999 Then
                      data_inflla.Recordset("matric") = 0
                   Else
                      data_inflla.Recordset("matric") = data_lla.Recordset("matric")
                   End If
                   data_inflla.Recordset("nombre") = data_lla.Recordset("nombre")
                   data_inflla.Recordset("edad") = data_lla.Recordset("edad")
                   data_inflla.Recordset("unied") = data_lla.Recordset("unied")
                   data_inflla.Recordset("categ") = data_lla.Recordset("categ")
                   data_inflla.Recordset("nomcat") = data_lla.Recordset("nomcat")
                   If data_lla.Recordset("ci") >= 999999999 Then
                      data_inflla.Recordset("ci") = 0
                   Else
                      data_inflla.Recordset("ci") = data_lla.Recordset("ci")
                   End If
                   data_inflla.Recordset("direcc") = data_lla.Recordset("direcc")
                   data_inflla.Recordset("telef") = data_lla.Recordset("telef")
                   data_inflla.Recordset("codzon") = data_lla.Recordset("codzon")
                   data_inflla.Recordset("base") = data_lla.Recordset("base")
                   data_inflla.Recordset("referen") = data_lla.Recordset("referen")
                   data_inflla.Recordset("motcon") = data_lla.Recordset("motcon")
                   data_inflla.Recordset("obsmot") = data_lla.Recordset("obsmot")
                   data_inflla.Recordset("codmot") = data_lla.Recordset("codmot")
                   If IsNull(data_inflla.Recordset("codmot")) = True Then
                      data_inflla.Recordset("pasado") = 0
                   Else
                      If data_inflla.Recordset("codmot") = "R" Then
                         data_inflla.Recordset("pasado") = 1
                      Else
                         If data_inflla.Recordset("codmot") = "A" Then
                            data_inflla.Recordset("pasado") = 2
                         Else
                            If data_inflla.Recordset("codmot") = "C" Then
                               data_inflla.Recordset("pasado") = 3
                            Else
                               data_inflla.Recordset("pasado") = 4
                            End If
                         End If
                      End If
                   End If
                   data_inflla.Recordset("descol") = data_lla.Recordset("descol")
                   data_inflla.Recordset("movilpas") = data_lla.Recordset("movilpas")
                   data_inflla.Recordset("pend") = data_lla.Recordset("pend")
                   If IsNull(data_lla.Recordset("fec_rea")) = True Then
                      data_inflla.Recordset("fec_rea") = data_lla.Recordset("fecpas")
                   Else
                      data_inflla.Recordset("fec_rea") = data_lla.Recordset("fec_rea")
                   End If
                   If IsNull(data_lla.Recordset("hor_rea")) = True Then
                      data_inflla.Recordset("hor_rea") = data_lla.Recordset("horpas")
                   Else
                      data_inflla.Recordset("hor_rea") = data_lla.Recordset("hor_rea")
                   End If
                   data_inflla.Recordset("fecpas") = data_lla.Recordset("fecpas")
                   data_inflla.Recordset("horpas") = data_lla.Recordset("horpas")
                   data_inflla.Recordset("fecsali") = data_lla.Recordset("fecsali")
                   data_inflla.Recordset("horsali") = data_lla.Recordset("horsali")
                   If IsNull(data_lla.Recordset("fec_llega")) = True Then
                      data_inflla.Recordset("fec_llega") = data_lla.Recordset("fecpas")
                   Else
                      data_inflla.Recordset("fec_llega") = data_lla.Recordset("fec_llega")
                   End If
                   If IsNull(data_lla.Recordset("hor_llega")) = True Then
                      data_inflla.Recordset("hor_llega") = data_lla.Recordset("horpas")
                   Else
                      data_inflla.Recordset("hor_llega") = data_lla.Recordset("hor_llega")
                   End If
                   data_inflla.Recordset("diag") = data_lla.Recordset("diag")
                   data_inflla.Recordset("colormot") = data_lla.Recordset("colormot")
                   data_inflla.Recordset("codmed") = data_lla.Recordset("codmed")
                   data_inflla.Recordset("obs") = data_lla.Recordset("obs")
                   data_inflla.Recordset("nommed") = data_lla.Recordset("nommed")
                   data_inflla.Recordset("trasla") = data_lla.Recordset("trasla")
                   If data_lla.Recordset("trasla") > 0 Then
                      Data1.RecordSource = "Select * from resplla where nro =" & data_lla.Recordset("nro")
                      Data1.Refresh
                      If Data1.Recordset.RecordCount > 0 Then
                         If IsNull(Data1.Recordset("descol")) = True Then
                         Else
                            data_inflla.Recordset("dcobr") = Data1.Recordset("descol")
                         End If
                      End If
                   End If
                   data_inflla.Recordset("lugar") = data_lla.Recordset("lugar")
                   data_inflla.Recordset("hsald") = data_lla.Recordset("hsald")
                   data_inflla.Recordset("hllega") = data_lla.Recordset("hllega")
                   data_inflla.Recordset("hzona") = data_lla.Recordset("hzona")
                   data_inflla.Recordset("movil_rea") = data_lla.Recordset("movil_rea")
                   data_inflla.Recordset("totdem") = data_lla.Recordset("totdem")
                   data_inflla.Recordset("totend") = data_lla.Recordset("totend")
                   data_inflla.Recordset("cancela") = data_lla.Recordset("cancela")
                   data_inflla.Recordset.Update
                   data_lla.Recordset.MoveNext
                Loop
     '           data_inflla.RecordSource = "select * from inflla order by fecha,hora"
    '           data_inflla.Refresh
    
                MiBaseact.Execute "Delete * from inflla where cancela =" & 1
                MiBaseact.Execute "Delete * from inflla where canteg in ('55','56','MSP')"
                data_inflla.Refresh
                data_inflla.Recordset.MoveFirst
                Do While Not data_inflla.Recordset.EOF
                   If IsNull(data_inflla.Recordset("hor_llega")) = True Then
                      data_inflla.Recordset.Edit
                      data_inflla.Recordset("totend") = "00:00"
                      data_inflla.Recordset("mm") = 0
                      data_inflla.Recordset.Update
                      data_inflla.Recordset.MoveNext
                   Else
                      If IsNull(data_inflla.Recordset("hor_llega")) = False Then
                         xhh = Val(Mid(data_inflla.Recordset("hor_llega"), 1, 2))
                         xmm = Val(Mid(data_inflla.Recordset("hor_llega"), 4, 2))
                      End If
                      If IsNull(data_inflla.Recordset("hor_rea")) = False Then
                         Xhhh = Val(Mid(data_inflla.Recordset("hor_rea"), 1, 2))
                         Xmmh = Val(Mid(data_inflla.Recordset("hor_rea"), 4, 2))
                      End If
                      xdemh = Xhhh - xhh
                      xdemm = Xmmh - xmm
                      If data_inflla.Recordset("fecha") < data_inflla.Recordset("fec_llega") Then
                         If xdemh < 0 Then
                            xdemh = Xhhh - xhh
                            xdemh = xdemh + 24
                         End If
                      Else
                         If IsNull(data_inflla.Recordset("fec_llega")) = True Then
                            xdemh = Xhhh - xhh
                            xdemh = xdemh + 24
                         Else
                            If xdemh < 0 Then
                               xdemh = xdemh + 24
                            End If
                         End If
                      End If
                      If xdemh > 0 Then
                         If xdemm < 0 Then
                            xdemm = xdemm + 60
                            xdemh = xdemh - 1
                         End If
                      Else
                         If Xmmh < xmm Then
                            xdemm = 0
                         Else
                            If xdemm < 0 Then
                               xdemm = xdemm + 60
                            End If
                         End If
                      End If
                      data_inflla.Recordset.Edit
                      If xdemh > 9 Then
                         If xdemm > 9 Then
                            data_inflla.Recordset("totend") = Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                         Else
                            data_inflla.Recordset("totend") = Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                         End If
                      Else
                         If xdemm > 9 Then
                            If xdemh < 0 Then
                               xdemh = 0
                            End If
                            data_inflla.Recordset("totend") = "0" + Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                         Else
                            If xdemh < 0 Then
                               xdemh = 0
                            End If
                            data_inflla.Recordset("totend") = "0" + Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                         End If
                      End If
                      If xdemh > 0 Then
                         data_inflla.Recordset("mm") = xdemh * 60 + xdemm
                      Else
                         data_inflla.Recordset("mm") = xdemm
                      End If
                      data_inflla.Recordset.Update
                      data_inflla.Recordset.MoveNext
                   End If
                Loop
                MsgBox "Proceso terminado"
                If Option2.Value = True Then
                   cr1.ReportFileName = App.path & "\infprodmed1.rpt"
                   cr1.ReportTitle = "INFORME DE TIEMPOS EN DOMICILIO LLAMADOS DESDE:" & md.Text & " HASTA:" & mh.Text
                Else
                   cr1.ReportFileName = App.path & "\infprodmed1n.rpt"
                   cr1.ReportTitle = "INFORME DE TIEMPOS EN DOMICILIO LLAMADOS DESDE:" & md.Text & " HASTA:" & mh.Text
                End If
                cr1.Action = 1
             End If
          End If
       End If
    End If
End If

frm_prodmed.MousePointer = 0

End Sub

Private Sub Command2_Click()
If Check2.Value = 1 Then
   Command3_Click
Else
        If txt_med.Text = 99 Then
           If Combo1.ListIndex = 0 Then
              If t_mov.Text = "" Then
                 If Check4.Value = 0 Then
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base =" & 0 & " and categ not in ('55','MSP','SAMC','SAMCB') and cancela is null order by codmed"
                    data_lla.Refresh
                 Else
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and categ not in ('55','MSP','SAMC','SAMCB') and cancela is null order by codmed"
                    data_lla.Refresh
                 End If
              Else
                 If Check4.Value = 0 Then
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and movilpas =" & t_mov.Text & " and base =" & 0 & " and cancela is null order by codmed"
                    data_lla.Refresh
                 Else
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and movilpas =" & t_mov.Text & " and cancela is null order by codmed"
                    data_lla.Refresh
                 End If
              End If
           Else
              If Combo1.ListIndex = 1 Then
                 If Check4.Value = 0 Then
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "V" & "' and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base =" & 0 & " and cancela is null order by codmed"
                    data_lla.Refresh
                 Else
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "V" & "' and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and cancela is null order by codmed"
                    data_lla.Refresh
                 End If
              Else
                 If Combo1.ListIndex = 2 Then
                    If Check4.Value = 0 Then
                       data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "A" & "' and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base =" & 0 & " and cancela is null order by codmed"
                       data_lla.Refresh
                    Else
                       data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "A" & "' and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and cancela is null order by codmed"
                       data_lla.Refresh
                    End If
                 Else
                    If Combo1.ListIndex = 3 Then
                       If Check4.Value = 0 Then
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "R" & "' and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base =" & 0 & " and cancela is null order by codmed"
                          data_lla.Refresh
                       Else
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "R" & "' and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and cancela is null order by codmed"
                          data_lla.Refresh
                       End If
                    Else
                       If Check4.Value = 0 Then
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base =" & 0 & " and categ not in ('55','MSP','SAMC','SAMCB') and cancela is null order by codmed"
                          data_lla.Refresh
                       Else
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and categ not in ('55','MSP','SAMC','SAMCB') and cancela is null order by codmed"
                          data_lla.Refresh
                       End If
                    End If
                 End If
              End If
           End If
        Else
           If Combo1.ListIndex = 0 Then
              If Check4.Value = 0 Then
                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmed =" & txt_med.Text & " and trasla in(1,2,3,4,5,6,7,8,9,10,11,13) and base =" & 0 & " and cancela is null order by codmed"
                 data_lla.Refresh
              Else
                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmed =" & txt_med.Text & " and trasla in(1,2,3,4,5,6,7,8,9,10,11,13) and cancela is null order by codmed"
                 data_lla.Refresh
              End If
           Else
              If Combo1.ListIndex = 1 Then
                 If Check4.Value = 0 Then
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmot ='" & "V" & "' and codmed =" & txt_med.Text & " and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base =" & 0 & " and cancela is null order by codmed"
                    data_lla.Refresh
                 Else
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmot ='" & "V" & "' and codmed =" & txt_med.Text & " and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and cancela is null order by codmed"
                    data_lla.Refresh
                 End If
              Else
                 If Combo1.ListIndex = 2 Then
                    If Check4.Value = 0 Then
                       data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmot ='" & "A" & "' and codmed =" & txt_med.Text & " and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base =" & 0 & " and cancela is null order by codmed"
                       data_lla.Refresh
                    Else
                       data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmot ='" & "A" & "' and codmed =" & txt_med.Text & " and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and cancela is null order by codmed"
                       data_lla.Refresh
                    End If
                 Else
                    If Combo1.ListIndex = 3 Then
                       If Check4.Value = 0 Then
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmot ='" & "R" & "' and codmed =" & txt_med.Text & " and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base =" & 0 & " and cancela is null order by codmed"
                          data_lla.Refresh
                       Else
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmot ='" & "R" & "' and codmed =" & txt_med.Text & " and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and cancela is null order by codmed"
                          data_lla.Refresh
                       End If
                    Else
                       If Check4.Value = 0 Then
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmed =" & txt_med.Text & " and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base =" & 0 & " and cancela is null order by codmed"
                          data_lla.Refresh
                       Else
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codmed =" & txt_med.Text & " and trasla in(1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and cancela is null order by codmed"
                          data_lla.Refresh
                       End If
                    End If
                 End If
              End If
           End If
        End If
        If data_lla.Recordset.RecordCount > 0 Then
           data_lla.Recordset.MoveFirst
           Do While Not data_lla.Recordset.EOF
              data_inflla.Recordset.AddNew
              data_inflla.Recordset("nro") = data_lla.Recordset("nro")
              data_inflla.Recordset("fecha") = data_lla.Recordset("fecha")
              data_inflla.Recordset("hora") = data_lla.Recordset("hora")
              data_inflla.Recordset("usuario") = data_lla.Recordset("usuario")
              If data_lla.Recordset("matric") >= 999999999 Then
                 data_inflla.Recordset("matric") = 0
              Else
                 data_inflla.Recordset("matric") = data_lla.Recordset("matric")
              End If
              data_inflla.Recordset("nombre") = data_lla.Recordset("nombre")
              data_inflla.Recordset("edad") = data_lla.Recordset("edad")
              data_inflla.Recordset("unied") = data_lla.Recordset("unied")
              data_inflla.Recordset("categ") = data_lla.Recordset("categ")
              data_inflla.Recordset("nomcat") = data_lla.Recordset("nomcat")
              If data_lla.Recordset("ci") >= 999999999 Then
                 data_inflla.Recordset("ci") = 0
              Else
                 data_inflla.Recordset("ci") = data_lla.Recordset("ci")
              End If
              data_inflla.Recordset("direcc") = data_lla.Recordset("direcc")
              data_inflla.Recordset("telef") = data_lla.Recordset("telef")
              data_inflla.Recordset("codzon") = data_lla.Recordset("codzon")
              data_inflla.Recordset("base") = data_lla.Recordset("base")
              data_inflla.Recordset("referen") = data_lla.Recordset("referen")
              data_inflla.Recordset("motcon") = data_lla.Recordset("motcon")
              data_inflla.Recordset("obsmot") = data_lla.Recordset("obsmot")
              data_inflla.Recordset("codmot") = data_lla.Recordset("codmot")
              If IsNull(data_inflla.Recordset("codmot")) = True Then
                 data_inflla.Recordset("pasado") = 0
              Else
                 If data_inflla.Recordset("codmot") = "R" Then
                    data_inflla.Recordset("pasado") = 1
                 Else
                    If data_inflla.Recordset("codmot") = "A" Then
                       data_inflla.Recordset("pasado") = 2
                    Else
                       If data_inflla.Recordset("codmot") = "C" Then
                          data_inflla.Recordset("pasado") = 3
                       Else
                          data_inflla.Recordset("pasado") = 4
                       End If
                    End If
                 End If
              End If
              data_inflla.Recordset("descol") = data_lla.Recordset("descol")
              data_inflla.Recordset("movilpas") = data_lla.Recordset("movilpas")
              data_inflla.Recordset("pend") = data_lla.Recordset("pend")
              If IsNull(data_lla.Recordset("fec_rea")) = True Then
                 data_inflla.Recordset("fec_rea") = data_lla.Recordset("fecpas")
              Else
                 data_inflla.Recordset("fec_rea") = data_lla.Recordset("fec_rea")
              End If
              If IsNull(data_lla.Recordset("hor_rea")) = True Then
                 data_inflla.Recordset("hor_rea") = data_lla.Recordset("horpas")
              Else
                 data_inflla.Recordset("hor_rea") = data_lla.Recordset("hor_rea")
              End If
              data_inflla.Recordset("fecpas") = data_lla.Recordset("fecpas")
              data_inflla.Recordset("horpas") = data_lla.Recordset("horpas")
              data_inflla.Recordset("fecsali") = data_lla.Recordset("fecsali")
              data_inflla.Recordset("horsali") = data_lla.Recordset("horsali")
              If IsNull(data_lla.Recordset("fec_llega")) = True Then
                 data_inflla.Recordset("fec_llega") = data_lla.Recordset("fecpas")
              Else
                 data_inflla.Recordset("fec_llega") = data_lla.Recordset("fec_llega")
              End If
              If IsNull(data_lla.Recordset("hor_llega")) = True Then
                 data_inflla.Recordset("hor_llega") = data_lla.Recordset("horpas")
              Else
                 data_inflla.Recordset("hor_llega") = data_lla.Recordset("hor_llega")
              End If
              data_inflla.Recordset("diag") = data_lla.Recordset("diag")
              data_inflla.Recordset("colormot") = data_lla.Recordset("colormot")
              data_inflla.Recordset("codmed") = data_lla.Recordset("codmed")
              data_inflla.Recordset("obs") = data_lla.Recordset("obs")
              data_inflla.Recordset("nommed") = data_lla.Recordset("nommed")
              data_inflla.Recordset("trasla") = data_lla.Recordset("trasla")
              If data_lla.Recordset("trasla") > 0 Then
                 Data1.RecordSource = "Select * from resplla where nro =" & data_lla.Recordset("nro")
                 Data1.Refresh
                 If Data1.Recordset.RecordCount > 0 Then
                    If IsNull(Data1.Recordset("descol")) = True Then
                    Else
                       data_inflla.Recordset("dcobr") = Data1.Recordset("descol")
                    End If
                 End If
              End If
              
              data_inflla.Recordset("lugar") = data_lla.Recordset("lugar")
              data_inflla.Recordset("hsald") = data_lla.Recordset("hsald")
              data_inflla.Recordset("hllega") = data_lla.Recordset("hllega")
              data_inflla.Recordset("hzona") = data_lla.Recordset("hzona")
              data_inflla.Recordset("movil_rea") = data_lla.Recordset("movil_rea")
              data_inflla.Recordset("totdem") = data_lla.Recordset("totdem")
              data_inflla.Recordset("totend") = data_lla.Recordset("totend")
              data_inflla.Recordset("cancela") = data_lla.Recordset("cancela")
              data_inflla.Recordset.Update
              data_lla.Recordset.MoveNext
           Loop
           data_inflla.RecordSource = "select * from inflla order by fecha,hora"
           data_inflla.Refresh
           If data_inflla.Recordset.RecordCount > 0 Then
              data_inflla.Recordset.MoveFirst
              Do While Not data_inflla.Recordset.EOF
                 If IsNull(data_inflla.Recordset("hsald")) = True Then
                    data_inflla.Recordset.Edit
                    data_inflla.Recordset("totend") = "00:00"
                    data_inflla.Recordset("mm") = 0
                    data_inflla.Recordset.Update
                    data_inflla.Recordset.MoveNext
                 Else
                    If IsNull(data_inflla.Recordset("hzona")) = True Then
                       data_inflla.Recordset.Edit
                       data_inflla.Recordset("totend") = "00:00"
                       data_inflla.Recordset("mm") = 0
                       data_inflla.Recordset.Update
                       data_inflla.Recordset.MoveNext
                    Else
                       If Len(data_inflla.Recordset("hsald")) < 5 Then
                          data_inflla.Recordset.Edit
                          data_inflla.Recordset("totend") = "00:00"
                          data_inflla.Recordset("mm") = 0
                          data_inflla.Recordset.Update
                          data_inflla.Recordset.MoveNext
                       Else
                          If Len(data_inflla.Recordset("hzona")) < 5 Then
                             data_inflla.Recordset.Edit
                             data_inflla.Recordset("totend") = "00:00"
                             data_inflla.Recordset("mm") = 0
                             data_inflla.Recordset.Update
                             data_inflla.Recordset.MoveNext
                          Else
                            If IsNull(data_inflla.Recordset("hsald")) = False Then
                               xhh = Val(Mid(data_inflla.Recordset("hsald"), 1, 2))
                               xmm = Val(Mid(data_inflla.Recordset("hsald"), 4, 2))
                            End If
                            If IsNull(data_inflla.Recordset("hzona")) = False Then
                               Xhhh = Val(Mid(data_inflla.Recordset("hzona"), 1, 2))
                               Xmmh = Val(Mid(data_inflla.Recordset("hzona"), 4, 2))
                            End If
                            xdemh = Xhhh - xhh
                            xdemm = Xmmh - xmm
                            If data_inflla.Recordset("fecha") < data_inflla.Recordset("fec_llega") Then
                               If xdemh < 0 Then
                                  xdemh = Xhhh - xhh
                                  xdemh = xdemh + 24
                               End If
                            Else
                               If IsNull(data_inflla.Recordset("fec_llega")) = True Then
                                  xdemh = Xhhh - xhh
                                  xdemh = xdemh + 24
                               Else
                                  If xdemh < 0 Then
                                     xdemh = xdemh + 24
                                  End If
                               End If
                            End If
                            If xdemh > 0 Then
                               If xdemm < 0 Then
                                  xdemm = xdemm + 60
                                  xdemh = xdemh - 1
                               End If
                            Else
                               If Xmmh < xmm Then
                                  xdemm = 0
                               Else
                                  If xdemm < 0 Then
                                     xdemm = xdemm + 60
                                  End If
                               End If
                            End If
                            data_inflla.Recordset.Edit
                            If xdemh > 9 Then
                               If xdemm > 9 Then
                                  data_inflla.Recordset("totend") = Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                               Else
                                  data_inflla.Recordset("totend") = Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                               End If
                            Else
                               If xdemm > 9 Then
                                  If xdemh < 0 Then
                                     xdemh = 0
                                  End If
                                  data_inflla.Recordset("totend") = "0" + Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                               Else
                                  If xdemh < 0 Then
                                     xdemh = 0
                                  End If
                                  data_inflla.Recordset("totend") = "0" + Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                               End If
                            End If
                            If xdemh > 0 Then
                               data_inflla.Recordset("mm") = xdemh * 60 + xdemm
                            Else
                               data_inflla.Recordset("mm") = xdemm
                            End If
                            data_inflla.Recordset.Update
                            data_inflla.Recordset.MoveNext
                          End If
                       End If
                    End If
                 End If
              Loop
              MsgBox "Proceso terminado"
              
              If Option2.Value = True Then
                 cr1.ReportFileName = App.path & "\infprodmed2.rpt"
                 cr1.ReportTitle = "INFORME DE TIEMPOS EN TRASLADOS DESDE:" & md.Text & " HASTA:" & mh.Text
              Else
                 cr1.ReportFileName = App.path & "\infprodmed2n.rpt"
                 cr1.ReportTitle = "INFORME DE TIEMPOS EN TRASLADOS DESDE:" & md.Text & " HASTA:" & mh.Text
              End If
              cr1.Action = 1
              
           End If
        End If
End If


End Sub

Private Sub Command3_Click()
        If txt_med.Text = 99 Then
           If Combo1.ListIndex = 0 Then
              If t_mov.Text = "" Then
                 If Check4.Value = 0 Then
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codzon in (1,2,3,5) and categ not in ('55','MSP','SAMC','SAMCB','50') and codmed <>" & 959 & " and cancela is null and movilpas not in (0,2015) order by codmed"
                    data_lla.Refresh
                 Else
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base >=" & 0 & " and codzon in (1,2,3,5) and categ not in ('55','MSP','SAMC','SAMCB','50') and codmed <>" & 959 & " and cancela is null and movilpas not in (0,2015) order by codmed"
'                    data_lla.RecordSource = "Select * from llamado where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# And base >=" & 0 & " and codzon in (1,2) and categ <>'" & "50" & "' and movilpas in (620) order by codmed"
                    data_lla.Refresh
                 End If
              Else
                 If Check4.Value = 0 Then
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codzon in (1,2,3,5) and categ not in ('55','MSP','SAMC','SAMCB','50') and codmed <>" & 959 & " and movilpas =" & t_mov.Text & " and cancela is null order by codmed"
                    data_lla.Refresh
                 Else
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base >=" & 0 & " and codzon in (1,2,3,5) and categ not in ('55','MSP','SAMC','SAMCB','50') and codmed <>" & 959 & " and movilpas =" & t_mov.Text & " and cancela is null order by codmed"
                    data_lla.Refresh
                 End If
              End If
           Else
              If Combo1.ListIndex = 1 Then
                 If Check4.Value = 0 Then
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "V" & "' And base =" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and codmed <>" & 959 & " and cancela is null order by codmed"
                    data_lla.Refresh
                 Else
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "V" & "' And base >=" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and codmed <>" & 959 & " and cancela is null order by codmed"
                    data_lla.Refresh
                 End If
              Else
                 If Combo1.ListIndex = 2 Then
                    If Check4.Value = 0 Then
                       data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "A" & "' And base =" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and codmed <>" & 959 & " and cancela is null order by codmed"
                       data_lla.Refresh
                    Else
                       data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "A" & "' And base >=" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and codmed <>" & 959 & " and cancela is null order by codmed"
                       data_lla.Refresh
                    End If
                 Else
                    If Combo1.ListIndex = 3 Then
                       If Check4.Value = 0 Then
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "R" & "' And base =" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and codmed <>" & 959 & " and cancela is null order by codmed"
                          data_lla.Refresh
                       Else
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "R" & "' And base >=" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and codmed <>" & 959 & " and cancela is null order by codmed"
                          data_lla.Refresh
                       End If
                    Else
                       If Check4.Value = 0 Then
'                          data_lla.RecordSource = "Select * from llamado where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# And base =" & 0 & " and codzon in (1,2) and categ <>'" & "50" & "' and codmed <>" & 959 & " and movilpas not in (620,602,601) order by codmed"
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codzon in (1,2,3,5) and categ not in ('55','MSP','SAMC','SAMCB','50') and cancela is null order by codmed"
                          data_lla.Refresh
                       Else
'                          data_lla.RecordSource = "Select * from llamado where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# And base >=" & 0 & " and codzon in (1,2) and categ <>'" & "50" & "' and codmed <>" & 959 & " and movilpas not in (620,602,601) order by codmed"
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base >=" & 0 & " and codzon in (1,2,3,5) and categ not in ('55','MSP','SAMC','SAMCB','50') and cancela is null order by codmed"
                          data_lla.Refresh
                       End If
                    End If
                 End If
              End If
           End If
        Else
           If Combo1.ListIndex = 0 Then
              If Check4.Value = 0 Then
                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmed =" & txt_med.Text & " And base =" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and cancela is null order by codmed"
                 data_lla.Refresh
              Else
                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmed =" & txt_med.Text & " And base >=" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and cancela is null order by codmed"
                 data_lla.Refresh
              End If
           Else
              If Combo1.ListIndex = 1 Then
                 If Check4.Value = 0 Then
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "V" & "' and codmed =" & txt_med.Text & " And base =" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and cancela is null order by codmed"
                    data_lla.Refresh
                 Else
                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "V" & "' and codmed =" & txt_med.Text & " And base >=" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and cancela is null order by codmed"
                    data_lla.Refresh
                 End If
              Else
                 If Combo1.ListIndex = 2 Then
                    If Check4.Value = 0 Then
                       data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "A" & "' and codmed =" & txt_med.Text & " And base =" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and cancela is null order by codmed"
                       data_lla.Refresh
                    Else
                       data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "A" & "' and codmed =" & txt_med.Text & " And base >=" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and cancela is null order by codmed"
                       data_lla.Refresh
                    End If
                 Else
                    If Combo1.ListIndex = 3 Then
                       If Check4.Value = 0 Then
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmot ='" & "R" & "' and codmed =" & txt_med.Text & " And base =" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and cancela is null order by codmed"
                          data_lla.Refresh
                       Else
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and codmot ='" & "R" & "' and codmed =" & txt_med.Text & " And base >=" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and cancela is null order by codmed"
                          data_lla.Refresh
                       End If
                    Else
                       If Check4.Value = 0 Then
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmed =" & txt_med.Text & " And base =" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and cancela is null order by codmed"
                          data_lla.Refresh
                       Else
                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codmed =" & txt_med.Text & " And base >=" & 0 & " and codzon in (1,2,3,5) and categ <>'" & "50" & "' and cancela is null order by codmed"
                          data_lla.Refresh
                       End If
                    End If
                 End If
              End If
           End If
        End If
        If data_lla.Recordset.RecordCount > 0 Then
           data_lla.Recordset.MoveFirst
           Do While Not data_lla.Recordset.EOF
              If IsNull(data_lla.Recordset("hor_rea")) = False Then
                 If data_lla.Recordset("hor_rea") >= "06:00" And data_lla.Recordset("hor_rea") <= "22:00" Then
                      data_inflla.Recordset.AddNew
                      data_inflla.Recordset("nro") = data_lla.Recordset("nro")
                      data_inflla.Recordset("fecha") = data_lla.Recordset("fecha")
                      data_inflla.Recordset("hora") = data_lla.Recordset("hora")
                      data_inflla.Recordset("usuario") = data_lla.Recordset("usuario")
                      If data_lla.Recordset("matric") >= 999999999 Then
                         data_inflla.Recordset("matric") = 0
                      Else
                         data_inflla.Recordset("matric") = data_lla.Recordset("matric")
                      End If
                      data_inflla.Recordset("nombre") = data_lla.Recordset("nombre")
                      data_inflla.Recordset("edad") = data_lla.Recordset("edad")
                      data_inflla.Recordset("unied") = data_lla.Recordset("unied")
                      data_inflla.Recordset("categ") = data_lla.Recordset("categ")
                      data_inflla.Recordset("nomcat") = data_lla.Recordset("nomcat")
                      If data_lla.Recordset("ci") >= 999999999 Then
                         data_inflla.Recordset("ci") = 0
                      Else
                         data_inflla.Recordset("ci") = data_lla.Recordset("ci")
                      End If
                      data_inflla.Recordset("direcc") = data_lla.Recordset("direcc")
                      data_inflla.Recordset("telef") = data_lla.Recordset("telef")
                      data_inflla.Recordset("codzon") = data_lla.Recordset("codzon")
                      data_inflla.Recordset("base") = data_lla.Recordset("base")
                      data_inflla.Recordset("referen") = data_lla.Recordset("referen")
                      data_inflla.Recordset("motcon") = data_lla.Recordset("motcon")
                      data_inflla.Recordset("obsmot") = data_lla.Recordset("obsmot")
                      data_inflla.Recordset("codmot") = data_lla.Recordset("codmot")
                      If IsNull(data_inflla.Recordset("codmot")) = True Then
                         data_inflla.Recordset("pasado") = 0
                      Else
                         If data_inflla.Recordset("codmot") = "R" Then
                            data_inflla.Recordset("pasado") = 1
                         Else
                            If data_inflla.Recordset("codmot") = "A" Then
                               data_inflla.Recordset("pasado") = 2
                            Else
                               If data_inflla.Recordset("codmot") = "C" Then
                                  data_inflla.Recordset("pasado") = 3
                               Else
                                  data_inflla.Recordset("pasado") = 4
                               End If
                            End If
                         End If
                      End If
                      data_inflla.Recordset("descol") = data_lla.Recordset("descol")
                      data_inflla.Recordset("movilpas") = data_lla.Recordset("movilpas")
                      data_inflla.Recordset("pend") = data_lla.Recordset("pend")
                      If IsNull(data_lla.Recordset("fec_rea")) = True Then
                         data_inflla.Recordset("fec_rea") = data_lla.Recordset("fecpas")
                      Else
                         data_inflla.Recordset("fec_rea") = data_lla.Recordset("fec_rea")
                      End If
                      If IsNull(data_lla.Recordset("hor_rea")) = True Then
                         data_inflla.Recordset("hor_rea") = data_lla.Recordset("horpas")
                      Else
                         data_inflla.Recordset("hor_rea") = data_lla.Recordset("hor_rea")
                      End If
                      data_inflla.Recordset("fecpas") = data_lla.Recordset("fecpas")
                      data_inflla.Recordset("horpas") = data_lla.Recordset("horpas")
                      data_inflla.Recordset("fecsali") = data_lla.Recordset("fecsali")
                      data_inflla.Recordset("horsali") = data_lla.Recordset("horsali")
                      If IsNull(data_lla.Recordset("fec_llega")) = True Then
                         data_inflla.Recordset("fec_llega") = data_lla.Recordset("fecpas")
                      Else
                         data_inflla.Recordset("fec_llega") = data_lla.Recordset("fec_llega")
                      End If
                      If IsNull(data_lla.Recordset("hor_llega")) = True Then
                         data_inflla.Recordset("hor_llega") = data_lla.Recordset("horpas")
                      Else
                         data_inflla.Recordset("hor_llega") = data_lla.Recordset("hor_llega")
                      End If
                      data_inflla.Recordset("diag") = data_lla.Recordset("diag")
                      data_inflla.Recordset("colormot") = data_lla.Recordset("colormot")
                      data_inflla.Recordset("codmed") = data_lla.Recordset("codmed")
                      data_inflla.Recordset("obs") = data_lla.Recordset("obs")
                      data_inflla.Recordset("nommed") = data_lla.Recordset("nommed")
                      data_inflla.Recordset("trasla") = data_lla.Recordset("trasla")
                        If data_lla.Recordset("trasla") > 0 Then
                           Data1.RecordSource = "Select * from resplla where nro =" & data_lla.Recordset("nro")
                           Data1.Refresh
                           If Data1.Recordset.RecordCount > 0 Then
                              If IsNull(Data1.Recordset("descol")) = True Then
                              Else
                                 data_inflla.Recordset("dcobr") = Data1.Recordset("descol")
                              End If
                           End If
                        End If
                      
                      data_inflla.Recordset("lugar") = data_lla.Recordset("lugar")
                      data_inflla.Recordset("hsald") = data_lla.Recordset("hsald")
                      data_inflla.Recordset("hllega") = data_lla.Recordset("hllega")
                      data_inflla.Recordset("hzona") = data_lla.Recordset("hzona")
                      data_inflla.Recordset("movil_rea") = data_lla.Recordset("movil_rea")
                      data_inflla.Recordset("movtras") = data_lla.Recordset("movtras")
                      data_inflla.Recordset("totdem") = data_lla.Recordset("totdem")
                      data_inflla.Recordset("totend") = data_lla.Recordset("totend")
                      data_inflla.Recordset("cancela") = data_lla.Recordset("cancela")
                      data_inflla.Recordset.Update
                 End If
              End If
              data_lla.Recordset.MoveNext
           Loop
           data_inflla.RecordSource = "select * from inflla order by fecha,hora"
           data_inflla.Refresh
           If data_inflla.Recordset.RecordCount > 0 Then
              data_inflla.Recordset.MoveFirst
              Do While Not data_inflla.Recordset.EOF
                 If IsNull(data_inflla.Recordset("cancela")) = False Then
                    If data_inflla.Recordset("cancela") = 1 Then
                       data_inflla.Recordset.Delete
                    Else
                    End If
                 Else
                 End If
                 data_inflla.Recordset.MoveNext
              Loop
              data_inflla.Recordset.MoveFirst
           End If
        End If
      data_inflla.RecordSource = "select * from inflla order by fecha,hora"
      data_inflla.Refresh
       MsgBox "Proceso terminado"
          If Option2.Value = True Then
             cr1.ReportFileName = App.path & "\infprodmed3.rpt"
             cr1.ReportTitle = "INFORME DE LLAMADOS HORARIO 06.00 A 22.00 DESDE:" & md.Text & " HASTA:" & mh.Text
             cr2.ReportFileName = App.path & "\infprodmed4.rpt"
             cr2.ReportTitle = "INFORME DE TRASLADOS POR MEDICO HORARIO 06.00 A 22.00 DESDE:" & md.Text & " HASTA:" & mh.Text
             cr3.ReportFileName = App.path & "\infprodmed5.rpt"
             cr3.ReportTitle = "INFORME DE TODOS LOS TRASLADOS POR MEDICO DE 06.00 A 22.00 DESDE:" & md.Text & " HASTA:" & mh.Text
          
          Else
             cr1.ReportFileName = App.path & "\infprodmed3n.rpt"
             cr1.ReportTitle = "INFORME DE LLAMADOS HORARIO 06.00 A 22.00 DESDE:" & md.Text & " HASTA:" & mh.Text
             cr2.ReportFileName = App.path & "\infprodmed4n.rpt"
             cr2.ReportTitle = "INFORME DE TRASLADOS POR MEDICO HORARIO 06.00 A 22.00 DESDE:" & md.Text & " HASTA:" & mh.Text
             cr3.ReportFileName = App.path & "\infprodmed5n.rpt"
             cr3.ReportTitle = "INFORME DE TODOS LOS TRASLADOS POR MEDICO DE 06.00 A 22.00 DESDE:" & md.Text & " HASTA:" & mh.Text
          
          End If
          cr1.Action = 1
          cr2.Action = 1
          cr3.Action = 1


End Sub

Private Sub Command4_Click()

Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub, Xdiurno, Xnocturno, Xradio, Xtraslasse, Xtotparahc, XtieneHC As Long
Dim Xarchtex As String
Dim Xnommedhc As String
Dim Xcodmed, Xcodmedhc As Integer
Dim Xlabrir As New Excel.Application
Dim Xporchc As Double
Dim Xfinde, XfindeCmt As Integer

Xporchc = 0
Xnommedhc = ""
Xcodmed = 0
Xtotreg = 0
Xsub = 0
Xdiurno = 0
Xnocturno = 0
Xradio = 0
Xtraslasse = 0
Xtotparahc = 0
XtieneHC = 0
Xfinde = 0
XfindeCmt = 0

Dim Xlabrir3 As New Excel.Application

If txt_med.Text <> 99 Then
   data_lla2.RecordSource = "Select * from llamado where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and movilpas not in (99,215,315,415,0) and cancela is null and codmed =" & txt_med.Text & " order by fecha"
Else
   If Trim(t_mov.Text) <> "" Then
      If Check6.Value = 1 Then
         data_lla2.RecordSource = "Select * from llamado where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and movilpas =" & t_mov.Text & " and cancela is null and codmed in (791,932,1382,1555,1594,1543,1629,1650,1647,1653,1671,1688,1695,1710,1721,1725,1557,1613,1484,1763,1777,1776,1785,1775,1780,1749,1812) order by codmed"
      Else
         data_lla2.RecordSource = "Select * from llamado where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and movilpas =" & t_mov.Text & " and cancela is null and codmed not in (0) order by codmed"
      End If
   Else
      If Check6.Value = 1 Then
         data_lla2.RecordSource = "Select * from llamado where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and movilpas not in (99,215,315,415,0) and cancela is null and codmed in (791,932,1382,1555,1594,1543,1650,1647,1653,1671,1688,1695,1710,1721,1725,1557,1613,1484,1629,1763,1777,1776,1785,1775,1780,1749,1812) order by codmed"
      Else
         data_lla2.RecordSource = "Select * from llamado where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and movilpas not in (99,215,315,415,0) and cancela is null and codmed not in (0) order by codmed"
      End If
   End If
End If
data_lla2.Refresh
bp1.Min = 0
bp1.Value = 0
If txt_med.Text = 99 Then
    If data_lla2.Recordset.RecordCount > 0 Then
       data_lla2.Recordset.MoveLast
       bp1.Max = data_lla2.Recordset.RecordCount + 1
       Xlin = 1
       XCol = 1
       Xtotreg = 0
       Xsub = 0
       Set Xobjexel22 = New Excel.Application
       Set Xlibexel22 = Xobjexel22.Workbooks.Add
       Set Xarchexel22 = Xlibexel22.Worksheets.Add
       Xarchexel22.Name = Trim("Controles")
       Xlibexel22.SaveAs ("C:\planillas\Controles por medico.xls")
       Xarchtex = "C:\planillas\Controles por medico.xls"
    
       Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
       Xlin = Xlin + 1
       XCol = XCol + 1
       Xarchexel22.Range("A1", "C3").Font.Size = 16
       Xarchexel22.Range("A" & Trim(str(Xlin)), "T" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
       Xarchexel22.Cells(Xlin, XCol) = "CONTROLES LLAMADOS POR MÉDICO DESDE: " & md.Text & " HASTA: " & mh.Text
            
       XCol = 1
       Xlin = Xlin + 2
       Xnrocan = Xnrocan + Xlin
                  
       Xarchexel22.Range("A" & Trim(str(Xlin)), "V" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
       Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
       Xarchexel22.Cells(Xlin, XCol) = "COD.MED"
       XCol = XCol + 1
       Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 25
       Xarchexel22.Cells(Xlin, XCol) = "NOMBRE del MEDICO"
       XCol = XCol + 1
       Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
       Xarchexel22.Cells(Xlin, XCol) = "LLAMADOS"
       XCol = XCol + 1
       Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
       Xarchexel22.Cells(Xlin, XCol) = "HCE REALIZA"
       XCol = XCol + 1
       Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
       Xarchexel22.Cells(Xlin, XCol) = "NOCTURNOS"
       XCol = XCol + 1
       Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
       Xarchexel22.Cells(Xlin, XCol) = "DIURNOS"
       XCol = XCol + 1
       Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
       Xarchexel22.Cells(Xlin, XCol) = "RADIO/CMT"
       XCol = XCol + 1
       Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
       Xarchexel22.Cells(Xlin, XCol) = "TRASL.ASSE"
       XCol = XCol + 1
       Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 10
       Xarchexel22.Cells(Xlin, XCol) = "%HCE"
       
       XCol = XCol + 1
       Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 10
       Xarchexel22.Cells(Xlin, XCol) = "ENCUESTA"
                        
       Xlin = Xlin + 1
       XCol = 1
       data_lla2.Recordset.MoveFirst
       Xcodmed = data_lla2.Recordset("codmed")
       data_medhce.RecordSource = "select * from medicos where med_cod =" & data_lla2.Recordset("codmed")
       data_medhce.Refresh
       If data_medhce.Recordset.RecordCount > 0 Then
          Xcodmedhc = data_medhce.Recordset("med_socnro")
       Else
          Xcodmedhc = 0
       End If
       Do While Not data_lla2.Recordset.EOF
          If Xcodmed = data_lla2.Recordset("codmed") Then
             Xsub = Xsub + 1
'             If Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 1 Or Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 7 Then
             If IsNull(data_lla2.Recordset("movilpas")) = False Then
                If data_lla2.Recordset("movilpas") = 10 Then
                   Xradio = Xradio + 1
                Else
                   If data_lla2.Recordset("base") <> 0 Then
                      Xradio = Xradio + 1
                   Else
                      If data_lla2.Recordset("movilpas") = 2015 Then
                         Xradio = Xradio + 1
                      Else
                         If IsNull(data_lla2.Recordset("horpas")) = False Then
                            If data_lla2.Recordset("horpas") <= "06:00" Then
                               Xnocturno = Xnocturno + 1
                            Else
                               If data_lla2.Recordset("horpas") >= "22:00" Then
                                  Xnocturno = Xnocturno + 1
                               Else
                                  Xdiurno = Xdiurno + 1
                               End If
                            End If
                         Else
                            Xdiurno = Xdiurno + 1
                         End If
                      End If
                   End If
                End If
             Else
                If IsNull(data_lla2.Recordset("horpas")) = False Then
                   If data_lla2.Recordset("horpas") <= "06:00" Then
                      Xnocturno = Xnocturno + 1
                   Else
                      If data_lla2.Recordset("horpas") >= "22:00" Then
                         Xnocturno = Xnocturno + 1
                      Else
                         Xdiurno = Xdiurno + 1
                      End If
                   End If
                Else
                   Xdiurno = Xdiurno + 1
                End If
             End If
             If IsNull(data_lla2.Recordset("categ")) = False Then
                If data_lla2.Recordset("categ") = "911" Or data_lla2.Recordset("categ") = "911B" Or _
                   data_lla2.Recordset("categ") = "CAAMEP" Then
                Else
                   Xtotparahc = Xtotparahc + 1
                End If
             Else
                Xtotparahc = Xtotparahc + 1
             End If
             If data_lla2.Recordset("categ") = "911" Or data_lla2.Recordset("categ") = "911B" Or _
                data_lla2.Recordset("categ") = "CAAMEP" Or data_lla2.Recordset("categ") = "MSP" Then
             Else
                data_hce.RecordSource = "select * from cabezal_hcdig where hc_codmed =" & Xcodmedhc & " and fecha =#" & Format(data_lla2.Recordset("fec_rea"), "yyyy/mm/dd") & "# and tipo_consd in ('Consulta Domicilio','Orientación Telefónica') and cednum =" & data_lla2.Recordset("ci")
                data_hce.Refresh
                If data_hce.Recordset.RecordCount > 0 Then
                   XtieneHC = XtieneHC + 1
                End If
             End If
          Else
             data_lla2.Recordset.MovePrevious
             data_llamsp.RecordSource = "select llamado.nro,llamado.fecha,llamado.ci,llamado.fec_rea,llamado.categ,resplla.movilpas from llamado " & _
             "inner join resplla on llamado.nro=resplla.nro where llamado.fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and llamado.fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and llamado.categ in ('MSP') and resplla.movilpas =" & data_lla2.Recordset("codmed")
             data_llamsp.Refresh
             If data_llamsp.Recordset.RecordCount > 0 Then
                data_llamsp.Recordset.MoveFirst
                Do While Not data_llamsp.Recordset.EOF
                   data_hce.RecordSource = "select * from cabezal_hcdig where hc_codmed =" & Xcodmedhc & " and fecha =#" & Format(data_llamsp.Recordset("fec_rea"), "yyyy/mm/dd") & "# and tipo_consd in ('Consulta Domicilio') and cednum =" & data_llamsp.Recordset("ci")
                   data_hce.Refresh
                   If data_hce.Recordset.RecordCount > 0 Then
                      XtieneHC = XtieneHC + 1
                   End If
                   data_llamsp.Recordset.MoveNext
                Loop
'                data_llamsp.Recordset.MoveLast
                Xtraslasse = data_llamsp.Recordset.RecordCount
             Else
                Xtraslasse = 0
             End If
             data_inflla.Recordset.AddNew
             data_inflla.Recordset("codmed") = data_lla2.Recordset("codmed")
             data_inflla.Recordset("nommed") = data_lla2.Recordset("nommed")
             data_inflla.Recordset("matric") = Xsub 'llamados
             data_inflla.Recordset("edad") = Xdiurno
             data_inflla.Recordset("codzon") = Xnocturno
             data_inflla.Recordset("pasado") = Xradio
             data_inflla.Recordset("movilpas") = Xtraslasse
             If XtieneHC > 0 Then
                data_inflla.Recordset("realiza") = XtieneHC
                If XtieneHC > Xtotparahc Then
                   data_inflla.Recordset("movil_rea") = 100
                Else
                   data_inflla.Recordset("movil_rea") = XtieneHC / Xtotparahc * 100
                End If
             Else
                data_inflla.Recordset("realiza") = 0
                data_inflla.Recordset("movil_rea") = 0
             End If
             data_inflla.Recordset("trasla") = 0
             data_inflla.Recordset.Update
             Xsub = 0
             Xdiurno = 0
             Xnocturno = 0
             Xradio = 0
             Xtraslasse = 0
             Xtotparahc = 0
             XtieneHC = 0
             data_lla2.Recordset.MoveNext
             If Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 1 Or Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 7 Then
                Xfinde = Xfinde + 1
             End If
             data_medhce.RecordSource = "select * from medicos where med_cod =" & data_lla2.Recordset("codmed")
             data_medhce.Refresh
             If data_medhce.Recordset.RecordCount > 0 Then
                If IsNull(data_medhce.Recordset("med_socnro")) = False Then
                   If IsNull(data_medhce.Recordset("med_socnro")) = False Then
                      Xcodmedhc = data_medhce.Recordset("med_socnro")
                   Else
                      Xcodmedhc = 0
                   End If
                Else
                   Xcodmedhc = 0
                End If
             Else
                Xcodmedhc = 0
             End If
             If data_lla2.Recordset("categ") = "911" Or data_lla2.Recordset("categ") = "911B" Or _
                data_lla2.Recordset("categ") = "CAAMEP" Or data_lla2.Recordset("categ") = "MSP" Then
             Else
                data_hce.RecordSource = "select * from cabezal_hcdig where hc_codmed =" & Xcodmedhc & " and fecha =#" & Format(data_lla2.Recordset("fec_rea"), "yyyy/mm/dd") & "# and tipo_consd in ('Consulta Domicilio','Orientación Telefónica') and cednum =" & data_lla2.Recordset("ci")
                data_hce.Refresh
                If data_hce.Recordset.RecordCount > 0 Then
                   XtieneHC = XtieneHC + 1
                End If
             End If
             Xsub = Xsub + 1
             If IsNull(data_lla2.Recordset("movilpas")) = False Then
                If data_lla2.Recordset("movilpas") = 10 Then
                   Xradio = Xradio + 1
                Else
                   If data_lla2.Recordset("base") <> 0 Then
                      Xradio = Xradio + 1
                   Else
                      If data_lla2.Recordset("movilpas") = 2015 Then
                         Xradio = Xradio + 1
                      Else
                         If IsNull(data_lla2.Recordset("horpas")) = False Then
                            If data_lla2.Recordset("horpas") <= "06:00" Then
                               Xnocturno = Xnocturno + 1
                            Else
                               If data_lla2.Recordset("horpas") >= "22:00" Then
                                  Xnocturno = Xnocturno + 1
                               Else
                                  Xdiurno = Xdiurno + 1
                               End If
                            End If
                         Else
                            Xdiurno = Xdiurno + 1
                         End If
                      End If
                   End If
                End If
             Else
                If IsNull(data_lla2.Recordset("horpas")) = False Then
                   If data_lla2.Recordset("horpas") <= "06:00" Then
                      Xnocturno = Xnocturno + 1
                   Else
                      If data_lla2.Recordset("horpas") >= "22:00" Then
                         Xnocturno = Xnocturno + 1
                      Else
                         Xdiurno = Xdiurno + 1
                      End If
                   End If
                Else
                   Xdiurno = Xdiurno + 1
                End If
             End If
             If IsNull(data_lla2.Recordset("trasla")) = False Then
                If data_lla2.Recordset("trasla") > 0 Then
                   If data_lla2.Recordset("categ") = "MSP" Then
                      Xtraslasse = Xtraslasse + 1
                   End If
                End If
             End If
             If IsNull(data_lla2.Recordset("categ")) = False Then
                If data_lla2.Recordset("categ") = "911" Or data_lla2.Recordset("categ") = "911B" Or _
                   data_lla2.Recordset("categ") = "CAAMEP" Then
                Else
                   Xtotparahc = Xtotparahc + 1
                End If
             Else
                Xtotparahc = Xtotparahc + 1
             End If
          
          End If
          Xcodmed = data_lla2.Recordset("codmed")
          data_lla2.Recordset.MoveNext
          bp1.Value = bp1.Value + 1
       Loop
       
        data_lla2.Recordset.MovePrevious
        
        data_llamsp.RecordSource = "select llamado.nro,llamado.fecha,llamado.ci,llamado.fec_rea,llamado.categ,resplla.movilpas from llamado " & _
        "inner join resplla on llamado.nro=resplla.nro where llamado.fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and llamado.fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and llamado.categ in ('MSP') and resplla.movilpas =" & data_lla2.Recordset("codmed")
        data_llamsp.Refresh
        If data_llamsp.Recordset.RecordCount > 0 Then
           data_llamsp.Recordset.MoveFirst
           Do While Not data_llamsp.Recordset.EOF
              data_hce.RecordSource = "select * from cabezal_hcdig where hc_codmed =" & Xcodmedhc & " and fecha =#" & Format(data_llamsp.Recordset("fec_rea"), "yyyy/mm/dd") & "# and tipo_consd in ('Consulta Domicilio') and cednum =" & data_llamsp.Recordset("ci")
              data_hce.Refresh
              If data_hce.Recordset.RecordCount > 0 Then
                 XtieneHC = XtieneHC + 1
              End If
              data_llamsp.Recordset.MoveNext
           Loop
           Xtraslasse = data_llamsp.Recordset.RecordCount
        Else
           Xtraslasse = 0
        End If
        data_inflla.Recordset.AddNew
        data_inflla.Recordset("codmed") = data_lla2.Recordset("codmed")
        data_inflla.Recordset("nommed") = data_lla2.Recordset("nommed")
        data_inflla.Recordset("matric") = Xsub 'llamados
        data_inflla.Recordset("edad") = Xdiurno
        data_inflla.Recordset("codzon") = Xnocturno
        data_inflla.Recordset("pasado") = Xradio
        data_inflla.Recordset("movilpas") = Xtraslasse
        If XtieneHC > 0 Then
           data_inflla.Recordset("realiza") = XtieneHC
           If XtieneHC > Xtotparahc Then
              data_inflla.Recordset("movil_rea") = 100
           Else
              data_inflla.Recordset("movil_rea") = XtieneHC / Xtotparahc * 100
           End If
        Else
           data_inflla.Recordset("realiza") = 0
           data_inflla.Recordset("movil_rea") = 0
        End If
        data_inflla.Recordset("trasla") = 0
        data_inflla.Recordset.Update
        
        Xsub = Xsub + 1
        If IsNull(data_lla2.Recordset("trasla")) = False Then
           If data_lla2.Recordset("trasla") > 0 Then
              If data_lla2.Recordset("categ") = "MSP" Then
                 Xtraslasse = Xtraslasse + 1
              End If
           End If
        End If
        If IsNull(data_lla2.Recordset("categ")) = False Then
           If data_lla2.Recordset("categ") = "911" Or data_lla2.Recordset("categ") = "911B" Or _
              data_lla2.Recordset("categ") = "CAAMEP" Then
           Else
              Xtotparahc = Xtotparahc + 1
           End If
        Else
           Xtotparahc = Xtotparahc + 1
        End If
        
       data_inflla.Refresh
       If data_inflla.Recordset.RecordCount > 0 Then
          data_inflla.Recordset.MoveFirst
          Do While Not data_inflla.Recordset.EOF
             Xarchexel22.Cells(Xlin, XCol) = data_inflla.Recordset("codmed")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = data_inflla.Recordset("nommed")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = data_inflla.Recordset("matric")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = data_inflla.Recordset("realiza")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = data_inflla.Recordset("codzon")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = data_inflla.Recordset("edad")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = data_inflla.Recordset("pasado")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = data_inflla.Recordset("movilpas")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = Format(data_inflla.Recordset("movil_rea"), "Standard")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = data_inflla.Recordset("trasla")
             data_inflla.Recordset.MoveNext
             XCol = 1
             Xlin = Xlin + 1
          Loop
          frm_prodmed.MousePointer = 0
          
          Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")
          Xlibexel22.Save
          Xlibexel22.Close
          Xobjexel22.Quit
       
          Xlabrir.Workbooks.Open Xarchtex, , False
          Xlabrir.Visible = True
          Xlabrir.WindowState = xlMaximized
       
       Else
          frm_prodmed.MousePointer = 0
          MsgBox "No hay registros para crear planilla."
       End If
    Else
       frm_prodmed.MousePointer = 0
       MsgBox "No hay registros"
    End If
Else
    If data_lla2.Recordset.RecordCount > 0 Then
       data_lla2.Recordset.MoveLast
       bp1.Max = data_lla2.Recordset.RecordCount + 1
       bp1.Min = 0
       Xlin = 1
       XCol = 1
       Xtotreg = 0
       Xsub = 0
       Set Xobjexel22 = New Excel.Application
       Set Xlibexel22 = Xobjexel22.Workbooks.Add
       Set Xarchexel22 = Xlibexel22.Worksheets.Add
       Xarchexel22.Name = Trim("Controles")
       Xlibexel22.SaveAs ("C:\planillas\Controles por medico.xls")
       Xarchtex = "C:\planillas\Controles por medico.xls"
    
       Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
       Xlin = Xlin + 1
       XCol = XCol + 1
       Xarchexel22.Range("A1", "C3").Font.Size = 16
       Xarchexel22.Range("A" & Trim(str(Xlin)), "T" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
       Xarchexel22.Cells(Xlin, XCol) = "LLAMADOS POR MÉDICO DESDE: " & md.Text & " HASTA: " & mh.Text
       XCol = 1
       Xlin = Xlin + 1
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.Size = 16
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.color = &H80FF&
       Xarchexel22.Cells(Xlin, XCol) = "MÉDICO: " & data_lla2.Recordset("nommed")
       
       XCol = 1
       Xlin = Xlin + 2
       Xnrocan = Xnrocan + Xlin
                  
       Xarchexel22.Range("A" & Trim(str(Xlin)), "V" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
       Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
       Xarchexel22.Cells(Xlin, XCol) = "COD.MED"
       XCol = XCol + 1
       Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 25
       Xarchexel22.Cells(Xlin, XCol) = "NOMBRE del MEDICO"
       XCol = XCol + 1
       Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
       Xarchexel22.Cells(Xlin, XCol) = "FECHA"
       XCol = XCol + 1
       Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 25
       Xarchexel22.Cells(Xlin, XCol) = "NOMBRE PACIENTE"
       XCol = XCol + 1
       Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
       Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
       XCol = XCol + 1
       Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
       Xarchexel22.Cells(Xlin, XCol) = "MOVIL"
       XCol = XCol + 1
       Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 14
       Xarchexel22.Cells(Xlin, XCol) = "HORA PASADO"
       XCol = XCol + 1
       Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
       Xarchexel22.Cells(Xlin, XCol) = "TIENE HCE?"
            
       Xlin = Xlin + 1
       XCol = 1
       data_lla2.Recordset.MoveFirst
       Xcodmed = data_lla2.Recordset("codmed")
       data_medhce.RecordSource = "select * from medicos where med_cod =" & data_lla2.Recordset("codmed")
       data_medhce.Refresh
       If data_medhce.Recordset.RecordCount > 0 Then
          Xcodmedhc = data_medhce.Recordset("med_socnro")
          Xnommedhc = data_medhce.Recordset("med_nombre")
       Else
          Xcodmedhc = 0
          Xnommedhc = "S/D"
       End If
       Xsub = 0
       Do While Not data_lla2.Recordset.EOF
             Xsub = Xsub + 1
             If IsNull(data_lla2.Recordset("movilpas")) = False Then
                If data_lla2.Recordset("movilpas") = 10 Then
                   Xradio = Xradio + 1
                Else
                   If data_lla2.Recordset("base") <> 0 Then
                      Xradio = Xradio + 1
                   Else
                      If data_lla2.Recordset("movilpas") = 2015 Then
                         Xradio = Xradio + 1
                      Else
                         If IsNull(data_lla2.Recordset("horpas")) = False Then
                            If data_lla2.Recordset("horpas") <= "06:00" Then
                               Xnocturno = Xnocturno + 1
                            Else
                               If data_lla2.Recordset("horpas") >= "22:00" Then
                                  Xnocturno = Xnocturno + 1
                               Else
                                  Xdiurno = Xdiurno + 1
                               End If
                            End If
                         Else
                            Xdiurno = Xdiurno + 1
                         End If
                      End If
                   End If
                End If
             Else
                If IsNull(data_lla2.Recordset("horpas")) = False Then
                   If data_lla2.Recordset("horpas") <= "06:00" Then
                      Xnocturno = Xnocturno + 1
                   Else
                      If data_lla2.Recordset("horpas") >= "22:00" Then
                         Xnocturno = Xnocturno + 1
                      Else
                         Xdiurno = Xdiurno + 1
                      End If
                   End If
                Else
                   Xdiurno = Xdiurno + 1
                End If
             End If
                          
             If data_lla2.Recordset("categ") = "911" Or data_lla2.Recordset("categ") = "911B" Or _
                data_lla2.Recordset("categ") = "CAAMEP" Then
             Else
                Xtotreg = Xtotreg + 1
             End If
             Xarchexel22.Cells(Xlin, XCol) = data_lla2.Recordset("codmed")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = data_lla2.Recordset("nommed")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_lla2.Recordset("fecha"), "dd/mm/yyyy")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = data_lla2.Recordset("nombre")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = data_lla2.Recordset("ci")
             XCol = XCol + 1
             If IsNull(data_lla2.Recordset("base")) = False Then
                If data_lla2.Recordset("base") > 0 Then
                   Xarchexel22.Cells(Xlin, XCol) = "BASE"
                Else
                   If data_lla2.Recordset("movilpas") = 2015 Then
                      Xarchexel22.Cells(Xlin, XCol) = "CMT"
                   Else
                      Xarchexel22.Cells(Xlin, XCol) = data_lla2.Recordset("movilpas")
                   End If
                End If
             Else
                Xarchexel22.Cells(Xlin, XCol) = data_lla2.Recordset("movilpas")
             End If
             XCol = XCol + 1
             If IsNull(data_lla2.Recordset("horpas")) = False Then
                Xarchexel22.Cells(Xlin, XCol) = data_lla2.Recordset("horpas")
             End If
             XCol = XCol + 1
             If data_lla2.Recordset("categ") = "911" Or data_lla2.Recordset("categ") = "911B" Or _
                data_lla2.Recordset("categ") = "CAAMEP" Then
                If data_lla2.Recordset("categ") = "911" Or data_lla2.Recordset("categ") = "911B" Then
                   Xarchexel22.Cells(Xlin, XCol) = "911"
                Else
                   Xarchexel22.Cells(Xlin, XCol) = "CAAMEPA"
                End If
             Else
                data_hce.RecordSource = "select * from cabezal_hcdig where hc_codmed =" & Xcodmedhc & " and fecha =#" & Format(data_lla2.Recordset("fec_rea"), "yyyy/mm/dd") & "# and tipo_consd in ('Consulta Domicilio','Orientación Telefónica') and cednum =" & data_lla2.Recordset("ci")
                data_hce.Refresh
                If data_hce.Recordset.RecordCount > 0 Then
                   Xarchexel22.Cells(Xlin, XCol) = "SI"
                   Xtotparahc = Xtotparahc + 1
                Else
                   Xarchexel22.Cells(Xlin, XCol) = "NO"
                End If
             End If
             Xlin = Xlin + 1
             XCol = 1
          data_lla2.Recordset.MoveNext
          bp1.Value = bp1.Value + 1
       Loop
       data_lla2.Recordset.MovePrevious
       
       data_llamsp.RecordSource = "select llamado.nro,llamado.fecha,llamado.ci,llamado.horpas,llamado.categ,llamado.base,llamado.movilpas,llamado.fec_rea,llamado.categ,llamado.codmed,llamado.nommed,llamado.nombre,resplla.movilpas from llamado " & _
       "inner join resplla on llamado.nro=resplla.nro where llamado.fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and llamado.fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and llamado.categ in ('MSP') and resplla.movilpas =" & data_lla2.Recordset("codmed")
       data_llamsp.Refresh
       If data_llamsp.Recordset.RecordCount > 0 Then
          data_llamsp.Recordset.MoveFirst
          Do While Not data_llamsp.Recordset.EOF
             
             Xarchexel22.Cells(Xlin, XCol) = data_llamsp.Recordset("resplla.movilpas")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = Xnommedhc
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_llamsp.Recordset("fecha"), "dd/mm/yyyy")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = data_llamsp.Recordset("nombre")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = data_llamsp.Recordset("ci")
             XCol = XCol + 1
             Xarchexel22.Cells(Xlin, XCol) = data_llamsp.Recordset("llamado.movilpas")
             XCol = XCol + 1
             If IsNull(data_llamsp.Recordset("horpas")) = False Then
                Xarchexel22.Cells(Xlin, XCol) = data_llamsp.Recordset("horpas")
             End If
             XCol = XCol + 1
             data_hce.RecordSource = "select * from cabezal_hcdig where hc_codmed =" & Xcodmedhc & " and fecha =#" & Format(data_llamsp.Recordset("fec_rea"), "yyyy/mm/dd") & "# and tipo_consd in ('Consulta Domicilio') and cednum =" & data_llamsp.Recordset("ci")
             data_hce.Refresh
             If data_hce.Recordset.RecordCount > 0 Then
                XtieneHC = XtieneHC + 1
                Xtotparahc = Xtotparahc + 1
                Xarchexel22.Cells(Xlin, XCol) = "SI -TRASL.ASSE"
             Else
                Xarchexel22.Cells(Xlin, XCol) = "NO -TRASL.ASSE"
             End If
             data_llamsp.Recordset.MoveNext
             Xlin = Xlin + 1
             XCol = 1
          Loop
          Xtraslasse = data_llamsp.Recordset.RecordCount
       Else
          Xtraslasse = 0
       End If
        
       XCol = 1
       Xlin = Xlin + 1
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.Size = 16
       Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE LLAMADOS: " & Xsub
       Xlin = Xlin + 1
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.Size = 16
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.color = &HC00000
       XCol = 1
       Xporchc = Xtotparahc / Xtotreg * 100
       Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE HCE REALIZADAS: " & Xtotparahc & " -->PORCENTAJE:" & Format(Xporchc, "Standard") & " %"
       Xlin = Xlin + 1
       XCol = 1
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.Size = 16
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.color = &HC00000
       Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE LLAMADOS DIURNOS: " & Xdiurno
       Xlin = Xlin + 1
       XCol = 1
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.Size = 16
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.color = &HC00000
       Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE LLAMADOS NOCTURNOS: " & Xnocturno
       Xlin = Xlin + 1
       XCol = 1
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.Size = 16
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.color = &HC00000
       Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE LLAMADOS DE RADIO: " & Xradio
       Xlin = Xlin + 1
       XCol = 1
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.Size = 16
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.color = &HC00000
       Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE TRASLADOS ASSE: " & Xtraslasse
       Xlin = Xlin + 1
                    
       XCol = 1
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.Size = 16
       Xarchexel22.Range("A" & Xlin, "C" & Xlin).Font.color = &HC00000
       Xarchexel22.Cells(Xlin, XCol) = "ENCUESTA: Sin Datos. (Está pendiente para desarrollar)"
       Xlin = Xlin + 1
          
       Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")
       Xlibexel22.Save
       Xlibexel22.Close
       Xobjexel22.Quit
       XtieneHC = 0
       
       frm_prodmed.MousePointer = 0
       
       Xlabrir.Workbooks.Open Xarchtex, , False
       Xlabrir.Visible = True
       Xlabrir.WindowState = xlMaximized
       
    Else
       frm_prodmed.MousePointer = 0
       MsgBox "No hay registros para crear planilla."
    End If

End If
frm_prodmed.MousePointer = 0




End Sub

Private Sub Command5_Click()


End Sub

Private Sub Command6_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub, Xdiurno, Xnocturno, Xradio, Xtraslasse, Xtotparahc, XtieneHC As Long
Dim Xarchtex As String
Dim Xnommedhc As String
Dim Xcodmed, Xcodmedhc As Integer
Dim Xlabrir As New Excel.Application
Dim Xporchc As Double
Dim Xfinde As Integer

Xporchc = 0
Xnommedhc = ""
Xcodmed = 0
Xtotreg = 0
Xsub = 0
Xdiurno = 0
Xnocturno = 0
Xradio = 0
Xtraslasse = 0
Xtotparahc = 0
XtieneHC = 0
Xfinde = 0
Dim Xlabrir3 As New Excel.Application

If txt_med.Text <> 99 Then
   data_lla2.RecordSource = "Select * from llamado where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and movilpas not in (99,215,315,415,0,5,202,217) and cancela is null and codmed =" & txt_med.Text & " order by fecha"
Else
   data_lla2.RecordSource = "Select * from llamado where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and movilpas not in (99,215,315,415,0,202,217,5) and cancela is null and codmed not in (0,959,440,845,959,658,705,579,536,1694,1105,407,751,849,435,717,652,439,414,437,702,465,1231,566,662,1004,516,556,979,1541,668,403,515,560,693) order by codmed"
End If
'414,435,437,439,465,536,556,566,574,579,652,658,668,684,705,751,979,1004,1045,1227,1231,1483,1541,1584,1631,1633,1667,1694,1734,1747,1750,1753,1760,1765,1779,1782,1793

data_lla2.Refresh
bp1.Min = 0
bp1.Value = 0
If txt_med.Text = 99 Then
    If data_lla2.Recordset.RecordCount > 0 Then
       data_lla2.Recordset.MoveLast
       bp1.Max = data_lla2.Recordset.RecordCount + 1
       Xlin = 1
       XCol = 1
       Xtotreg = 0
       Xsub = 0
       Set Xobjexel22 = New Excel.Application
       Set Xlibexel22 = Xobjexel22.Workbooks.Add
       Set Xarchexel22 = Xlibexel22.Worksheets.Add
       Xarchexel22.Name = Trim("Controles")
       Xlibexel22.SaveAs ("C:\planillas\Controles por medico.xls")
       Xarchtex = "C:\planillas\Controles por medico.xls"
    
       Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
       Xlin = Xlin + 1
       XCol = XCol + 1
       Xarchexel22.Range("A1", "C3").Font.Size = 16
       Xarchexel22.Range("A" & Trim(str(Xlin)), "T" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
       Xarchexel22.Cells(Xlin, XCol) = "CONTROLES SERVICIOS EN FIN DE SEMANA POR MÉDICO DESDE: " & md.Text & " HASTA: " & mh.Text
            
       XCol = 1
       Xlin = Xlin + 2
       Xnrocan = Xnrocan + Xlin
                  
       Xarchexel22.Range("A" & Trim(str(Xlin)), "V" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
       Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
       Xarchexel22.Cells(Xlin, XCol) = "COD.MED"
       XCol = XCol + 1
       Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 25
       Xarchexel22.Cells(Xlin, XCol) = "NOMBRE del MEDICO"
       XCol = XCol + 1
       Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 12
       Xarchexel22.Cells(Xlin, XCol) = "LLAMADOS"
       XCol = XCol + 1
       Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
       Xarchexel22.Cells(Xlin, XCol) = "POLIC."
       XCol = XCol + 1
       Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
       Xarchexel22.Cells(Xlin, XCol) = "CMT"
            
       Xlin = Xlin + 1
       XCol = 1
       data_lla2.Recordset.MoveFirst
       Xcodmed = data_lla2.Recordset("codmed")
       data_medhce.RecordSource = "select * from medicos where med_cod =" & data_lla2.Recordset("codmed")
       data_medhce.Refresh
       If data_medhce.Recordset.RecordCount > 0 Then
          Xcodmedhc = data_medhce.Recordset("med_socnro")
       Else
          Xcodmedhc = 0
       End If
       Do While Not data_lla2.Recordset.EOF
          If Xcodmed = data_lla2.Recordset("codmed") Then
             If Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 1 Or Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 7 Then
                If data_lla2.Recordset("movilpas") = 2015 Then
                   Xfinde = Xfinde + 1 'cmt finde
                Else
                   Xsub = Xsub + 1 'presencial finde
                End If
             End If
          Else
             data_lla2.Recordset.MovePrevious
             If Xsub > 0 Then
                If Xfinde > 0 Then
                   data_inflla.Recordset.AddNew
                   data_inflla.Recordset("codmed") = data_lla2.Recordset("codmed")
                   data_inflla.Recordset("nommed") = data_lla2.Recordset("nommed")
                   data_inflla.Recordset("matric") = Xsub 'llamados find
                   data_inflla.Recordset("edad") = Xfinde 'cmt finde
                   data_inflla.Recordset.Update
                Else
                   data_inflla.Recordset.AddNew
                   data_inflla.Recordset("codmed") = data_lla2.Recordset("codmed")
                   data_inflla.Recordset("nommed") = data_lla2.Recordset("nommed")
                   data_inflla.Recordset("matric") = Xsub 'llamados find
                   data_inflla.Recordset("edad") = 0 'cmt finde
                   data_inflla.Recordset.Update
                End If
             Else
                If Xfinde > 0 Then
                   data_inflla.Recordset.AddNew
                   data_inflla.Recordset("codmed") = data_lla2.Recordset("codmed")
                   data_inflla.Recordset("nommed") = data_lla2.Recordset("nommed")
                   data_inflla.Recordset("matric") = 0 'llamados find
                   data_inflla.Recordset("edad") = Xfinde 'cmt finde
                   data_inflla.Recordset.Update
                End If
             End If
             data_lla2.Recordset.MoveNext
             
             Xsub = 0
             Xfinde = 0
             If Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 1 Or Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 7 Then
                If data_lla2.Recordset("movilpas") = 2015 Then
                   Xfinde = Xfinde + 1
                Else
                   Xsub = Xsub + 1
                End If
             End If
          End If
          Xcodmed = data_lla2.Recordset("codmed")
          data_lla2.Recordset.MoveNext
          bp1.Value = bp1.Value + 1
       Loop
       data_lla2.Recordset.MovePrevious
       If Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 1 Or Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 7 Then
          If data_lla2.Recordset("movilpas") = 2015 Then
             Xfinde = Xfinde + 1 'cmt finde
          Else
             Xsub = Xsub + 1 'presencial finde
          End If
       End If
       If Xsub > 0 Then
          If Xfinde > 0 Then
             data_inflla.Recordset.AddNew
             data_inflla.Recordset("codmed") = data_lla2.Recordset("codmed")
             data_inflla.Recordset("nommed") = data_lla2.Recordset("nommed")
             data_inflla.Recordset("matric") = Xsub 'llamados find
             data_inflla.Recordset("edad") = Xfinde 'cmt finde
             data_inflla.Recordset.Update
          Else
             data_inflla.Recordset.AddNew
             data_inflla.Recordset("codmed") = data_lla2.Recordset("codmed")
             data_inflla.Recordset("nommed") = data_lla2.Recordset("nommed")
             data_inflla.Recordset("matric") = Xsub 'llamados find
             data_inflla.Recordset("edad") = 0 'cmt finde
             data_inflla.Recordset.Update
          End If
       Else
          If Xfinde > 0 Then
             data_inflla.Recordset.AddNew
             data_inflla.Recordset("codmed") = data_lla2.Recordset("codmed")
             data_inflla.Recordset("nommed") = data_lla2.Recordset("nommed")
             data_inflla.Recordset("matric") = 0 'llamados find
             data_inflla.Recordset("edad") = Xfinde 'cmt finde
             data_inflla.Recordset.Update
          End If
       End If
    End If

    If txt_med.Text <> 99 Then
       data_lla2.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and cod_prod in (10001,10003,10005) and nro_med_a=" & txt_med.Text & " order by nro_med_a"
    Else
       data_lla2.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and cod_prod in (10001,10003,10005) and nro_med_a not in (440,0,959,658,705,579,536,1694,1105,407,751,849,435,717,652,439,414,437,702,465,1231,566,662,1004,516,556,979,1541,668,403,515,560,693) order by nro_med_a"
    End If
    data_lla2.Refresh
    If data_lla2.Recordset.RecordCount > 0 Then
       data_lla2.Recordset.MoveLast
       bp1.Max = bp1.Value + data_lla2.Recordset.RecordCount + 1
       data_lla2.Recordset.MoveFirst
       Xcodmed = data_lla2.Recordset("nro_med_a")
       Do While Not data_lla2.Recordset.EOF
          If Xcodmed = data_lla2.Recordset("nro_med_a") Then
             If Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 1 Or Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 7 Then
                Xsub = Xsub + 1 'presencial finde polic
             End If
          Else
             data_lla2.Recordset.MovePrevious
             data_inflla.RecordSource = "select * from inflla where codmed =" & data_lla2.Recordset("nro_med_a")
             data_inflla.Refresh
             If data_inflla.Recordset.RecordCount > 0 Then
                data_inflla.Recordset.Edit
                data_inflla.Recordset("ci") = Xsub 'polic presenc
                data_inflla.Recordset.Update
             Else
                If Xsub > 0 Then
                   data_inflla.Recordset.AddNew
                   data_inflla.Recordset("codmed") = data_lla2.Recordset("nro_med_a")
                   data_inflla.Recordset("nommed") = data_lla2.Recordset("nom_med_a")
                   data_inflla.Recordset("ci") = Xsub
                   data_inflla.Recordset.Update
                End If
             End If
             data_lla2.Recordset.MoveNext
             Xsub = 0
             If Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 1 Or Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 7 Then
                Xsub = Xsub + 1
             End If
          End If
          Xcodmed = data_lla2.Recordset("nro_med_a")
          data_lla2.Recordset.MoveNext
          bp1.Value = bp1.Value + 1
       Loop
       data_lla2.Recordset.MovePrevious
       data_inflla.RecordSource = "select * from inflla where codmed =" & data_lla2.Recordset("nro_med_a")
       data_inflla.Refresh
       If data_inflla.Recordset.RecordCount > 0 Then
          data_inflla.Recordset.Edit
          data_inflla.Recordset("ci") = Xsub 'polic presenc
          data_inflla.Recordset.Update
       Else
          If Xsub > 0 Then
             data_inflla.Recordset.AddNew
             data_inflla.Recordset("codmed") = data_lla2.Recordset("nro_med_a")
             data_inflla.Recordset("nommed") = data_lla2.Recordset("nom_med_a")
             data_inflla.Recordset("ci") = Xsub
             data_inflla.Recordset.Update
          End If
       End If
       data_lla2.Recordset.MoveNext
       Xsub = 0
    End If
'cmt finde polic
    If txt_med.Text <> 99 Then
       data_lla2.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and cod_prod in (10050) and nro_med_a=" & txt_med.Text & " order by nro_med_a"
    Else
       data_lla2.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and cod_prod in (10050) and nro_med_a not in (440,0,959,658,705,579,536,1694,1105,407,751,849,435,717,652,439,414,437,702,465,1231,566,662,1004,516,556,979,1541,668,403,515,560,693) order by nro_med_a"
    End If
    data_lla2.Refresh
    If data_lla2.Recordset.RecordCount > 0 Then
       data_lla2.Recordset.MoveLast
       bp1.Max = bp1.Value + data_lla2.Recordset.RecordCount + 1
       data_lla2.Recordset.MoveFirst
       Xcodmed = data_lla2.Recordset("nro_med_a")
       Do While Not data_lla2.Recordset.EOF
          If Xcodmed = data_lla2.Recordset("nro_med_a") Then
             If Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 1 Or Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 7 Then
                Xsub = Xsub + 1 'presencial finde polic
             End If
          Else
             data_lla2.Recordset.MovePrevious
             data_inflla.RecordSource = "select * from inflla where codmed =" & data_lla2.Recordset("nro_med_a")
             data_inflla.Refresh
             If data_inflla.Recordset.RecordCount > 0 Then
                data_inflla.Recordset.Edit
                data_inflla.Recordset("edad") = data_inflla.Recordset("edad") + Xsub 'polic presenc
                data_inflla.Recordset.Update
             Else
                If Xsub > 0 Then
                   data_inflla.Recordset.AddNew
                   data_inflla.Recordset("codmed") = data_lla2.Recordset("nro_med_a")
                   data_inflla.Recordset("nommed") = data_lla2.Recordset("nom_med_a")
                   data_inflla.Recordset("edad") = Xsub
                   data_inflla.Recordset.Update
                End If
             End If
             data_lla2.Recordset.MoveNext
             Xsub = 0
             If Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 1 Or Weekday(DateSerial(Year(data_lla2.Recordset("fecha")), Month(data_lla2.Recordset("fecha")), Day(data_lla2.Recordset("fecha")))) = 7 Then
                Xsub = Xsub + 1
             End If
          End If
          Xcodmed = data_lla2.Recordset("nro_med_a")
          data_lla2.Recordset.MoveNext
          bp1.Value = bp1.Value + 1
       Loop
       data_lla2.Recordset.MovePrevious
       data_inflla.RecordSource = "select * from inflla where codmed =" & data_lla2.Recordset("nro_med_a")
       data_inflla.Refresh
       If data_inflla.Recordset.RecordCount > 0 Then
          data_inflla.Recordset.Edit
          data_inflla.Recordset("edad") = data_inflla.Recordset("edad") + Xsub 'polic presenc
          data_inflla.Recordset.Update
       Else
          If Xsub > 0 Then
             data_inflla.Recordset.AddNew
             data_inflla.Recordset("codmed") = data_lla2.Recordset("nro_med_a")
             data_inflla.Recordset("nommed") = data_lla2.Recordset("nom_med_a")
             data_inflla.Recordset("edad") = Xsub
             data_inflla.Recordset.Update
          End If
       End If
    End If
    data_inflla.RecordSource = "Select * from inflla"
    data_inflla.Refresh
    XCol = 1
    If data_inflla.Recordset.RecordCount > 0 Then
       data_inflla.Recordset.MoveFirst
       Do While Not data_inflla.Recordset.EOF
          Xarchexel22.Cells(Xlin, XCol) = data_inflla.Recordset("codmed")
          XCol = XCol + 1
          Xarchexel22.Cells(Xlin, XCol) = data_inflla.Recordset("nommed")
          XCol = XCol + 1
          Xarchexel22.Cells(Xlin, XCol) = data_inflla.Recordset("matric")
          XCol = XCol + 1
          Xarchexel22.Cells(Xlin, XCol) = data_inflla.Recordset("ci")
          XCol = XCol + 1
          Xarchexel22.Cells(Xlin, XCol) = data_inflla.Recordset("edad")
          Xlin = Xlin + 1
          XCol = 1
          data_inflla.Recordset.MoveNext
       Loop
    End If
       
    frm_prodmed.MousePointer = 0
    Xlin = Xlin + 1
    XCol = 1
    Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")
    Xlibexel22.Save
    Xlibexel22.Close
    Xobjexel22.Quit
    
    Xlabrir.Workbooks.Open Xarchtex, , False
    Xlabrir.Visible = True
    Xlabrir.WindowState = xlMaximized
Else
    If data_lla2.Recordset.RecordCount > 0 Then



    End If
End If
frm_prodmed.MousePointer = 0



End Sub

Private Sub Form_Load()
data_med.DatabaseName = App.path & "\medicos.mdb"
data_med.RecordSource = "medicos"
data_med.Refresh
'data_hs.DatabaseName = App.Path & "\medconhoras.mdb"
'data_hs.RecordSource = "tabmed"
'data_hs.Refresh
data_lla.ConnectionString = "dsn=" & Xconexrmt
'data_lla.RecordSource = "llamado"
'data_lla.Refresh
data_inflla.DatabaseName = App.path & "\informes.mdb"
'data_inflla.RecordSource = "inflla"
'data_inflla.Refresh
Combo1.ListIndex = 0

Data1.ConnectionString = "dsn=" & Xconexrmt
data_lla2.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_hce.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_medhce.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_llamsp.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_llaresp.Connect = "odbc;dsn=" & Xconexrmt & ";"


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
    txt_med.SetFocus
End If

End Sub

Private Sub mhd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhh.SetFocus
End If

End Sub

Private Sub mhh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_med.SetFocus
End If

End Sub

Private Sub txt_med_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_proc.SetFocus
End If

End Sub
