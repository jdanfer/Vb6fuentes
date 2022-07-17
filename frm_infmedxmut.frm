VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infmedxmut 
   BackColor       =   &H0080FF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes de medicación por mutualista"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infmedxmut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_lin 
      Height          =   375
      Left            =   1320
      Top             =   3840
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
      Caption         =   "data_lin"
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
      Left            =   3480
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
      Height          =   495
      Left            =   4320
      Picture         =   "frm_infmedxmut.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Salir"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_infmedxmut.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Procesar"
      Top             =   3840
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Datos del informe"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSAdodcLib.Adodc data_cli 
         Height          =   375
         Left            =   240
         Top             =   1920
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
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080FF80&
         Caption         =   "Generar planilla electrónica (EXCEL)"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Value           =   1  'Checked
         Width           =   4335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FF80&
         Caption         =   "Resumen"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   2520
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080FF80&
         Caption         =   "Detalle"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.TextBox t_base 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Text            =   "99"
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_infmedxmut.frx":109E
         Left            =   1320
         List            =   "frm_infmedxmut.frx":10B4
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   3015
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   0
         X2              =   4680
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FFFF&
         Caption         =   "99 = Todas"
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BASE:"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mutualista:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha:"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label labd 
      Caption         =   "Label5"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label labm 
      Caption         =   "Label5"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label laba 
      Caption         =   "Label5"
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1080
      Picture         =   "frm_infmedxmut.frx":10F3
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1815
   End
End
Attribute VB_Name = "frm_infmedxmut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_base.SetFocus
End If

End Sub

Private Sub Command1_Click()
Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet
Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xarchtex As String
Dim Xlabrir As New Excel.Application

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
MiBaseact.Execute "Delete * from infcli"
laba.Caption = ""
labm.Caption = ""
labd.Caption = ""
XCol = 1
Xlin = 1
Xnrocan = 1
frm_infmedxmut.MousePointer = 11
Command1.Enabled = False
If mfd.Text <> "__/__/____" And mfh.Text <> "__/__/____" Then
   
   data_inf.RecordSource = "infcli"
   data_inf.Refresh
   If t_base.Text = 99 Then
      If Combo1.ListIndex = 0 Then
         data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.ced_socio,linmmdd.fact,linmmdd.nom_medic,linmmdd.cod_prod," & _
         "linmmdd.zona,linmmdd.base,linmmdd.convenio,linmmdd.nro_flia,linmmdd.tot_lin,clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on " & _
         "linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and linmmdd.nro_flia =" & 6 & " order by linmmdd.fecha"
      Else
         If Combo1.ListIndex = 1 Then
            data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.ced_socio,linmmdd.fact,linmmdd.nom_medic,linmmdd.cod_prod," & _
            "linmmdd.zona,linmmdd.base,linmmdd.convenio,linmmdd.nro_flia,linmmdd.tot_lin,clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on " & _
            "linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 60103 & " order by linmmdd.fecha"
'            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 60103 & " order by fecha"
         Else
            If Combo1.ListIndex = 2 Then
               data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.ced_socio,linmmdd.fact,linmmdd.nom_medic,linmmdd.cod_prod," & _
               "linmmdd.zona,linmmdd.base,linmmdd.convenio,linmmdd.nro_flia,linmmdd.tot_lin,clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on " & _
               "linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 60108 & " order by linmmdd.fecha"
'               data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 60108 & " order by fecha"
            Else
               If Combo1.ListIndex = 3 Then
                  data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.ced_socio,linmmdd.fact,linmmdd.nom_medic,linmmdd.cod_prod," & _
                 "linmmdd.zona,linmmdd.base,linmmdd.convenio,linmmdd.nro_flia,linmmdd.tot_lin,clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on " & _
                 "linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 60107 & " order by linmmdd.fecha"
'                  data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 60107 & " order by fecha"
               Else
                  If Combo1.ListIndex = 4 Then
                     data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.ced_socio,linmmdd.fact,linmmdd.nom_medic,linmmdd.cod_prod," & _
                     "linmmdd.zona,linmmdd.base,linmmdd.convenio,linmmdd.nro_flia,linmmdd.tot_lin,clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on " & _
                     "linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (60105,60106) order by linmmdd.fecha"
'                     data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (60105,60106) order by fecha"
                  Else
                     If Combo1.ListIndex = 5 Then
                        data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.ced_socio,linmmdd.fact,linmmdd.nom_medic,linmmdd.cod_prod," & _
                        "linmmdd.zona,linmmdd.base,linmmdd.convenio,linmmdd.nro_flia,linmmdd.tot_lin,clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on " & _
                        "linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (60109) order by linmmdd.fecha"
'                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 60109 & " order by fecha"
                     Else
                        data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.ced_socio,linmmdd.fact,linmmdd.nom_medic,linmmdd.cod_prod," & _
                        "linmmdd.zona,linmmdd.base,linmmdd.convenio,linmmdd.nro_flia,linmmdd.tot_lin,clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on " & _
                        "linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia in (6) order by linmmdd.fecha"
'                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia =" & 6
                     End If
                  End If
               End If
            End If
         End If
      End If
   Else
      If Combo1.ListIndex = 0 Then
         data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.ced_socio,linmmdd.fact,linmmdd.nom_medic,linmmdd.cod_prod," & _
         "linmmdd.zona,linmmdd.base,linmmdd.convenio,linmmdd.nro_flia,linmmdd.tot_lin,clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on " & _
         "linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia in (6) and base =" & t_base.Text & " order by linmmdd.fecha"
'         data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and base =" & t_base.Text & " order by fecha"
      Else
         If Combo1.ListIndex = 1 Then
            data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.ced_socio,linmmdd.fact,linmmdd.nom_medic,linmmdd.cod_prod," & _
            "linmmdd.zona,linmmdd.base,linmmdd.convenio,linmmdd.nro_flia,linmmdd.tot_lin,clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on " & _
            "linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (60103) and base =" & t_base.Text & " order by linmmdd.fecha"
'            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 60103 & " and base =" & t_base.Text & " order by fecha"
         Else
            If Combo1.ListIndex = 2 Then
               data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.ced_socio,linmmdd.fact,linmmdd.nom_medic,linmmdd.cod_prod," & _
               "linmmdd.zona,linmmdd.base,linmmdd.convenio,linmmdd.nro_flia,linmmdd.tot_lin,clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on " & _
               "linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (60108) and base =" & t_base.Text & " order by linmmdd.fecha"
'               data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 60108 & " and base =" & t_base.Text & " order by fecha"
            Else
               If Combo1.ListIndex = 3 Then
                  data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.ced_socio,linmmdd.fact,linmmdd.nom_medic,linmmdd.cod_prod," & _
                  "linmmdd.zona,linmmdd.base,linmmdd.convenio,linmmdd.nro_flia,linmmdd.tot_lin,clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on " & _
                  "linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (60107) and base =" & t_base.Text & " order by linmmdd.fecha"
'                  data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 60107 & " and base =" & t_base.Text & " order by fecha"
               Else
                  If Combo1.ListIndex = 4 Then
                     data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.ced_socio,linmmdd.fact,linmmdd.nom_medic,linmmdd.cod_prod," & _
                     "linmmdd.zona,linmmdd.base,linmmdd.convenio,linmmdd.nro_flia,linmmdd.tot_lin,clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on " & _
                     "linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (60105,60106) and base =" & t_base.Text & " order by linmmdd.fecha"
'                     data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (60105,60106) and base =" & t_base.Text & " order by fecha"
                  Else
                     If Combo1.ListIndex = 5 Then
                        data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.ced_socio,linmmdd.fact,linmmdd.nom_medic,linmmdd.cod_prod," & _
                        "linmmdd.zona,linmmdd.base,linmmdd.convenio,linmmdd.nro_flia,linmmdd.tot_lin,clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on " & _
                        "linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod in (60109) and base =" & t_base.Text & " order by linmmdd.fecha"
'                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 60109 & " and base =" & t_base.Text & " order by fecha"
                     Else
                        data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.ced_socio,linmmdd.fact,linmmdd.nom_medic,linmmdd.cod_prod," & _
                        "linmmdd.zona,linmmdd.base,linmmdd.convenio,linmmdd.nro_flia,linmmdd.tot_lin,clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on " & _
                        "linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia in (6) and base =" & t_base.Text & " order by linmmdd.fecha"
'                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and base =" & t_base.Text
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
   data_lin.Refresh
   Dim Xnomsin As String
   If data_lin.Recordset.RecordCount > 0 Then
      data_lin.Recordset.MoveFirst
      Do While Not data_lin.Recordset.EOF
         Xnomsin = Replace(data_lin.Recordset("nom_cli"), "'", chr(37))
         If Combo1.Text = "UNIVERSAL" Then
            MiBaseact.Execute "Insert into infcli (cl_fnac,cl_fecing,cl_apellid,cl_codigo,cl_telefon,cl_direcci,cl_entre," & _
            "cl_nomcobr,cl_codced,cl_codconv,cl_cedula,cl_zona" & _
            ") values ('" & data_lin.Recordset("fecha") & "','" & data_lin.Recordset("cl_fnac") & "','" & Xnomsin & "'," & _
            data_lin.Recordset("cod_cli") & ",'" & Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact"))) & "'," & _
            "'" & data_lin.Recordset("nom_medic") & "','" & data_lin.Recordset("zona") & "','" & Combo1.Text & "'," & _
            data_lin.Recordset("base") & ",'" & data_lin.Recordset("convenio") & "'," & data_lin.Recordset("tot_lin") & "," & _
            "'" & Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact"))) & "')"
         Else
            MiBaseact.Execute "Insert into infcli (cl_fnac,cl_fecing,cl_apellid,cl_codigo,cl_telefon,cl_direcci,cl_entre," & _
            "cl_nomcobr,cl_codced,cl_codconv,cl_cedula,cl_zona" & _
            ") values ('" & data_lin.Recordset("fecha") & "','" & data_lin.Recordset("cl_fnac") & "','" & Xnomsin & "'," & _
            data_lin.Recordset("cod_cli") & ",'" & Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact"))) & "'," & _
            "'" & data_lin.Recordset("zona") & "','" & data_lin.Recordset("zona") & "','" & Combo1.Text & "'," & _
            data_lin.Recordset("base") & ",'" & data_lin.Recordset("convenio") & "'," & data_lin.Recordset("tot_lin") & "," & _
            "'" & Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact"))) & "')"
         End If
         
         data_lin.Recordset.MoveNext
         Xnrocan = Xnrocan + 1
      Loop
      data_inf.RecordSource = "Select * from infcli order by cl_direcci"
      data_inf.Refresh
      data_inf.Recordset.MoveFirst
      If Check1.Value = 1 Then
         If Combo1.Text = "UNIVERSAL" Then
             Set Xobjexel = New Excel.Application
             Set Xlibexel = Xobjexel.Workbooks.Add
             Set Xarchexel = Xlibexel.Worksheets.Add
             Xarchexel.Name = Trim(Combo1.Text)
             Xlibexel.SaveAs ("C:\planillas\" & Trim(Combo1.Text) & ".xls")
             Xarchtex = "C:\planillas\" & Trim(Combo1.Text) & ".xls"
             Xarchexel.Range("A1", "C5").Font.Size = 16
             Xarchexel.Range("A1", "C5").Font.Bold = True
             Xlin = Xlin + 1
             XCol = XCol + 1
             Xarchexel.Cells(Xlin, XCol) = "PEDIDO DE SEDE SECUNDARIA COLECTIVA"
             Xlin = Xlin + 1
             'XCol = XCol + 1
             Xarchexel.Cells(Xlin, XCol) = "MUTUALISTA:"
             XCol = XCol + 1
             Xarchexel.Cells(Xlin, XCol) = Combo1.Text
             Xlin = Xlin + 1
             XCol = 2
             Xarchexel.Cells(Xlin, XCol) = "FECHAS COMPRENDIDAS"
             XCol = XCol + 1
             Xarchexel.Cells(Xlin, XCol) = Format(mfd.Text, "dd/mm/yyyy") & " A: " & Format(mfh.Text, "dd/mm/yyyy")
             XCol = 2
             Xlin = Xlin + 1
             Xarchexel.Cells(Xlin, XCol) = "NUMERO DE PEDIDO SAPP"
             XCol = XCol + 1
'             Xarchexel.Range("A1", "C3").Font.Size = 16
'             Xarchexel.Cells(Xlin, XCol) = "MEDICACION ENTREGADA " & Trim(Combo1.Text) & " DESDE: " & mfd.Text & " HASTA: " & mfh.Text
'             Xarchexel.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 200)
             XCol = 1
             Xlin = Xlin + 2
             Xnrocan = Xnrocan + Xlin
             Xarchexel.Range("A8", "L" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
             Xarchexel.Range("A8", "L" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
             Xarchexel.Range("A8", "L" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
             Xarchexel.Range("A8", "L" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
             Xarchexel.Range("A8", "L" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
             Xarchexel.Range("A8", "L" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
             Xarchexel.Range("A" & Trim(str(Xlin)), "L" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
             Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
             Xarchexel.Cells(Xlin, XCol) = "NRO."
             XCol = XCol + 1
             Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 15
             Xarchexel.Cells(Xlin, XCol) = "ZONA"
             XCol = XCol + 1
             Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 40
             Xarchexel.Cells(Xlin, XCol) = "NOMBRE"
             XCol = XCol + 1
             Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
             Xarchexel.Cells(Xlin, XCol) = "DOCUMENTO"
             XCol = XCol + 1
             Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 40
             Xarchexel.Cells(Xlin, XCol) = "MEDICAMENTO"
             XCol = XCol + 1
             Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
             Xarchexel.Cells(Xlin, XCol) = "FECHA"
             XCol = XCol + 1
             Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
             Xarchexel.Cells(Xlin, XCol) = "PSICO"
             XCol = XCol + 1
             Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
             Xarchexel.Cells(Xlin, XCol) = "C/DUPLICADO"
             XCol = XCol + 1
             Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 8
             Xarchexel.Cells(Xlin, XCol) = "VALES"
             XCol = XCol + 1
             Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 8
             Xarchexel.Cells(Xlin, XCol) = "AÑOS"
             XCol = XCol + 1
             Xarchexel.Range("K" & Trim(str(Xlin))).ColumnWidth = 8
             Xarchexel.Cells(Xlin, XCol) = "MESES"
             XCol = XCol + 1
             Xarchexel.Range("L" & Trim(str(Xlin))).ColumnWidth = 8
             Xarchexel.Cells(Xlin, XCol) = "DIAS"
             
             Xlin = Xlin + 1
             XCol = 1
             Xnrocan = 1
             If data_inf.Recordset.RecordCount > 0 Then
                data_inf.Recordset.MoveLast
                data_inf.Recordset.MoveFirst
                Do While Not data_inf.Recordset.EOF
                   If IsNull(data_inf.Recordset("cl_fecing")) = False Then
                      CalculaEdad (data_inf.Recordset("cl_fecing"))
                   Else
                      laba.Caption = 0
                      labm.Caption = 0
                      labd.Caption = 0
                   End If
                   Xarchexel.Cells(Xlin, XCol) = Xnrocan
                   XCol = XCol + 1
                   If data_inf.Recordset("cl_codced") = 1 Or _
                      data_inf.Recordset("cl_codced") = 2 Or _
                      data_inf.Recordset("cl_codced") = 3 Or _
                      data_inf.Recordset("cl_codced") = 4 Or _
                      data_inf.Recordset("cl_codced") = 18 Or _
                      data_inf.Recordset("cl_codced") = 92 Then
                      Xarchexel.Cells(Xlin, XCol) = "SALINAS"
                   Else
                      If data_inf.Recordset("cl_codced") = 6 Or _
                         data_inf.Recordset("cl_codced") = 17 Then
                         Xarchexel.Cells(Xlin, XCol) = "BS.BS."
                      Else
                         Xarchexel.Cells(Xlin, XCol) = "TOLEDO"
                      End If
                   End If
                   XCol = XCol + 1
                   Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
                   XCol = XCol + 1
                   Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon") & " CI"
                   XCol = XCol + 1
                   Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
                   XCol = XCol + 1
                   Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
                   XCol = XCol + 1
                   Xarchexel.Cells(Xlin, XCol) = "COMUN"
                   XCol = XCol + 1
                   Xarchexel.Cells(Xlin, XCol) = laba.Caption
                   XCol = XCol + 1
                   Xarchexel.Cells(Xlin, XCol) = labm.Caption
                   XCol = XCol + 1
                   Xarchexel.Cells(Xlin, XCol) = labd.Caption
                   data_inf.Recordset.MoveNext
                   Xlin = Xlin + 1
                   Xcolfija = XCol
                   XCol = 1
                   Xnrocan = Xnrocan + 1
                Loop
    '            Xlin = Xlin - 1
    '            Xarchexel.Cells(Xlin, Xcolfija) = Xcanxdia
                Xlibexel.Save
        '            Xlibexel.Application
                Xlibexel.Close
                Xobjexel.Quit
                frm_infmedxmut.MousePointer = 0
                Command1.Enabled = True
'                Shell frm_menu.data_usuac.Recordset("destino") & "excel.exe " & Xarchtex, vbMaximizedFocus
                Xlabrir.Workbooks.Open Xarchtex, , False
                Xlabrir.Visible = True
                Xlabrir.WindowState = xlMaximized
                
             End If
         Else
             If Combo1.Text = "C.GALICIA" Then
                 Set Xobjexel = New Excel.Application
                 Set Xlibexel = Xobjexel.Workbooks.Add
                 Set Xarchexel = Xlibexel.Worksheets.Add
                 Xarchexel.Name = Trim(Combo1.Text)
                 Xlibexel.SaveAs ("C:\planillas\" & Trim(Combo1.Text) & ".xls")
                 Xarchtex = "C:\planillas\" & Trim(Combo1.Text) & ".xls"
                 Xarchexel.Range("A1", "C5").Font.Size = 16
                 Xarchexel.Range("A1", "C5").Font.Bold = True
                 Xlin = Xlin + 1
                 XCol = XCol + 1
                 Xarchexel.Cells(Xlin, XCol) = "PEDIDO DE SEDE SECUNDARIA COLECTIVA"
                 Xlin = Xlin + 1
                 'XCol = XCol + 1
                 Xarchexel.Cells(Xlin, XCol) = "MUTUALISTA:"
                 XCol = XCol + 1
                 Xarchexel.Cells(Xlin, XCol) = Combo1.Text
                 Xlin = Xlin + 1
                 XCol = 2
                 Xarchexel.Cells(Xlin, XCol) = "FECHAS COMPRENDIDAS"
                 XCol = XCol + 1
                 Xarchexel.Cells(Xlin, XCol) = Format(mfd.Text, "dd/mm/yyyy") & " A: " & Format(mfh.Text, "dd/mm/yyyy")
                 XCol = 2
                 Xlin = Xlin + 1
                 Xarchexel.Cells(Xlin, XCol) = "NUMERO DE PEDIDO SAPP"
                 XCol = XCol + 1
                 XCol = 1
                 Xlin = Xlin + 2
                 Xnrocan = Xnrocan + Xlin
                 Xarchexel.Range("A8", "L" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                 Xarchexel.Range("A8", "L" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
                 Xarchexel.Range("A8", "L" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
                 Xarchexel.Range("A8", "L" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
                 Xarchexel.Range("A8", "L" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
                 Xarchexel.Range("A8", "L" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
                 Xarchexel.Range("A" & Trim(str(Xlin)), "L" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
                 Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
'                 Xarchexel.Cells(Xlin, XCol) = "NRO."
 '                XCol = XCol + 1
'                 Xarchexel.Range("A" & Trim(Str(Xlin))).ColumnWidth = 10
                 Xarchexel.Cells(Xlin, XCol) = "ZONA"
                 XCol = XCol + 1
                 Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
                 Xarchexel.Cells(Xlin, XCol) = "NOMBRE"
                 XCol = XCol + 1
                 Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
                 Xarchexel.Cells(Xlin, XCol) = "Nro.HC"
                 XCol = XCol + 1
                 Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
                 Xarchexel.Cells(Xlin, XCol) = "DOCUMENTO"
                 XCol = XCol + 1
                 Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 35
                 Xarchexel.Cells(Xlin, XCol) = "Fármaco solicitado"
                 XCol = XCol + 1
                 Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
                 Xarchexel.Cells(Xlin, XCol) = "Nro.recetas"
                 Xlin = Xlin + 1
                 Xarchexel.Cells(Xlin, XCol) = "q se envían"
                 Xlin = Xlin - 1
                 XCol = XCol + 1
                 Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
                 Xarchexel.Cells(Xlin, XCol) = "Cant.Solic."
                 XCol = XCol + 1
                 Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
                 Xarchexel.Cells(Xlin, XCol) = "Cant.enviada"
                 XCol = XCol + 1
                 Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 10
                 Xarchexel.Cells(Xlin, XCol) = "No.RECETA"
                 XCol = XCol + 1
                 Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 8
                 Xarchexel.Cells(Xlin, XCol) = "AÑOS"
                 XCol = XCol + 1
                 Xarchexel.Range("K" & Trim(str(Xlin))).ColumnWidth = 8
                 Xarchexel.Cells(Xlin, XCol) = "MESES"
                 XCol = XCol + 1
                 Xarchexel.Range("L" & Trim(str(Xlin))).ColumnWidth = 8
                 Xarchexel.Cells(Xlin, XCol) = "DIAS"
                 Xlin = Xlin + 2
                 XCol = 1
                 Xnrocan = 1
                 If data_inf.Recordset.RecordCount > 0 Then
                    data_inf.Recordset.MoveLast
                    data_inf.Recordset.MoveFirst
                    Do While Not data_inf.Recordset.EOF
'                       Xarchexel.Cells(Xlin, XCol) = Xnrocan
 '                      XCol = XCol + 1
                       If IsNull(data_inf.Recordset("cl_fecing")) = False Then
                          CalculaEdad (data_inf.Recordset("cl_fecing"))
                       Else
                          laba.Caption = 0
                          labm.Caption = 0
                          labd.Caption = 0
                       End If
                       If data_inf.Recordset("cl_codced") = 1 Or _
                          data_inf.Recordset("cl_codced") = 2 Or _
                          data_inf.Recordset("cl_codced") = 3 Or _
                          data_inf.Recordset("cl_codced") = 4 Or _
                          data_inf.Recordset("cl_codced") = 18 Or _
                          data_inf.Recordset("cl_codced") = 92 Then
                          Xarchexel.Cells(Xlin, XCol) = "SALINAS"
                       Else
                          If data_inf.Recordset("cl_codced") = 6 Or _
                             data_inf.Recordset("cl_codced") = 17 Then
                             Xarchexel.Cells(Xlin, XCol) = "BS.BS."
                          Else
                             Xarchexel.Cells(Xlin, XCol) = "TOLEDO"
                          End If
                       End If
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nrosocm")
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon") & " CI"
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = 1
                       data_inf.Recordset.MoveNext
                       Xlin = Xlin + 1
                       Xcolfija = XCol
                       XCol = 1
                       Xnrocan = Xnrocan + 1
                    Loop
        '            Xlin = Xlin - 1
        '            Xarchexel.Cells(Xlin, Xcolfija) = Xcanxdia
                    Xlibexel.Save
            '            Xlibexel.Application
                    Xlibexel.Close
                    Xobjexel.Quit
                    frm_infmedxmut.MousePointer = 0
                    Command1.Enabled = True
'                    Shell frm_menu.data_usuac.Recordset("destino") & "excel.exe " & Xarchtex, vbMaximizedFocus
                    Xlabrir.Workbooks.Open Xarchtex, , False
                    Xlabrir.Visible = True
                    Xlabrir.WindowState = xlMaximized
                 End If
             Else
                 Set Xobjexel = New Excel.Application
                 Set Xlibexel = Xobjexel.Workbooks.Add
                 Set Xarchexel = Xlibexel.Worksheets.Add
                 Xarchexel.Name = Trim(Combo1.Text)
                 Xlibexel.SaveAs ("C:\planillas\" & Trim(Combo1.Text) & ".xls")
                 Xarchtex = "C:\planillas\" & Trim(Combo1.Text) & ".xls"
                 Xarchexel.Range("A1", "C5").Font.Size = 16
                 Xarchexel.Range("A1", "C5").Font.Bold = True
                 Xlin = Xlin + 1
                 XCol = XCol + 1
                 Xarchexel.Cells(Xlin, XCol) = "PEDIDO DE SEDE SECUNDARIA COLECTIVA"
                 Xlin = Xlin + 1
                 'XCol = XCol + 1
                 Xarchexel.Cells(Xlin, XCol) = "MUTUALISTA:"
                 XCol = XCol + 1
                 Xarchexel.Cells(Xlin, XCol) = Combo1.Text
                 Xlin = Xlin + 1
                 XCol = 2
                 Xarchexel.Cells(Xlin, XCol) = "FECHAS COMPRENDIDAS"
                 XCol = XCol + 1
                 Xarchexel.Cells(Xlin, XCol) = Format(mfd.Text, "dd/mm/yyyy") & " A: " & Format(mfh.Text, "dd/mm/yyyy")
                 XCol = 2
                 Xlin = Xlin + 1
                 Xarchexel.Cells(Xlin, XCol) = "NUMERO DE PEDIDO SAPP"
                 XCol = XCol + 1
                 XCol = 1
                 Xlin = Xlin + 2
                 Xnrocan = Xnrocan + Xlin
                 Xarchexel.Range("A8", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                 Xarchexel.Range("A8", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
                 Xarchexel.Range("A8", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
                 Xarchexel.Range("A8", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
                 Xarchexel.Range("A8", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
                 Xarchexel.Range("A8", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
                 Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
                 Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
                 Xarchexel.Cells(Xlin, XCol) = "NRO."
                 XCol = XCol + 1
                 Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 15
                 Xarchexel.Cells(Xlin, XCol) = "ZONA"
                 XCol = XCol + 1
                 Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 40
                 Xarchexel.Cells(Xlin, XCol) = "NOMBRE"
                 XCol = XCol + 1
                 Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
                 Xarchexel.Cells(Xlin, XCol) = "DOCUMENTO"
                 XCol = XCol + 1
                 Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 40
                 Xarchexel.Cells(Xlin, XCol) = "MEDICAMENTO"
                 XCol = XCol + 1
                 Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
                 Xarchexel.Cells(Xlin, XCol) = "OBSERVACION"
                 XCol = XCol + 1
                 Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 8
                 Xarchexel.Cells(Xlin, XCol) = "AÑOS"
                 XCol = XCol + 1
                 Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 8
                 Xarchexel.Cells(Xlin, XCol) = "MESES"
                 XCol = XCol + 1
                 Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 8
                 Xarchexel.Cells(Xlin, XCol) = "DIAS"
                        
                 Xlin = Xlin + 1
                 XCol = 1
                 Xnrocan = 1
                 If data_inf.Recordset.RecordCount > 0 Then
                    data_inf.Recordset.MoveLast
                    data_inf.Recordset.MoveFirst
                    Do While Not data_inf.Recordset.EOF
                       If IsNull(data_inf.Recordset("cl_fecing")) = False Then
                          CalculaEdad (data_inf.Recordset("cl_fecing"))
                       Else
                          laba.Caption = 0
                          labm.Caption = 0
                          labd.Caption = 0
                       End If
                       Xarchexel.Cells(Xlin, XCol) = Xnrocan
                       XCol = XCol + 1
                       If data_inf.Recordset("cl_codced") = 1 Or _
                          data_inf.Recordset("cl_codced") = 2 Or _
                          data_inf.Recordset("cl_codced") = 3 Or _
                          data_inf.Recordset("cl_codced") = 4 Or _
                          data_inf.Recordset("cl_codced") = 18 Or _
                          data_inf.Recordset("cl_codced") = 92 Then
                          Xarchexel.Cells(Xlin, XCol) = "SALINAS"
                       Else
                          If data_inf.Recordset("cl_codced") = 6 Or _
                             data_inf.Recordset("cl_codced") = 17 Then
                             Xarchexel.Cells(Xlin, XCol) = "BS.BS."
                          Else
                             Xarchexel.Cells(Xlin, XCol) = "TOLEDO"
                          End If
                       End If
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon") & " CI"
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = "COMUN"
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = laba.Caption
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = labm.Caption
                       XCol = XCol + 1
                       Xarchexel.Cells(Xlin, XCol) = labd.Caption
                       data_inf.Recordset.MoveNext
                       Xlin = Xlin + 1
                       Xcolfija = XCol
                       XCol = 1
                       Xnrocan = Xnrocan + 1
                    Loop
        '            Xlin = Xlin - 1
        '            Xarchexel.Cells(Xlin, Xcolfija) = Xcanxdia
                    Xlibexel.Save
            '            Xlibexel.Application
                    Xlibexel.Close
                    Xobjexel.Quit
                    frm_infmedxmut.MousePointer = 0
                    Command1.Enabled = True
'                    Shell frm_menu.data_usuac.Recordset("destino") & "excel.exe " & Xarchtex, vbMaximizedFocus
                    Xlabrir.Workbooks.Open Xarchtex, , False
                    Xlabrir.Visible = True
                    Xlabrir.WindowState = xlMaximized
                 End If
             End If
         End If
      Else
         If data_inf.Recordset.RecordCount > 0 Then
            data_inf.Recordset.MoveFirst
            Xnrocan = 1
            Do While Not data_inf.Recordset.EOF
               data_inf.Recordset.Edit
               data_inf.Recordset("cl_nrovend") = Xnrocan
               data_inf.Recordset.Update
               data_inf.Recordset.MoveNext
               Xnrocan = Xnrocan + 1
            Loop
            data_inf.Recordset.MoveFirst
         End If
         frm_infmedxmut.MousePointer = 0
         Command1.Enabled = True
         If Option1.Value = True Then
            cr1.ReportFileName = App.path & "\infmedxmut.rpt"
            cr1.ReportTitle = "Informe medicación mutualista: " & Combo1.Text & " DESDE:" & mfd.Text & " HASTA:" & mfh.Text
            cr1.Action = 1
         Else
            cr1.ReportFileName = App.path & "\infmedxmutn.rpt"
            cr1.ReportTitle = "Informe medicación mutualista: " & Combo1.Text & " DESDE:" & mfd.Text & " HASTA:" & mfh.Text
            cr1.Action = 1
         End If
      End If
   Else
      frm_infmedxmut.MousePointer = 0
      Command1.Enabled = True
      MsgBox "No hay registros"
   End If
Else
   frm_infmedxmut.MousePointer = 0
   Command1.Enabled = True
   MsgBox "No ingresó fechas"
End If
frm_infmedxmut.MousePointer = 0
Command1.Enabled = True

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
'data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lin.ConnectionString = "dsn=" & Xconexrmt
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"

data_cli.ConnectionString = "dsn=" & Xconexrmt
'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"

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

Private Sub t_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub
Private Sub CalculaEdad(ByVal FNaci As Date)
Dim FAct As String
Dim Anios As String
Dim Meses As String
Dim Dias As String
Dim newday As String
Dim newmonth As String
Dim newyear As String

FAct = Format(Now, "dd/MM/yyyy")
FNaci = Format(FNaci, "dd/MM/yyyy")

'Calcula los años
Anios = DateDiff("yyyy", CDate(Format(FNaci, "dd/MM/yyyy")), data_inf.Recordset("cl_fnac"))
'Si el mes actual es menor que el mes de la fecha de nacimiento entonces
If Month(CDate(FAct)) < Month(CDate(FNaci)) Then
 'Restele uno a los años
 Anios = Anios - 1
 newmonth = Month(CDate(FAct)) + 12
 Else
 'Deje el mes actual tal y como estan
 newmonth = Month(CDate(FAct))
 End If

 'Si el mes actual es igual al mes de la fecha de nacimiento entonces
If Month(CDate(FAct)) = Month(CDate(FNaci)) Then
 'Si el día de la fecha actual es menor al día de la fecha de nacimiento
 If Day(CDate(FAct)) < Day(CDate(FNaci)) Then
 'Restele uno a los años
 Anios = Anios - 1
 End If
End If

If Day(CDate(FAct)) < Day(CDate(FNaci)) Then

   If Month(FNaci) = 1 Or Month(FNaci) = 3 Or Month(FNaci) = 5 Or _
      Month(FNaci) = 7 Or Month(FNaci) = 8 Or Month(FNaci) = 10 Or _
      Month(FNaci) = 12 Then
      newday = Day(CDate(FAct)) + 31
   Else
      If Month(FNaci) = 2 Then
         newday = Day(CDate(FAct)) + 28
      Else
         newday = Day(CDate(FAct)) + 30
      End If
   End If
   newmonth = newmonth - 1
Else
   newday = Day(CDate(FAct))
End If

If Month(CDate(FNaci)) = Month(Date) Then
   
   Meses = 0
Else
   Meses = newmonth - Month(CDate(FNaci))
End If

If Meses < 0 And Anios = 0 Then
   Meses = Meses + 12
End If

Dias = newday - Day(CDate(FNaci))

If FNaci <= FAct Then

'Me.TextBox3.Text = Anios & " Años, " & Meses & " Meses, " & Dias & " Dias."
   laba.Caption = Anios
   If Month(Date) = Month(FNaci) Then
      If Day(Date) > Day(FNaci) Then
         Meses = Meses
      Else
         If Day(Date) = Day(FNaci) Then
            Meses = 0
         Else
            Meses = 11
         End If
      End If
   End If
   labm.Caption = Meses
   labd.Caption = Dias
Else
   MsgBox "Fecha Inválida"
   laba.Caption = 0
   labm.Caption = 0
   labd.Caption = 0
End If

End Sub

Public Sub Modif_caracter()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
                          
'ConbdSapp.Execute "UPDATE linmmdd SET zona = REPLACE(zona, " '", '-') where fecha >='"& mfd. and nro_flia=6
'Xrecclii.Close
ConbdSapp.Close

End Sub


