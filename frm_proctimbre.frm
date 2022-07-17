VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_proctimbre 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesar timbres para emisión"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   Icon            =   "frm_proctimbre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2760
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "Por socio:"
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
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc data_cli 
      Height          =   375
      Left            =   480
      Top             =   120
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
   Begin MSAdodcLib.Adodc data_llamod 
      Height          =   495
      Left            =   3480
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "data_llamod"
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
   Begin MSAdodcLib.Adodc data_lla 
      Height          =   375
      Left            =   1800
      Top             =   2040
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
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Crystal.CrystalReport cr3 
      Left            =   4200
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crrr 
      Left            =   2640
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crr 
      Left            =   6480
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_emiserv 
      Caption         =   "data_emiserv"
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
      Top             =   2040
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.Data data_conv 
      Caption         =   "data_conv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.Data data_emitiq 
      Caption         =   "data_emitiq"
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
      Top             =   1920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Picture         =   "frm_proctimbre.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Picture         =   "frm_proctimbre.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin MSMask.MaskEdBox mh 
      Height          =   375
      Left            =   4680
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
   Begin MSMask.MaskEdBox md 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
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
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
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
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   2760
      Picture         =   "frm_proctimbre.frx":0F56
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "frm_proctimbre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MaskEdBox2_Change()

End Sub

Private Sub Command1_Click()
Dim Ximptim As Integer
data_emitiq.DatabaseName = App.path & "\env_tiq.mdb"

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)

Set MiBaseact = Unasesact.OpenDatabase(App.path & "\env_tiq.mdb")

MiBaseact.Execute "Delete * from emitiq"
MiBaseact.Execute "Delete * from emiserv"

data_emitiq.RecordSource = "emitiq"
data_emitiq.Refresh

If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      frm_proctimbre.MousePointer = 11
      Data1.Recordset.FindFirst "codest =" & 995
      If Not Data1.Recordset.NoMatch Then
         Ximptim = Data1.Recordset("cons")
      Else
         Ximptim = 53
      End If
      If Check1.Value = 1 Then
         If Text1.Text <> "" Then
            data_lla.RecordSource = "Select * from llamado where matric =" & Text1.Text & " and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
            data_lla.Refresh
         Else
            data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
            data_lla.Refresh
         End If
      Else
         data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
         data_lla.Refresh
      End If
      If data_lla.Recordset.RecordCount > 0 Then
         data_lla.Recordset.MoveFirst
         Do While Not data_lla.Recordset.EOF
            data_llamod.RecordSource = "Select * from resplla where nro =" & data_lla.Recordset("nrolla") & " and mm in (2,3)"
            data_llamod.Refresh
            If data_llamod.Recordset.RecordCount > 0 Then
            Else
                If IsNull(data_lla.Recordset("categ")) = False Then
'                   data_conv.Recordset.FindFirst "cnv_codigo ='" & data_lla.Recordset("categ") & "'"
                   data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lla.Recordset("categ") & "'"
                   data_conv.Refresh
                   If data_conv.Recordset.RecordCount > 0 Then
                      If data_conv.Recordset("cnv_colrec") = "R" Or data_conv.Recordset("cnv_colrec") = "M" Then
                         If IsNull(data_conv.Recordset("cnv_grupo")) = True Then
                               If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_codigo") <> "SEGAM" Then
                                  If IsNull(data_lla.Recordset("base")) = False Then
                                     If data_lla.Recordset("base") = 0 Then
                                        If data_conv.Recordset("cnv_codigo") = "CASA3" Or data_conv.Recordset("cnv_codigo") = "CCOMSP" Or _
                                           data_conv.Recordset("cnv_codigo") = "CCSA" Or data_conv.Recordset("cnv_codigo") = "CLUBBH" Or _
                                           data_conv.Recordset("cnv_codigo") = "IMP2" Or data_conv.Recordset("cnv_codigo") = "REFIN" Or _
                                           data_conv.Recordset("cnv_codigo") = "INAUCO" Or data_conv.Recordset("cnv_codigo") = "SEGAM" Or _
                                           data_conv.Recordset("cnv_codigo") = "EMERNT" Then
                                        Else
                                            data_emitiq.Recordset.AddNew
                                            data_emitiq.Recordset("mat") = data_lla.Recordset("matric")
                                            data_emitiq.Recordset("nombre") = Mid(data_lla.Recordset("nombre"), 1, 50)
                                            data_emitiq.Recordset("imp") = Ximptim
                                            data_emitiq.Recordset("fecha") = data_lla.Recordset("fecha")
                                            data_emitiq.Recordset("categ") = data_lla.Recordset("categ")
                                            data_emitiq.Recordset("est") = data_lla.Recordset("cancela")
                                            data_emitiq.Recordset("movil") = data_lla.Recordset("movilpas")
                                            data_emitiq.Recordset.Update
                                        End If
                                     End If
                                  Else
                                     If data_conv.Recordset("cnv_codigo") = "CASA3" Or _
                                        data_conv.Recordset("cnv_codigo") = "CCSA" Or _
                                        data_conv.Recordset("cnv_codigo") = "IMP2" Or _
                                        data_conv.Recordset("cnv_codigo") = "INAUCO" Then
                                     Else
                                        data_emitiq.Recordset.AddNew
                                        data_emitiq.Recordset("mat") = data_lla.Recordset("matric")
                                        data_emitiq.Recordset("nombre") = Mid(data_lla.Recordset("nombre"), 1, 50)
                                        data_emitiq.Recordset("imp") = Ximptim
                                        data_emitiq.Recordset("fecha") = data_lla.Recordset("fecha")
                                        data_emitiq.Recordset("categ") = data_lla.Recordset("categ")
                                        data_emitiq.Recordset("est") = data_lla.Recordset("cancela")
                                        data_emitiq.Recordset("movil") = data_lla.Recordset("movilpas")
                                        data_emitiq.Recordset.Update
                                     End If
                                  End If
                               End If
                         Else
                            If data_conv.Recordset("cnv_grupo") = "" Then
                               If data_conv.Recordset("cnv_cant_r") = 2 Then
                                  If IsNull(data_lla.Recordset("base")) = False Then
                                     If data_lla.Recordset("base") = 0 Then
                                        data_emitiq.Recordset.AddNew
                                        data_emitiq.Recordset("mat") = data_lla.Recordset("matric")
                                        data_emitiq.Recordset("nombre") = Mid(data_lla.Recordset("nombre"), 1, 50)
                                        data_emitiq.Recordset("imp") = Ximptim
                                        data_emitiq.Recordset("fecha") = data_lla.Recordset("fecha")
                                        data_emitiq.Recordset("categ") = data_lla.Recordset("categ")
                                        data_emitiq.Recordset("est") = data_lla.Recordset("cancela")
                                        data_emitiq.Recordset("movil") = data_lla.Recordset("movilpas")
                                        data_emitiq.Recordset.Update
                                     End If
                                  Else
                                     data_emitiq.Recordset.AddNew
                                     data_emitiq.Recordset("mat") = data_lla.Recordset("matric")
                                     data_emitiq.Recordset("nombre") = Mid(data_lla.Recordset("nombre"), 1, 50)
                                     data_emitiq.Recordset("imp") = Ximptim
                                     data_emitiq.Recordset("fecha") = data_lla.Recordset("fecha")
                                     data_emitiq.Recordset("categ") = data_lla.Recordset("categ")
                                     data_emitiq.Recordset("est") = data_lla.Recordset("cancela")
                                     data_emitiq.Recordset("movil") = data_lla.Recordset("movilpas")
                                     data_emitiq.Recordset.Update
                                  End If
                               End If
                         
                            End If
                         End If
                      Else
                         If data_conv.Recordset("cnv_colrec") = "A" Then
                            If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
                                If data_conv.Recordset("cnv_codigo") = "SUAIN" Then
                                   If IsNull(data_lla.Recordset("matric")) = False Then
                                         If data_lla.Recordset("matric") <> 0 Then
                                            data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lla.Recordset("matric")
                                            data_cli.Refresh
                                            If data_cli.Recordset.RecordCount > 0 Then
                                               If data_cli.Recordset("cl_nrocobr") = 650 Or _
                                                  data_cli.Recordset("cl_nrocobr") = 33 Or _
                                                  data_cli.Recordset("cl_nrocobr") = 22 Or _
                                                  data_cli.Recordset("cl_nrocobr") = 2 Or _
                                                  data_cli.Recordset("cl_nrocobr") = 1 Or _
                                                  data_cli.Recordset("cl_nrocobr") = 10 Or _
                                                  data_cli.Recordset("cl_nrocobr") = 3 Or _
                                                  data_cli.Recordset("cl_nrocobr") = 698 Then
                                                  If data_conv.Recordset("cnv_codigo") = "TALA50" Then
    '                                                 data_emitiq.Recordset.AddNew
    '                                                 data_emitiq.Recordset("mat") = data_lla.Recordset("matric")
    '                                                 data_emitiq.Recordset("nombre") = Mid(data_lla.Recordset("nombre"), 1, 50)
    '                                                 data_emitiq.Recordset("imp") = Ximptim
    '                                                 data_emitiq.Recordset("fecha") = data_lla.Recordset("fecha")
    '                                                 data_emitiq.Recordset("categ") = data_lla.Recordset("categ")
    '                                                 data_emitiq.Recordset.Update
                                                  End If
                                               Else
                                                  If IsNull(data_lla.Recordset("base")) = False Then
                                                     If data_lla.Recordset("base") = 0 Then
                                                        data_emitiq.Recordset.AddNew
                                                        data_emitiq.Recordset("mat") = data_lla.Recordset("matric")
                                                        data_emitiq.Recordset("nombre") = Mid(data_lla.Recordset("nombre"), 1, 50)
                                                        data_emitiq.Recordset("imp") = Ximptim
                                                        data_emitiq.Recordset("fecha") = data_lla.Recordset("fecha")
                                                        data_emitiq.Recordset("categ") = data_lla.Recordset("categ")
                                                        data_emitiq.Recordset("est") = data_lla.Recordset("cancela")
                                                        data_emitiq.Recordset("movil") = data_lla.Recordset("movilpas")
                                                        data_emitiq.Recordset.Update
                                                     End If
                                                  Else
                                                     data_emitiq.Recordset.AddNew
                                                     data_emitiq.Recordset("mat") = data_lla.Recordset("matric")
                                                     data_emitiq.Recordset("nombre") = Mid(data_lla.Recordset("nombre"), 1, 50)
                                                     data_emitiq.Recordset("imp") = Ximptim
                                                     data_emitiq.Recordset("fecha") = data_lla.Recordset("fecha")
                                                     data_emitiq.Recordset("categ") = data_lla.Recordset("categ")
                                                     data_emitiq.Recordset("est") = data_lla.Recordset("cancela")
                                                     data_emitiq.Recordset("movil") = data_lla.Recordset("movilpas")
                                                     data_emitiq.Recordset.Update
                                                  End If
                                               End If
                                            End If
                                         End If
                                   End If
                                Else
                                   If data_conv.Recordset("cnv_codigo") = "CALF34" Or data_conv.Recordset("cnv_codigo") = "CALF35" Or _
                                      data_conv.Recordset("cnv_codigo") = "CALF27" Or data_conv.Recordset("CNV_CODIGO") = "CALF25" Or _
                                      data_conv.Recordset("CNV_CODIGO") = "CALF26" Then
                                      If IsNull(data_lla.Recordset("base")) = False Then
                                         If data_lla.Recordset("base") = 0 Then
                                            data_emitiq.Recordset.AddNew
                                            data_emitiq.Recordset("mat") = data_lla.Recordset("matric")
                                            data_emitiq.Recordset("nombre") = Mid(data_lla.Recordset("nombre"), 1, 50)
                                            data_emitiq.Recordset("imp") = Ximptim
                                            data_emitiq.Recordset("fecha") = data_lla.Recordset("fecha")
                                            data_emitiq.Recordset("categ") = data_lla.Recordset("categ")
                                            data_emitiq.Recordset("est") = data_lla.Recordset("cancela")
                                            data_emitiq.Recordset("movil") = data_lla.Recordset("movilpas")
                                            data_emitiq.Recordset.Update
                                         End If
                                      Else
                                         data_emitiq.Recordset.AddNew
                                         data_emitiq.Recordset("mat") = data_lla.Recordset("matric")
                                         data_emitiq.Recordset("nombre") = Mid(data_lla.Recordset("nombre"), 1, 50)
                                         data_emitiq.Recordset("imp") = Ximptim
                                         data_emitiq.Recordset("fecha") = data_lla.Recordset("fecha")
                                         data_emitiq.Recordset("categ") = data_lla.Recordset("categ")
                                         data_emitiq.Recordset("est") = data_lla.Recordset("cancela")
                                         data_emitiq.Recordset("movil") = data_lla.Recordset("movilpas")
                                         data_emitiq.Recordset.Update
                                      End If
                                   Else
                                      If data_conv.Recordset("cnv_codigo") = "CALF34" Or data_conv.Recordset("cnv_codigo") = "CALF35" Or _
                                           data_conv.Recordset("cnv_codigo") = "CALF27" Or data_conv.Recordset("CNV_CODIGO") = "CALF26" Or _
                                           data_conv.Recordset("CNV_CODIGO") = "CALF25" Then
                                         If IsNull(data_lla.Recordset("base")) = False Then
                                            If data_lla.Recordset("base") = 0 Then
                                                data_emitiq.Recordset.AddNew
                                                data_emitiq.Recordset("mat") = data_lla.Recordset("matric")
                                                data_emitiq.Recordset("nombre") = Mid(data_lla.Recordset("nombre"), 1, 50)
                                                data_emitiq.Recordset("imp") = Ximptim
                                                data_emitiq.Recordset("fecha") = data_lla.Recordset("fecha")
                                                data_emitiq.Recordset("categ") = data_lla.Recordset("categ")
                                                data_emitiq.Recordset("est") = data_lla.Recordset("cancela")
                                                data_emitiq.Recordset("movil") = data_lla.Recordset("movilpas")
                                                data_emitiq.Recordset.Update
                                            End If
                                         Else
                                            data_emitiq.Recordset.AddNew
                                            data_emitiq.Recordset("mat") = data_lla.Recordset("matric")
                                            data_emitiq.Recordset("nombre") = Mid(data_lla.Recordset("nombre"), 1, 50)
                                            data_emitiq.Recordset("imp") = Ximptim
                                            data_emitiq.Recordset("fecha") = data_lla.Recordset("fecha")
                                            data_emitiq.Recordset("categ") = data_lla.Recordset("categ")
                                            data_emitiq.Recordset("est") = data_lla.Recordset("cancela")
                                            data_emitiq.Recordset("movil") = data_lla.Recordset("movilpas")
                                            data_emitiq.Recordset.Update
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
            data_lla.Recordset.MoveNext
         Loop
         If data_emitiq.Recordset.RecordCount > 0 Then
            data_emitiq.Recordset.MoveFirst
            Do While Not data_emitiq.Recordset.EOF
               If IsNull(data_emitiq.Recordset("mat")) = False Then
                  If data_emitiq.Recordset("mat") = 0 Then
                     data_emitiq.Recordset.Edit
                     data_emitiq.Recordset("est") = 78
                     data_emitiq.Recordset.Update
                  Else
                     If data_emitiq.Recordset("mat") >= 99999999 Then
                        data_emitiq.Recordset.Edit
                        data_emitiq.Recordset("est") = 78
                        data_emitiq.Recordset.Update
                     Else
                        data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_emitiq.Recordset("mat")
                        data_cli.Refresh
                        If data_cli.Recordset.RecordCount > 0 Then
                           If data_cli.Recordset("cl_grupo") = 501 Or data_cli.Recordset("cl_codconv") = "SOC" Or data_cli.Recordset("cl_nrocobr") = 0 Or _
                              data_cli.Recordset("cl_nrocobr") = 700 Or data_cli.Recordset("cl_nrocobr") = 206 Or data_cli.Recordset("cl_nrocobr") = 696 Then
                              data_emitiq.Recordset.Edit
                              data_emitiq.Recordset("est") = 78
                              data_emitiq.Recordset.Update
                           Else
                              If IsNull(data_emitiq.Recordset("est")) = False Then
                                 If data_emitiq.Recordset("est") = 1 Or data_emitiq.Recordset("categ") = "SEGAM" Then
                                    data_emitiq.Recordset.Edit
                                    data_emitiq.Recordset("est") = 78
                                    data_emitiq.Recordset.Update
                                 Else
                                    data_emitiq.Recordset.Edit
                                    data_emitiq.Recordset("cob") = data_cli.Recordset("cl_nrocobr")
                                    If IsNull(data_cli.Recordset("estado")) = True Then
                                       data_emitiq.Recordset("est") = 9
                                    Else
                                       If data_cli.Recordset("estado") = 1 Then
                                          data_emitiq.Recordset("est") = 9
                                       Else
                                          data_emitiq.Recordset("est") = 8
                                       End If
                                    End If
                                    data_emitiq.Recordset.Update
                                 End If
                              Else
                                 data_emitiq.Recordset.Edit
                                 data_emitiq.Recordset("cob") = data_cli.Recordset("cl_nrocobr")
                                 If IsNull(data_cli.Recordset("estado")) = True Then
                                    data_emitiq.Recordset("est") = 9
                                 Else
                                    If data_cli.Recordset("estado") = 1 Then
                                       data_emitiq.Recordset("est") = 9
                                    Else
                                       data_emitiq.Recordset("est") = 8
                                    End If
                                 End If
                                 data_emitiq.Recordset.Update
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
               data_emitiq.Recordset.MoveNext
            Loop
            MiBaseact.Execute "Delete from emitiq where est =" & 78
            MiBaseact.Execute "delete from emitiq where est =" & 8
            MiBaseact.Execute "delete from emitiq where categ in ('CCOMSP','CLUBBH','REFIN','SEGAM','SMI16','UNIVSC')"
            MiBaseact.Execute "delete from emitiq where cob =" & 0
            MiBaseact.Execute "delete from emitiq where cob is null"
         
         End If
         If data_emiserv.Recordset.RecordCount > 0 Then
            data_emiserv.Recordset.MoveFirst
            Do While Not data_emiserv.Recordset.EOF
               data_emiserv.Recordset.Delete
               data_emiserv.Recordset.MoveNext
            Loop
         End If
'''         Command3_Click
         
         data_emiserv.RecordSource = "Select * from emiserv"
         data_emiserv.Refresh
         MsgBox "Proceso terminado", vbInformation, "Mensaje"
'         crr.ReportTitle = "FECHA DESDE: " & md.Text & " HASTA: " & mh.Text
         data_emitiq.RecordSource = "Select * from emitiq"
         data_emitiq.Refresh
         
         cr3.ReportFileName = App.path & "\inftimbres.rpt"
         cr3.ReportTitle = "FECHA DESDE: " & md.Text & " HASTA: " & mh.Text
         cr3.Action = 1
         
'         crrr.ReportFileName = App.path & "\infservicios.rpt"
'         crrr.Action = 1
         
      Else
         MsgBox "No existen registros para procesar", vbExclamation, "Mensaje"
      End If
   End If
End If
frm_proctimbre.MousePointer = 0

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
data_emitiq.DatabaseName = App.path & "\env_cp.mdb"

If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      frm_proctimbre.MousePointer = 11
      Data1.Recordset.FindFirst "codest =" & 995
      If Not Data1.Recordset.NoMatch Then
         Ximptim = Data1.Recordset("cons")
      Else
         Ximptim = 53
      End If
      data_emitiq.RecordSource = "emitiq"
      data_emitiq.Refresh
      If data_emitiq.Recordset.RecordCount > 0 Then
         data_emitiq.Recordset.MoveFirst
         Do While Not data_emitiq.Recordset.EOF
            data_emitiq.Recordset.Delete
            data_emitiq.Recordset.MoveNext
         Loop
      End If
      data_lla.RecordSource = "Select * from llamado where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
      data_lla.Refresh
      If data_lla.Recordset.RecordCount > 0 Then
         data_lla.Recordset.MoveFirst
         Do While Not data_lla.Recordset.EOF
            If IsNull(data_lla.Recordset("categ")) = False Then
               If data_lla.Recordset("categ") = "911" Or data_lla.Recordset("categ") = "911B" Or _
                  data_lla.Recordset("categ") = "CASH" Or data_lla.Recordset("categ") = "SEMM" Or _
                  data_lla.Recordset("categ") = "SEMM1" Or data_lla.Recordset("categ") = "CCASMU" Or _
                  data_lla.Recordset("categ") = "1727" Or data_lla.Recordset("categ") = "CPS" Or _
                  data_lla.Recordset("categ") = "CPSSA" Then
                  If data_lla.Recordset("movilpas") = 4 Or _
                     data_lla.Recordset("movilpas") = 103 Or _
                     data_lla.Recordset("movilpas") = 202 Or _
                     data_lla.Recordset("movilpas") = 161 Or _
                     data_lla.Recordset("movilpas") = 301 Or _
                     data_lla.Recordset("movilpas") = 306 Or _
                     data_lla.Recordset("movilpas") = 203 Or _
                     data_lla.Recordset("movilpas") = 1 Or _
                     data_lla.Recordset("movilpas") = 207 Then
                     data_emitiq.Recordset.AddNew
                     data_emitiq.Recordset("mat") = data_lla.Recordset("matric")
                     data_emitiq.Recordset("nombre") = Mid(data_lla.Recordset("nombre"), 1, 50)
                     data_emitiq.Recordset("imp") = Ximptim
                     data_emitiq.Recordset("fecha") = data_lla.Recordset("fecha")
                     data_emitiq.Recordset("categ") = data_lla.Recordset("categ")
                     data_emitiq.Recordset("est") = data_lla.Recordset("cancela")
                     data_emitiq.Recordset("movil") = data_lla.Recordset("movilpas")
                     data_emitiq.Recordset.Update
                  End If
               Else
                  data_conv.Recordset.FindFirst "cnv_codigo ='" & data_lla.Recordset("categ") & "'"
                  If Not data_conv.Recordset.NoMatch Then
                     If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
                        If data_conv.Recordset("cnv_grupo") <> "" Then
                        Else
                           If data_lla.Recordset("categ") = "MSP" Or data_lla.Recordset("categ") = "50" Or _
                              data_lla.Recordset("categ") = "55" Or data_lla.Recordset("categ") = "ASOCES" Or _
                              data_lla.Recordset("categ") = "IMPASA" Or data_lla.Recordset("categ") = "UNIVS" Or _
                              data_lla.Recordset("categ") = "GANOS" Or data_lla.Recordset("categ") = "CASANO" Or _
                              data_lla.Recordset("categ") = "CASANR" Or data_lla.Recordset("categ") = "CCSA" Or _
                              data_lla.Recordset("categ") = "CCNOS" Or data_lla.Recordset("categ") = "CCNRE" Or _
                              data_lla.Recordset("categ") = "HEVAN" Or data_lla.Recordset("categ") = "HEVANO" Or _
                              data_lla.Recordset("categ") = "HEVANR" Or data_lla.Recordset("categ") = "SMIN" Or _
                              data_lla.Recordset("categ") = "SMINR" Or data_lla.Recordset("categ") = "IMPNO" Then
                           Else
                            If data_lla.Recordset("movilpas") = 4 Or _
                               data_lla.Recordset("movilpas") = 103 Or _
                               data_lla.Recordset("movilpas") = 202 Or _
                               data_lla.Recordset("movilpas") = 161 Or _
                               data_lla.Recordset("movilpas") = 301 Or _
                               data_lla.Recordset("movilpas") = 306 Or _
                               data_lla.Recordset("movilpas") = 203 Or _
                               data_lla.Recordset("movilpas") = 1 Or _
                               data_lla.Recordset("movilpas") = 207 Then
'                              If data_lla.Recordset("movilpas") = 4 Or _
'                                 data_lla.Recordset("movilpas") = 15 Or _
'                                 data_lla.Recordset("movilpas") = 20 Or _
'                                 data_lla.Recordset("movilpas") = 16 Or _
'                                 data_lla.Recordset("movilpas") = 24 Or _
'                                 data_lla.Recordset("movilpas") = 25 Or _
'                                 data_lla.Recordset("movilpas") = 161 Or _
'                                 data_lla.Recordset("movilpas") = 19 Then
                                 data_emitiq.Recordset.AddNew
                                 data_emitiq.Recordset("mat") = data_lla.Recordset("matric")
                                 data_emitiq.Recordset("nombre") = Mid(data_lla.Recordset("nombre"), 1, 50)
                                 data_emitiq.Recordset("imp") = Ximptim
                                 data_emitiq.Recordset("fecha") = data_lla.Recordset("fecha")
                                 data_emitiq.Recordset("categ") = data_lla.Recordset("categ")
                                 data_emitiq.Recordset("est") = data_lla.Recordset("cancela")
                                 data_emitiq.Recordset("movil") = data_lla.Recordset("movilpas")
                                 data_emitiq.Recordset.Update
                              End If
                           End If
                        End If
                     Else
                        If data_lla.Recordset("categ") = "MSP" Or data_lla.Recordset("categ") = "50" Or _
                           data_lla.Recordset("categ") = "55" Or data_lla.Recordset("categ") = "ASOCES" Or _
                           data_lla.Recordset("categ") = "IMPASA" Then
                        Else
                           If data_lla.Recordset("movilpas") = 4 Or _
                              data_lla.Recordset("movilpas") = 103 Or _
                              data_lla.Recordset("movilpas") = 202 Or _
                              data_lla.Recordset("movilpas") = 161 Or _
                              data_lla.Recordset("movilpas") = 301 Or _
                              data_lla.Recordset("movilpas") = 306 Or _
                              data_lla.Recordset("movilpas") = 203 Or _
                              data_lla.Recordset("movilpas") = 1 Or _
                              data_lla.Recordset("movilpas") = 207 Then
                              data_emitiq.Recordset.AddNew
                              data_emitiq.Recordset("mat") = data_lla.Recordset("matric")
                              data_emitiq.Recordset("nombre") = Mid(data_lla.Recordset("nombre"), 1, 50)
                              data_emitiq.Recordset("imp") = Ximptim
                              data_emitiq.Recordset("fecha") = data_lla.Recordset("fecha")
                              data_emitiq.Recordset("categ") = data_lla.Recordset("categ")
                              data_emitiq.Recordset("est") = data_lla.Recordset("cancela")
                              data_emitiq.Recordset("movil") = data_lla.Recordset("movilpas")
                              data_emitiq.Recordset.Update
                           End If
                        End If
                     End If
                  End If
               End If
            End If
            data_lla.Recordset.MoveNext
         Loop
      End If
   End If
End If

'         data_emitiq.RecordSource = "Select * from emitiq"
'         data_emitiq.Refresh
'         data_emiserv.RecordSource = "Select * from emiserv"
'         data_emiserv.Refresh
'         MsgBox "Proceso terminado", vbInformation, "Mensaje"
'         crr.ReportFileName = App.Path & "\inftimbres.rpt"
'         crr.ReportTitle = "FECHA DESDE: " & md.Text & " HASTA: " & mh.Text
'         crr.Action = 1
'         cr3.ReportFileName = App.Path & "\inftimbresb.rpt"
'         cr3.Action = 1
         
'         crrr.ReportFileName = App.Path & "\infservicios.rpt"
'         crrr.Action = 1
         
'      Else
'         MsgBox "No existen registros para procesar", vbExclamation, "Mensaje"
'      End If
'   End If
'End If
'frm_proctimbre.MousePointer = 0

End Sub

Private Sub Form_Load()
data_lla.ConnectionString = "dsn=" & Xconexrmt

'data_lla.RecordSource = "llamado"
'data_lla.Refresh
data_conv.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_conv.RecordSource = "convenio"
data_conv.Refresh
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "estudios"
Data1.Refresh
data_cli.ConnectionString = "dsn=" & Xconexrmt
'data_cli.RecordSource = "clientes"
'data_cli.Refresh

data_llamod.ConnectionString = "dsn=" & Xconexrmt

data_emitiq.DatabaseName = App.path & "\env_tiq.mdb"
data_emiserv.DatabaseName = App.path & "\env_tiq.mdb"
data_emiserv.RecordSource = "emiserv"
data_emiserv.Refresh


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
   Command1.SetFocus
End If

End Sub
