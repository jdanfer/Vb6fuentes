VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infabmvar 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de Altas/Bajas de socios"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infabmvar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frm_infabmvar.frx":0442
   ScaleHeight     =   7620
   ScaleWidth      =   8310
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_abm 
      Height          =   375
      Left            =   1800
      Top             =   6840
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
      Caption         =   "data_abm"
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
   Begin VB.Data data_clideu 
      Caption         =   "data_clideu"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSComctlLib.ProgressBar barr 
      Height          =   375
      Left            =   600
      TabIndex        =   24
      Top             =   6360
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Data data_convloc 
      Caption         =   "data_convloc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   4920
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   6000
      TabIndex        =   22
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   6000
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   6000
      TabIndex        =   19
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Data data_zon 
      Caption         =   "data_zon"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tipo de informe"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   240
      TabIndex        =   15
      Top             =   5160
      Width           =   7815
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FF8080&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF8080&
         Caption         =   "Resumen"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin Crystal.CrystalReport reporte 
      Left            =   7200
      Top             =   6360
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      MouseIcon       =   "frm_infabmvar.frx":074C
      MousePointer    =   99  'Custom
      Picture         =   "frm_infabmvar.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   6840
      Width           =   735
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
      Height          =   615
      Left            =   600
      MouseIcon       =   "frm_infabmvar.frx":0FE0
      MousePointer    =   99  'Custom
      Picture         =   "frm_infabmvar.frx":12EA
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Procesar"
      Top             =   6840
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos de informe"
      Height          =   4935
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7815
      Begin MSAdodcLib.Adodc data_conv 
         Height          =   375
         Left            =   4920
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
         Caption         =   "data_conv"
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
         Left            =   5040
         Top             =   3720
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
      Begin VB.CheckBox Check3 
         BackColor       =   &H00800000&
         Caption         =   "ORDENAR POR MOTIVOS DE BAJA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   4560
         Width           =   4455
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00800000&
         Caption         =   "INCLUIR DEUDA DEL SOCIO"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   4080
         Width           =   4455
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00800000&
         Caption         =   "INCLUIR DATOS PATRONIMICOS"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3600
         Width           =   4455
      End
      Begin VB.ComboBox Combo3 
         Height          =   360
         ItemData        =   "frm_infabmvar.frx":1874
         Left            =   2400
         List            =   "frm_infabmvar.frx":188D
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3000
         Width           =   4455
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "frm_infabmvar.frx":18C5
         Left            =   2400
         List            =   "frm_infabmvar.frx":18DE
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2400
         Width           =   4455
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_infabmvar.frx":1916
         Left            =   2400
         List            =   "frm_infabmvar.frx":192F
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1800
         Width           =   4455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "Bajas"
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Activos"
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   1200
         Value           =   -1  'True
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   4800
         TabIndex        =   3
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         Left            =   2400
         TabIndex        =   0
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         BackColor       =   &H00C0FFFF&
         Caption         =   "Con fecha en blanco se toman todos los datos del padrón.-"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   6615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         X1              =   0
         X2              =   7800
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "Selección por ---->>>"
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Tipo de socios:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Rango de Fechas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   120
      Picture         =   "frm_infabmvar.frx":1970
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   1095
   End
End
Attribute VB_Name = "frm_infabmvar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check2.Value = 1 Then
   Check2.Value = 0
End If
If Check3.Value = 1 Then
   Check3.Value = 0
End If
   
End Sub

Private Sub Check2_Click()
If Check1.Value = 1 Then
   Check1.Value = 0
End If
If Check3.Value = 1 Then
   Check3.Value = 0
End If

End Sub

Private Sub Check3_Click()
If Option2.Value = True Then

Else
   Check3.Value = 0
End If
If Check1.Value = 1 Then
   Check1.Value = 0
End If
If Check2.Value = 1 Then
   Check2.Value = 0
End If

End Sub

Private Sub Combo1_Click()
On Error GoTo Elconve
If Combo1.Text = "CONVENIO" Then
   frm_opsconv.Show vbModal
Else
   If Combo1.Text = "EDAD" Then
      frm_opsedad.Show vbModal
   Else
      If Combo1.Text = "COBRADOR" Then
         frm_opscob.Show vbModal
      Else
         If Combo1.Text = "PROMOTOR" Then
            frm_opspro.Show vbModal
         Else
            If Combo1.Text = "RADIO" Then
               frm_opszona.Show vbModal
            Else
               If Combo1.Text = "MUTUALISTA" Then
                  Combo2.Clear
                  Carga_mutuales
               End If
            End If
         End If
      End If
   End If
End If

Exit Sub

Elconve:
        If Err.Number = 364 Then
''           MsgBox "Vea"
        End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo2.SetFocus
End If

End Sub

Private Sub Combo2_Click()
On Error GoTo Elconve2
If Combo2.Text = "CONVENIO" Then
   frm_opsconv.Show vbModal
Else
   If Combo2.Text = "EDAD" Then
      frm_opsedad.Show vbModal
   Else
      If Combo2.Text = "COBRADOR" Then
         frm_opscob.Show vbModal
      Else
         If Combo2.Text = "PROMOTOR" Then
            frm_opspro.Show vbModal
         Else
            If Combo2.Text = "RADIO" Then
               frm_opszona.Show vbModal
            End If
         End If
      End If
   End If
End If

Exit Sub

Elconve2:
        If Err.Number = 364 Then
''           MsgBox "Vea"
        End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo3.SetFocus
End If

End Sub

Private Sub Combo3_Click()
On Error GoTo Elconve3
If Combo3.Text = "CONVENIO" Then
   frm_opsconv.Show vbModal
Else
   If Combo3.Text = "EDAD" Then
      frm_opsedad.Show vbModal
   Else
      If Combo3.Text = "COBRADOR" Then
         frm_opscob.Show vbModal
      Else
         If Combo3.Text = "PROMOTOR" Then
            frm_opspro.Show vbModal
         Else
            If Combo3.Text = "RADIO" Then
               frm_opszona.Show vbModal
            End If
         End If
      End If
   End If
End If

Exit Sub

Elconve3:
        If Err.Number = 364 Then
''           MsgBox "Vea"
        End If

End Sub

Private Sub Command1_Click()
Dim XQueopeli As String
Dim XQueopnro As Long
Dim Xcadsel, Xcadorden As String
Dim Xnrosex As Long
Dim Xnrosexs As String
Dim Xfecedd, Xfecedh As Date
Dim Xdiad, Xdiah, Xdiabis As Long
Xop1 = 0
Xop2 = 0
Xop3 = 0
Xop4 = 0
Xop5 = 0
Xop6 = 0
Xcadsel = ""
Xcadorden = ""
barr.Visible = True
barr.Value = 0
frm_infabmvar.MousePointer = 11
Command1.Enabled = False
Command2.Enabled = False

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
MiBaseact.Execute "Delete * from infcli"
data_inf.RecordSource = "infcli"
data_inf.Refresh

DoEvents
If Combo1.Text = "" And Combo2.Text = "" And Combo3.Text = "" Then
   XQueopnro = 0
Else
   If Combo2.Text = "" And Combo3.Text = "" And Combo1.Text <> "" Then
      XQueopnro = 1
      XQueopeli = Combo1.Text
      Command3_Click
   Else
      If Combo1.Text = Combo2.Text Or Combo1.Text = Combo3.Text Then
         MsgBox "Verifique opciones de selección", vbInformation
      Else
         If Combo2.Text = Combo3.Text Then
            MsgBox "Verifique opciones de selección", vbInformation
         Else
            If Combo1.Text = "CONVENIO" Or Combo2.Text = "CONVENIO" Or Combo3.Text = "CONVENIO" Or Combo1.Text = "MUTUALISTA" Then
               Xop1 = 1
            Else
               Xop1 = 0
            End If
            If Combo1.Text = "COBRADOR" Or Combo2.Text = "COBRADOR" Or Combo3.Text = "COBRADOR" Then
               Xop2 = 1
            Else
               Xop2 = 0
            End If
            If Combo1.Text = "PROMOTOR" Or Combo2.Text = "PROMOTOR" Or Combo3.Text = "PROMOTOR" Then
               Xop3 = 1
            Else
               Xop3 = 0
            End If
            If Combo1.Text = "RADIO" Or Combo2.Text = "RADIO" Or Combo3.Text = "RADIO" Then
               Xop4 = 1
            Else
               Xop4 = 0
            End If
            If Combo1.Text = "SEXO" Or Combo2.Text = "SEXO" Or Combo3.Text = "SEXO" Then
               Xop5 = 1
               Xnrosexs = InputBox("Ingrese SEXO (0=TODOS, 1=MASC,2=FEM):", "Selección de Sexo")
               If Xnrosexs = "" Then
                  Xnrosex = 0
               Else
                  Xnrosex = Val(Xnrosexs)
               End If
            Else
               Xop5 = 0
            End If
            If Combo1.Text = "EDAD" Or Combo2.Text = "EDAD" Or Combo3.Text = "EDAD" Then
               Xop6 = 1
            Else
               Xop6 = 0
            End If
            
            If Xop2 = 1 Then
               If Wopscob <> 0 Then
                  Xcadsel = "cl_nrocobr =" & Wopscob
               End If
            End If
            
            If Xop3 = 1 Then
               If Wopspro <> 0 Then
                  If Len(Xcadsel) > 0 Then
                     Xcadsel = Xcadsel + " And cl_nrovend =" & Wopspro
                  Else
                     Xcadsel = "cl_nrovend =" & Wopspro
                  End If
               End If
            End If
            If Xop4 = 1 Then
                If Len(Xcadsel) > 0 Then
                   If Wopszon = 1 Then
                      Xcadsel = Xcadsel + " And cl_grupo >=" & 101 & " And cl_grupo <=" & 104
                   End If
                   If Wopszon = 2 Then
                      Xcadsel = Xcadsel + " And cl_grupo >=" & 201 & " And cl_grupo <=" & 209
                   End If
                   If Wopszon = 3 Then
                      Xcadsel = Xcadsel + " And cl_grupo >=" & 301 & " And cl_grupo <=" & 312
                   End If
                   If Wopszon = 4 Then
                      Xcadsel = Xcadsel + " And cl_grupo >=" & 401 & " And cl_grupo <=" & 419
                   End If
                   If Wopszon = 5 Then
                      Xcadsel = Xcadsel + " And cl_grupo >=" & 500 & " And cl_grupo <=" & 501
                   End If
                   If Wopszon = 6 Then
                      Xcadsel = Xcadsel + " And cl_grupo in (600,601,602,603,604,605,606,607,608,609,610,624)"
                   End If
                   If Wopszon = 7 Then
                      Xcadsel = Xcadsel + " And cl_grupo >=" & 700 & " And cl_grupo <=" & 722
                   End If
                   If Wopszon = 8 Then
                      Xcadsel = Xcadsel + " And cl_grupo =" & 800 & " or cl_grupo =" & 650
                   End If
                   If Wopszon = 9 Then
                      Xcadsel = Xcadsel + " And cl_grupo >=" & 801 & " And cl_grupo >=" & 803
                   End If
                   If Wopszon = 10 Then
                      Xcadsel = Xcadsel + " And cl_grupo =" & 810
                   End If
                   If Wopszon = 11 Then
                      Xcadsel = Xcadsel + " And cl_grupo >=" & 630 & " And cl_grupo >=" & 640
                   End If
                   If Wopszon = 12 Then
                      Xcadsel = Xcadsel + " And cl_grupo =" & 671
                   End If
                   If Wopszon = 13 Then
                      Xcadsel = Xcadsel + " And cl_grupo =" & 670 & " or cl_grupo =" & 672 & " or cl_grupo =" & 673 & " or cl_grupo =" & 674
                   End If
                   If Wopszon = 15 Then
                      Xcadsel = Xcadsel + " And cl_grupo =" & 815 & " or cl_grupo =" & 816
                   End If
                Else
                   If Wopszon = 1 Then
                      Xcadsel = "cl_grupo >=" & 101 & " And cl_grupo <=" & 104
                   End If
                   If Wopszon = 2 Then
                      Xcadsel = "cl_grupo >=" & 201 & " And cl_grupo <=" & 209
                   End If
                   If Wopszon = 3 Then
                      Xcadsel = "cl_grupo >=" & 301 & " And cl_grupo <=" & 312
                   End If
                   If Wopszon = 4 Then
                      Xcadsel = "cl_grupo >=" & 401 & " And cl_grupo <=" & 419
                   End If
                   If Wopszon = 5 Then
                      Xcadsel = "cl_grupo >=" & 500 & " And cl_grupo <=" & 501
                   End If
                   If Wopszon = 6 Then
                      Xcadsel = "cl_grupo >=" & 600 & " And cl_grupo <=" & 624
                   End If
                   If Wopszon = 7 Then
                      Xcadsel = "cl_grupo >=" & 700 & " And cl_grupo <=" & 722
                   End If
                   If Wopszon = 8 Then
                      Xcadsel = "cl_grupo =" & 800 & " or cl_grupo =" & 650
                   End If
                   If Wopszon = 9 Then
                      Xcadsel = "cl_grupo >=" & 801 & " And cl_grupo <=" & 803
                   End If
                   If Wopszon = 10 Then
                      Xcadsel = "cl_grupo =" & 810
                   End If
                   If Wopszon = 11 Then
                      Xcadsel = "cl_grupo >=" & 630 & " And cl_grupo <=" & 640
                   End If
                   If Wopszon = 12 Then
                      Xcadsel = "cl_grupo =" & 671
                   End If
                   If Wopszon = 13 Then
                      Xcadsel = "cl_grupo =" & 670 & " or cl_grupo =" & 672 & " or cl_grupo =" & 673 & " or cl_grupo =" & 674
                   End If
                   If Wopszon = 15 Then
                      Xcadsel = "cl_grupo >=" & 815 & " or cl_grupo <=" & 816
                   End If
                End If
            End If
            
            If Xop5 = 1 Then 'Sexo
               If Xnrosex <> 0 Then
                  If Len(Xcadsel) > 0 Then
                     Xcadsel = Xcadsel + " And cl_sexo =" & Xnrosex
                  Else
                     Xcadsel = "cl_sexo =" & Xnrosex
                  End If
               End If
            End If
            
            If Xop6 = 1 Then 'Edades
               If Wopsed = 1 Or Wopsed = 3 Then
                  If Wopsedd = 0 And Wopsedh = 110 Then
                  Else
                     Xdiad = Wopsedd * 365
                     Xdiad = Xdiad + 365
                     Xdiah = Wopsedh * 365
                     Xdiah = Xdiah + 365
                     If Wopsedd > 4 Then
                        Xdiabis = Wopsedd / 4
                        Xdiabis = Int(Xdiabis)
                        Xdiad = Xdiad + Xdiabis
                     End If
                     Xdiabis = 0
                     If Wopsedh > 4 Then
                        Xdiabis = Wopsedh / 4
                        Xdiabis = Int(Xdiabis)
                        Xdiah = Xdiah + Xdiabis
                     End If
                     Xfecedd = Date - Xdiad
                     Xfecedh = Date - Xdiah
                     If Wopsedd = 0 Then
                        Xfecedd = Date
                     End If
                     If Len(Xcadsel) > 0 Then
                        Xcadsel = Xcadsel + " And cl_fnac >='" & Format(Xfecedh, "yyyy-mm-dd") & "' And cl_fnac <='" & Format(Xfecedd, "yyyy-mm-dd") & "'"
                     Else
                        Xcadsel = "cl_fnac >='" & Format(Xfecedh, "yyyy-mm-dd") & "' And cl_fnac <='" & Format(Xfecedd, "yyyy-mm-dd") & "'"
                     End If
                  End If
               End If
               If Wopsed = 2 Then
                  If Len(Xcadsel) > 0 Then
                     Xcadsel = Xcadsel + " And month(cl_fnac) =" & Wopsedd
                  Else
                     Xcadsel = "month(cl_fnac) =" & Wopsedd
                  End If
               End If
            End If
            If Option1.Value = True Then
               If md.Text <> "__/__/____" And mh.Text <> "__/__/____" Then
'                  data_abm.RecordSource = "Select * from abmsocio where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and desc ='BAJA'"
'                  data_abm.Refresh
                  If Combo1.Text = "MUTUALISTA" Then
                     data_cli.RecordSource = "Select * from clientes where estado in (1) And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_socmnom ='" & Combo2.Text & "'"
                     data_cli.Refresh
                  Else
                     If Len(Xcadsel) > 0 Then
                        If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
                           data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' and cl_codconv <='" & Xlehas & "'"
                        Else
                           data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
                        End If
                        data_cli.Refresh
                     Else
                        If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
                           data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Xlehas & "'"
                        Else
                           data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
                        End If
                        data_cli.Refresh
                     End If
                  End If
               Else
                  md.Text = CDate("01/01/1980")
                  mh.Text = Date
                  If Len(Xcadsel) > 0 Then
                     If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
                        data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Xlehas & "'"
                     Else
                        data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
                     End If
                     data_cli.Refresh
                  Else
                     If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
                        data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Xlehas & "'"
                     Else
                        data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
                     End If
                     data_cli.Refresh
                  End If
                  md.Text = Date
                  mh.Text = Date
                  data_abm.RecordSource = "Select * from abmsocio where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and desc ='BAJA'"
                  data_abm.Refresh
                  md.Text = "__/__/____"
                  mh.Text = "__/__/____"
               End If
            Else
               If Option2.Value = True Then
                  If md.Text <> "__/__/____" And mh.Text <> "__/__/____" Then
                     data_abm.RecordSource = "Select * from abmsocio where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and desc ='BAJA'"
                     data_abm.Refresh
                     If Len(Xcadsel) > 0 Then
                        If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
                           data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Xlehas & "'"
                        Else
                           data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
                        End If
                        data_cli.Refresh
                     Else
                        If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
                           data_cli.RecordSource = "Select * from clientes where fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Xlehas & "'"
                        Else
                           data_cli.RecordSource = "Select * from clientes where fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
                        End If
                        data_cli.Refresh
                     End If
                  Else
                     md.Text = CDate("01/01/1980")
                     mh.Text = Date
                     If Len(Xcadsel) > 0 Then
                        If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
                           data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Xlehas & "'"
                        Else
                           data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
                        End If
                        data_cli.Refresh
                     Else
                        If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
                           data_cli.RecordSource = "Select * from clientes where fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Xlehas & "'"
                        Else
                           data_cli.RecordSource = "Select * from clientes where fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
                        End If
                        data_cli.Refresh
                     End If
                     md.Text = Date
                     mh.Text = Date
                     data_abm.RecordSource = "Select * from abmsocio where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and desc ='BAJA'"
                     data_abm.Refresh
                     md.Text = "__/__/____"
                     mh.Text = "__/__/____"
                  End If
               End If
            End If
            If data_cli.Recordset.RecordCount > 0 Then
               If Xop1 = 1 Then
                  Command5_Click
               Else
                  data_cli.Recordset.MoveLast
                  barr.Max = data_cli.Recordset.RecordCount
                  barr.Value = 0
                  data_cli.Recordset.MoveFirst
                  Do While Not data_cli.Recordset.EOF
                     If data_cli.Recordset("estado") = 3 Then
                        data_cli.Recordset.MoveNext
                     Else
                        data_inf.Recordset.AddNew
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                        If IsNull(data_cli.Recordset("cl_codced")) = False Then
                           If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                              data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                           End If
                        End If
                        data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                        data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                        data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                        data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                        data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                        data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                        data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                        data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                        data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                        data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                        data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                        data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                        data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                        data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                        data_inf.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
                        data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                        data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                        data_inf.Recordset("estado") = data_cli.Recordset("estado")
                        data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                        If data_cli.Recordset("cl_sexo") = 2 Then
                           data_inf.Recordset("cl_diacobr") = "FEMENINO"
                           data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                        Else
                           If data_cli.Recordset("cl_sexo") = 1 Then
                              data_inf.Recordset("cl_diacobr") = "MASCULINO"
                              data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                           Else
                              data_inf.Recordset("cl_diacobr") = "SIN DATO"
                              data_inf.Recordset("cl_sexo") = 3
                           End If
                        End If
                        If Check3.Value = 1 Then
'                           data_abm.Recordset.FindFirst "cl_codigo =" & data_cli.Recordset("cl_codigo")
                           data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo")
                           data_abm.Refresh
                           If data_abm.Recordset.RecordCount > 0 Then
                              data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                           Else
                              data_inf.Recordset("info_debit") = "SIN DATOS"
                           End If
                        End If
                        data_inf.Recordset.Update
                        data_cli.Recordset.MoveNext
                     End If
                     barr.Value = barr.Value + 1
                  Loop
                  DoEvents
                  Command4.Visible = True
                  Command4_Click
                  Command4.Visible = False
               End If
            Else
               MsgBox "No existen registros con ésta selección", vbInformation, "Informes de socios"
               
            End If
         End If
      End If
   End If
End If
Command1.Enabled = True
Command2.Enabled = True
Xcadsel = ""
frm_infabmvar.MousePointer = 0
barr.Visible = False

End Sub

Private Sub Command2_Click()
Unload Me

End Sub


Private Sub Command3_Click()
If Combo1.Text = "CONVENIO" Or Combo2.Text = "CONVENIO" Or Combo3.Text = "CONVENIO" Or Combo1.Text = "MUTUALISTA" Then
   Xop1 = 1
Else
   Xop1 = 0
End If
If Combo1.Text = "COBRADOR" Or Combo2.Text = "COBRADOR" Or Combo3.Text = "COBRADOR" Then
   Xop2 = 1
Else
   Xop2 = 0
End If
If Combo1.Text = "PROMOTOR" Or Combo2.Text = "PROMOTOR" Or Combo3.Text = "PROMOTOR" Then
   Xop3 = 1
Else
   Xop3 = 0
End If
If Combo1.Text = "RADIO" Or Combo2.Text = "RADIO" Or Combo3.Text = "RADIO" Then
   Xop4 = 1
Else
   Xop4 = 0
End If
If Combo1.Text = "SEXO" Or Combo2.Text = "SEXO" Or Combo3.Text = "SEXO" Then
   Xop5 = 1
   Xnrosexs = InputBox("Ingrese SEXO (0=TODOS, 1=MASC,2=FEM):", "Selección de Sexo")
   If Xnrosexs = "" Then
      Xnrosex = 0
   Else
      Xnrosex = Val(Xnrosexs)
   End If
Else
   Xop5 = 0
End If
If Combo1.Text = "EDAD" Or Combo2.Text = "EDAD" Or Combo3.Text = "EDAD" Then
   Xop6 = 1
Else
   Xop6 = 0
End If

If Xop2 = 1 Then
   If Wopscob <> 0 Then
      Xcadsel = "cl_nrocobr =" & Wopscob
   End If
End If

If Xop3 = 1 Then
   If Wopspro <> 0 Then
      If Len(Xcadsel) > 0 Then
         Xcadsel = Xcadsel + " And cl_nrovend =" & Wopspro
      Else
         Xcadsel = "cl_nrovend =" & Wopspro
      End If
   End If
End If

If Xop4 = 1 Then
'   If Xnrorad <> 0 Then
   If Wopszon >= 1 Then
      If Len(Xcadsel) > 0 Then
         If Wopszon = 1 Then
            Xcadsel = Xcadsel + " And cl_grupo >=" & 101 & " And cl_grupo <=" & 104
         End If
         If Wopszon = 2 Then
            Xcadsel = Xcadsel + " And cl_grupo >=" & 201 & " And cl_grupo <=" & 209
         End If
         If Wopszon = 3 Then
            Xcadsel = Xcadsel + " And cl_grupo >=" & 301 & " And cl_grupo <=" & 312
         End If
         If Wopszon = 4 Then
            Xcadsel = Xcadsel + " And cl_grupo >=" & 401 & " And cl_grupo <=" & 419
         End If
         If Wopszon = 5 Then
            Xcadsel = Xcadsel + " And cl_grupo IN (500,501)"
         End If
         If Wopszon = 6 Then
            Xcadsel = Xcadsel + " And cl_grupo >=" & 600 & " And cl_grupo <=" & 624
         End If
         If Wopszon = 7 Then
            Xcadsel = Xcadsel + " And cl_grupo >=" & 700 & " And cl_grupo <=" & 722
         End If
         If Wopszon = 8 Then
            Xcadsel = Xcadsel + " And cl_grupo =" & 800 & " or cl_grupo =" & 650
         End If
         If Wopszon = 9 Then
            Xcadsel = Xcadsel + " And cl_grupo >=" & 801 & " And cl_grupo >=" & 803
         End If
         If Wopszon = 10 Then
            Xcadsel = Xcadsel + " And cl_grupo =" & 810
         End If
         If Wopszon = 11 Then
            Xcadsel = Xcadsel + " And cl_grupo >=" & 630 & " And cl_grupo >=" & 640
         End If
         If Wopszon = 12 Then
            Xcadsel = Xcadsel + " And cl_grupo =" & 671
         End If
         If Wopszon = 13 Then
            Xcadsel = Xcadsel + " And cl_grupo =" & 670 & " or cl_grupo =" & 672 & " or cl_grupo =" & 673 & " or cl_grupo =" & 674
         End If
         If Wopszon = 15 Then
            Xcadsel = Xcadsel + " And cl_grupo =" & 815 & " or cl_grupo =" & 816
         End If
      Else
         If Wopszon = 1 Then
            Xcadsel = "cl_grupo >=" & 101 & " And cl_grupo <=" & 104
         End If
         If Wopszon = 2 Then
            Xcadsel = "cl_grupo >=" & 201 & " And cl_grupo <=" & 209
         End If
         If Wopszon = 3 Then
            Xcadsel = "cl_grupo >=" & 301 & " And cl_grupo <=" & 312
         End If
         If Wopszon = 4 Then
            Xcadsel = "cl_grupo >=" & 401 & " And cl_grupo <=" & 419
         End If
         If Wopszon = 5 Then
            Xcadsel = "cl_grupo >=" & 500 & " And cl_grupo <=" & 501
         End If
         If Wopszon = 6 Then
            Xcadsel = "cl_grupo >=" & 600 & " And cl_grupo <=" & 624
         End If
         If Wopszon = 7 Then
            Xcadsel = "cl_grupo >=" & 700 & " And cl_grupo <=" & 722
         End If
         If Wopszon = 8 Then
            Xcadsel = "cl_grupo =" & 800 & " or cl_grupo =" & 650
         End If
         If Wopszon = 9 Then
            Xcadsel = "cl_grupo >=" & 801 & " And cl_grupo <=" & 803
         End If
         If Wopszon = 10 Then
            Xcadsel = "cl_grupo =" & 810
         End If
         If Wopszon = 11 Then
            Xcadsel = "cl_grupo >=" & 630 & " And cl_grupo <=" & 640
         End If
         If Wopszon = 12 Then
            Xcadsel = "cl_grupo =" & 671
         End If
         If Wopszon = 13 Then
            Xcadsel = "cl_grupo =" & 670 & " or cl_grupo =" & 672 & " or cl_grupo =" & 673 & " or cl_grupo =" & 674
         End If
         If Wopszon = 15 Then
            Xcadsel = "cl_grupo >=" & 815 & " or cl_grupo <=" & 816
         End If
      End If
   End If
End If

If Xop5 = 1 Then 'Sexo
   If Xnrosex <> 0 Then
      If Len(Xcadsel) > 0 Then
         Xcadsel = Xcadsel + " And cl_sexo =" & Xnrosex
      Else
         Xcadsel = "cl_sexo =" & Xnrosex
      End If
   End If
End If

If Xop6 = 1 Then 'Edades
   If Wopsed = 1 Or Wopsed = 3 Then
      If Wopsedd = 0 And Wopsedh = 110 Then
      Else
         Xdiad = Wopsedd * 365
         Xdiad = Xdiad + 365
         Xdiah = Wopsedh * 365
         Xdiah = Xdiah + 365
         If Wopsedd > 4 Then
            Xdiabis = Wopsedd / 4
            Xdiabis = Int(Xdiabis)
            Xdiad = Xdiad + Xdiabis
         End If
         Xdiabis = 0
         If Wopsedh > 4 Then
            Xdiabis = Wopsedh / 4
            Xdiabis = Int(Xdiabis)
            Xdiah = Xdiah + Xdiabis
         End If
         Xfecedd = Date - Xdiad
         Xfecedh = Date - Xdiah
         If Wopsedd = 0 Then
            Xfecedd = Date
         End If
         If Len(Xcadsel) > 0 Then
            Xcadsel = Xcadsel + " And cl_fnac >='" & Format(Xfecedh, "yyyy-mm-dd") & "' And cl_fnac <='" & Format(Xfecedd, "yyyy-mm-dd") & "'"
         Else
            Xcadsel = "cl_fnac >='" & Format(Xfecedh, "yyyy-mm-dd") & "' And cl_fnac <='" & Format(Xfecedd, "yyyy-mm-dd") & "'"
         End If
      End If
   End If
   If Wopsed = 2 Then
      If Len(Xcadsel) > 0 Then
         Xcadsel = Xcadsel + " And month(cl_fnac) =" & Wopsedd
      Else
         Xcadsel = "month(cl_fnac) =" & Wopsedd
      End If
   End If
End If
'''aca
If Option1.Value = True Then
   If md.Text <> "__/__/____" And mh.Text <> "__/__/____" Then
      If Xop1 = 1 Then
         data_abm.RecordSource = "Select * from abmsocio where fecha >='" & Format(Date, "yyyy-mm-dd") & "' and fecha <='" & Format(Date, "yyyy-mm-dd") & "'"
'      data_abm.RecordSource = "Select * from abmsocio where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and desc ='" & "BAJA" & "'"
      Else
         data_abm.RecordSource = "Select * from abmsocio where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
      End If
      data_abm.Refresh
      If Len(Xcadsel) > 0 Then
         If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
            If Xop1 = 1 Then
               data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Wopsconvd & "'"
            Else
               data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Xlehas & "'"
            End If
         Else
'Wopsconvd
            If Xop1 = 1 Then
               data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_codconv ='" & Wopsconvd & "'"
            Else
               data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
            End If
         End If
         data_cli.Refresh
      Else
         If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
            data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Xlehas & "'"
         Else
            If Xop1 = 1 Then
               If Wopsconvd = "TODOS" Then
                  data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
               Else
                  data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_codconv ='" & Wopsconvd & "'"
               End If
            Else
               data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
            End If
         End If
         data_cli.Refresh
      End If
   Else
      md.Text = CDate("01/01/1980")
      mh.Text = Date
      If Len(Xcadsel) > 0 Then
         If Len(xlendes) > 0 And Len(xlenhas) > 0 Then
            data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Xlehas & "'"
         Else
            data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
         End If
         data_cli.Refresh
      Else
         If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
            data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Xlehas & "'"
         Else
            data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And cl_fecing >='" & Format(md.Text, "yyyy-mm-dd") & "' And cl_fecing <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
         End If
         data_cli.Refresh
      End If
      md.Text = Date
      mh.Text = Date
      data_abm.RecordSource = "Select * from abmsocio where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and desc ='BAJA'"
      data_abm.Refresh
      md.Text = "__/__/____"
      mh.Text = "__/__/____"
   End If
Else
   If Option2.Value = True Then
      If md.Text <> "__/__/____" And mh.Text <> "__/__/____" Then
         data_abm.RecordSource = "Select * from abmsocio where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and abmsocio.desc in ('BAJA')"
         data_abm.Refresh
         If Len(Xcadsel) > 0 Then
            If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
               data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv >='" & Xlehas & "'"
            Else
               data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
            End If
            data_cli.Refresh
         Else
            If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
               data_cli.RecordSource = "Select * from clientes where fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Xlehas & "'"
            Else
               data_cli.RecordSource = "Select * from clientes where fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
            End If
            data_cli.Refresh
         End If
      Else
         md.Text = CDate("01/01/1980")
         mh.Text = Date
         If Len(Xcadsel) > 0 Then
            If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
               data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Xlehas & "'"
            Else
               data_cli.RecordSource = "Select * from clientes where " & Xcadsel & " And fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
            End If
            data_cli.Refresh
         Else
            If Len(Xledes) > 0 And Len(Xlehas) > 0 Then
               data_cli.RecordSource = "Select * from clientes where And fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cl_codconv >='" & Xledes & "' And cl_codconv <='" & Xlehas & "'"
            Else
               data_cli.RecordSource = "Select * from clientes where And fecha_baja >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha_baja <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
            End If
            data_cli.Refresh
         End If
         md.Text = Date
         mh.Text = Date
         data_abm.RecordSource = "Select * from abmsocio where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and desc ='BAJA'"
         data_abm.Refresh
         md.Text = "__/__/____"
         mh.Text = "__/__/____"
      End If
   End If
End If
If data_cli.Recordset.RecordCount > 0 Then
   If Xop1 = 1 Then
      Command6_Click
   Else
      data_cli.Recordset.MoveLast
      barr.Max = data_cli.Recordset.RecordCount
      barr.Value = 0
      data_cli.Recordset.MoveFirst
      Do While Not data_cli.Recordset.EOF
         If data_cli.Recordset("estado") = 3 Then
            data_cli.Recordset.MoveNext
         Else
            data_inf.Recordset.AddNew
            data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
            data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
            If IsNull(data_cli.Recordset("cl_codced")) = False Then
               If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                  data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
               End If
            End If
            data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
            data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
            data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
            data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
            data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
            data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
            data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
            data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
            data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
            data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
            data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
            data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
            data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
            data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
            data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
            data_inf.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
            data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
            data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
            data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
            data_inf.Recordset("estado") = data_cli.Recordset("estado")
            data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
            If data_cli.Recordset("cl_sexo") = 2 Then
               data_inf.Recordset("cl_diacobr") = "FEMENINO"
               data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
            Else
               If data_cli.Recordset("cl_sexo") = 1 Then
                  data_inf.Recordset("cl_diacobr") = "MASCULINO"
                  data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
               Else
                  data_inf.Recordset("cl_diacobr") = "SIN DATO"
                  data_inf.Recordset("cl_sexo") = 3
               End If
            End If
            If Check3.Value = 1 Then
'               data_abm.Recordset.FindFirst "cl_codigo =" & data_cli.Recordset("cl_codigo")
               data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo")
               data_abm.Refresh
               If data_abm.Recordset.RecordCount > 0 Then
                  data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
               Else
                  data_inf.Recordset("info_debit") = "SIN DATOS"
               End If
            End If
            data_inf.Recordset.Update
            data_cli.Recordset.MoveNext
         End If
         barr.Value = barr.Value + 1
      Loop
      DoEvents
   End If
   Command4.Visible = True
   Command4_Click
   Command4.Visible = False
Else
   MsgBox "No existen registros", vbInformation, "Informes"
End If
barr.Visible = False

End Sub

Private Sub Command4_Click()
data_inf.RecordSource = "Select * from infcli"
data_inf.Refresh

If Check1.Value = 1 Then
   If Combo1.Text = "MUTUALISTA" Then
      reporte.ReportFileName = App.path & "\infvabpatd.rpt"
      reporte.ReportTitle = "INFORME DE SOCIOS POR PRESTADOR: " & Combo2.Text & " CON DATOS PATRONIMICOS  FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
      reporte.Action = 1
   Else
      reporte.ReportFileName = App.path & "\infvabpatd.rpt"
      reporte.ReportTitle = "INFORME DE SOCIOS CON DATOS PATRONIMICOS  FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
      reporte.Action = 1
   End If
Else
   If Check2.Value = 1 Then
      reporte.ReportFileName = App.path & "\infvabprod.rpt"
      reporte.ReportTitle = "INFORME DE SOCIOS CON DEUDA ACTUAL  FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
      reporte.Action = 1
   Else
      If Check3.Value = 1 Then
         If Option3.Value = True Then
            reporte.ReportFileName = App.path & "\infvabmotn.rpt"
            reporte.ReportTitle = "INFORME DE SOCIOS CON MOTIVO DE BAJA FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
            reporte.Action = 1
         Else
            reporte.ReportFileName = App.path & "\infvabmotd.rpt"
            reporte.ReportTitle = "INFORME DE SOCIOS CON MOTIVO DE BAJA FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
            reporte.Action = 1
         End If
      Else
        If Xop1 = 1 Then
           If Combo1.Text = "MUTUALISTA" Then
              If Option3.Value = True Then
                 reporte.ReportFileName = App.path & "\infvabcnvn.rpt"
                 reporte.ReportTitle = "INFORME DE SOCIOS POR PRESTADOR: " & Combo2.Text & " FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
                 reporte.Action = 1
              Else
                 reporte.ReportFileName = App.path & "\infvabcnvd.rpt"
                 reporte.ReportTitle = "INFORME DE SOCIOS POR PRESTADOR: " & Combo2.Text & " FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
                 reporte.Action = 1
              End If
           Else
              If Option3.Value = True Then
                 reporte.ReportFileName = App.path & "\infvabcnvn.rpt"
                 reporte.ReportTitle = "INFORME DE SOCIOS ORDENADOS POR CONVENIO  FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
                 reporte.Action = 1
              Else
                 reporte.ReportFileName = App.path & "\infvabcnvd.rpt"
                 reporte.ReportTitle = "INFORME DE SOCIOS ORDENADOS POR CONVENIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
                 reporte.Action = 1
              End If
           End If
        Else
           If Xop2 = 1 Then
              If Option3.Value = True Then
                 reporte.ReportFileName = App.path & "\infvabcobn.rpt"
                 reporte.ReportTitle = "INFORME DE SOCIOS ORDENADOS POR COBRADOR  FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
                 reporte.Action = 1
              Else
                 reporte.ReportFileName = App.path & "\infvabcobd.rpt"
                 reporte.ReportTitle = "INFORME DE SOCIOS ORDENADOS POR COBRADOR FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
                 reporte.Action = 1
              End If
           Else
              If Xop3 = 1 Then
                 If Option3.Value = True Then
                    reporte.ReportFileName = App.path & "\infvabpron.rpt"
                    reporte.ReportTitle = "INFORME DE SOCIOS ORDENADOS POR PROMOTOR  FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
                    reporte.Action = 1
                 Else
                    reporte.ReportFileName = App.path & "\infvabprod.rpt"
                    reporte.ReportTitle = "INFORME DE SOCIOS ORDENADOS POR PROMOTOR FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
                    reporte.Action = 1
                 End If
              Else
                 If Xop4 = 1 Then
                    If Option3.Value = True Then
                      reporte.ReportFileName = App.path & "\infvabradn.rpt"
                      reporte.ReportTitle = "INFORME DE SOCIOS ORDENADOS POR RADIOS  FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
                      reporte.Action = 1
                    Else
                      reporte.ReportFileName = App.path & "\infvabradd.rpt"
                      reporte.ReportTitle = "INFORME DE SOCIOS ORDENADOS POR RADIOS  FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
                      reporte.Action = 1
                    End If
                 Else
                    If Xop5 = 1 Then
                       If Option3.Value = True Then
                          reporte.ReportFileName = App.path & "\infvabsexn.rpt"
                          reporte.ReportTitle = "INFORME DE SOCIOS ORDENADOS POR SEXO  FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
                          reporte.Action = 1
                       Else
                          reporte.ReportFileName = App.path & "\infvabsexd.rpt"
                          reporte.ReportTitle = "INFORME DE SOCIOS ORDENADOS POR SEXO  FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
                          reporte.Action = 1
                       End If
                    Else
                       If Xop6 = 1 Then
                          If Option3.Value = True Then
                             reporte.ReportFileName = App.path & "\infvabedan.rpt"
                             reporte.ReportTitle = "INFORME DE SOCIOS ORDENADOS POR EDAD  FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
                             reporte.Action = 1
                          Else
                             reporte.ReportFileName = App.path & "\infvabedad.rpt"
                             reporte.ReportTitle = "INFORME DE SOCIOS ORDENADOS POR EDAD  FECHA: " & Format(md.Text, "dd/mm/yyyy") & " AL " & Format(mh.Text, "dd/mm/yyyy")
                             reporte.Action = 1
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

End Sub

Private Sub Command5_Click()
barr.Visible = True
barr.Value = 0
If data_cli.Recordset.RecordCount > 0 Then
   data_cli.Recordset.MoveLast
   barr.Max = data_cli.Recordset.RecordCount + 10
   barr.Value = 0
   data_cli.Recordset.MoveFirst
End If
If Xop1 = 1 Then
   Do While Not data_cli.Recordset.EOF
      If data_cli.Recordset("estado") = 3 Then
         data_cli.Recordset.MoveNext
      Else
         If Wopsconv = 1 Then 'Todos los convenios
            data_inf.Recordset.AddNew
            data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
            data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
            If IsNull(data_cli.Recordset("cl_codced")) = False Then
               If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                  data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
               End If
            End If
            data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
            data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
            data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
            data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
            data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
            data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
            data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
            data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
            data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
            data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
            data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
            data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
            data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
            data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
            data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
            data_inf.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
            data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
            data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
            data_inf.Recordset("estado") = data_cli.Recordset("estado")
            data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
            If data_cli.Recordset("cl_sexo") = 2 Then
               data_inf.Recordset("cl_diacobr") = "FEMENINO"
               data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
            Else
               If data_cli.Recordset("cl_sexo") = 1 Then
                  data_inf.Recordset("cl_diacobr") = "MASCULINO"
                  data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
               Else
                  data_inf.Recordset("cl_diacobr") = "SIN DATO"
                  data_inf.Recordset("cl_sexo") = 3
               End If
            End If
            If Check3.Value = 1 Then
'               data_abm.Recordset.FindFirst "cl_codigo =" & data_cli.Recordset("cl_codigo")
               data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo")
               data_abm.Refresh
               If data_abm.Recordset.RecordCount > 0 Then
                  data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
               Else
                  data_inf.Recordset("info_debit") = "SIN DATOS"
               End If
            End If
            data_inf.Recordset.Update
            data_cli.Recordset.MoveNext
         Else
            If Wopsconv = 2 Then 'Mutuales todos
'               data_conv.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Wopsconvd Then
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     If IsNull(data_cli.Recordset("cl_codced")) = False Then
                        If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                           data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                        End If
                     End If
                     data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                     data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                     data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                     data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                     data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                     data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                     data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                     data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                     data_inf.Recordset("cl_nomconv") = Mid(data_conv.Recordset("cnv_desc"), 1, 30)
                     data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                     data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                     data_inf.Recordset("estado") = data_cli.Recordset("estado")
                     data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                     If data_cli.Recordset("cl_sexo") = 2 Then
                        data_inf.Recordset("cl_diacobr") = "FEMENINO"
                        data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                     Else
                        If data_cli.Recordset("cl_sexo") = 1 Then
                           data_inf.Recordset("cl_diacobr") = "MASCULINO"
                           data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                        Else
                           data_inf.Recordset("cl_diacobr") = "SIN DATO"
                           data_inf.Recordset("cl_sexo") = 3
                        End If
                     End If
                     If Check3.Value = 1 Then
                        data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo")
                        data_abm.Refresh
                        If data_abm.Recordset.RecordCount > 0 Then
                           data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                        Else
                           data_inf.Recordset("info_debit") = "SIN DATOS"
                        End If
                     End If
                     data_inf.Recordset.Update
                     data_cli.Recordset.MoveNext
                  Else
                     data_cli.Recordset.MoveNext
                  End If
               Else
                  data_cli.Recordset.MoveNext
               End If
            Else
               If Wopsconv = 5 Then 'Mutuales con complementos
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                  data_conv.Refresh
                  If data_conv.Recordset.RecordCount > 0 Then
                     If data_conv.Recordset("cnv_grupo") = Wopsconvd And data_conv.Recordset("cnv_precio") <> 0 Then
                        data_inf.Recordset.AddNew
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                        If IsNull(data_cli.Recordset("cl_codced")) = False Then
                           If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                              data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                           End If
                        End If
                        data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                        data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                        data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                        data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                        data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                        data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                        data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                        data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                        data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                        data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                        data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                        data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                        data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                        data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_inf.Recordset("cl_nomconv") = Mid(data_conv.Recordset("cnv_desc"), 1, 30)
                        data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                        data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                        data_inf.Recordset("estado") = data_cli.Recordset("estado")
                        data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                        data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                        If data_cli.Recordset("cl_sexo") = 2 Then
                           data_inf.Recordset("cl_diacobr") = "FEMENINO"
                           data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                        Else
                           If data_cli.Recordset("cl_sexo") = 1 Then
                              data_inf.Recordset("cl_diacobr") = "MASCULINO"
                              data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                           Else
                              data_inf.Recordset("cl_diacobr") = "SIN DATO"
                              data_inf.Recordset("cl_sexo") = 3
                           End If
                        End If
                        If Check3.Value = 1 Then
                           data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo")
                           data_abm.Refresh
                           If data_abm.Recordset.RecordCount > 0 Then
                              data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                           Else
                              data_inf.Recordset("info_debit") = "SIN DATOS"
                           End If
                        End If
                        data_inf.Recordset.Update
                        data_cli.Recordset.MoveNext
                     Else
                        data_cli.Recordset.MoveNext
                     End If
                  Else
                     data_cli.Recordset.MoveNext
                  End If
               Else
                  If Wopsconv = 6 Then 'Mutuales sin complemento
                     data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.Refresh
                     If data_conv.Recordset.RecordCount > 0 Then
                        If data_conv.Recordset("cnv_grupo") = Wopsconvd And data_conv.Recordset("cnv_precio") = 0 Then
                           data_inf.Recordset.AddNew
                           data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                           data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                           If IsNull(data_cli.Recordset("cl_codced")) = False Then
                              If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                 data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                              End If
                           End If
                           data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                           data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                           data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                           data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                           data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                           data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                           data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                           data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                           data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                           data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                           data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                           data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                           data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                           data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                           data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                           data_inf.Recordset("cl_nomconv") = Mid(data_conv.Recordset("cnv_desc"), 1, 30)
                           data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                           data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                           data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                           data_inf.Recordset("estado") = data_cli.Recordset("estado")
                           data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                           If data_cli.Recordset("cl_sexo") = 2 Then
                              data_inf.Recordset("cl_diacobr") = "FEMENINO"
                              data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                           Else
                              If data_cli.Recordset("cl_sexo") = 1 Then
                                 data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                 data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                              Else
                                 data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                 data_inf.Recordset("cl_sexo") = 3
                              End If
                           End If
                           If Check3.Value = 1 Then
                              data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo")
                              data_abm.Refresh
                              If data_abm.Recordset.RecordCount > 0 Then
                                 data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                              Else
                                 data_inf.Recordset("info_debit") = "SIN DATOS"
                              End If
                           End If
                           data_inf.Recordset.Update
                           data_cli.Recordset.MoveNext
                        Else
                           data_cli.Recordset.MoveNext
                        End If
                     Else
                        data_cli.Recordset.MoveNext
                     End If
                  Else
                     If Wopsconv = 3 Then 'Grupos de sapp
                        data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                        data_conv.Refresh
                        If data_conv.Recordset.RecordCount > 0 Then
                           If Wopsconvd = "TODOS" Then
                              If data_conv.Recordset("cnv_colrec") = "M" Or data_conv.Recordset("cnv_colrec") = "V" Or data_conv.Recordset("cnv_colrec") = "R" Or data_conv.Recordset("cnv_colrec") = "A" Then
                                 data_inf.Recordset.AddNew
                                 data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                 data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                 If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                    If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                       data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                    End If
                                 End If
                                 data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                 data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                                 data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                                 data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                                 data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                 data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                 data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                                 data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                                 data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                                 data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                                 data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                                 data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                                 data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                                 data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                                 data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                                 data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                 data_inf.Recordset("cl_nomconv") = Mid(data_conv.Recordset("cnv_desc"), 1, 30)
                                 data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                                 data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                                 data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                 data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                                 If data_cli.Recordset("cl_sexo") = 2 Then
                                    data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                    data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                 Else
                                    If data_cli.Recordset("cl_sexo") = 1 Then
                                       data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                       data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                    Else
                                       data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                       data_inf.Recordset("cl_sexo") = 3
                                    End If
                                 End If
                                 If Check3.Value = 1 Then
                                    data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo")
                                    data_abm.Refresh
                                    If data_abm.Recordset.RecordCount > 0 Then
                                       data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                    Else
                                       data_inf.Recordset("info_debit") = "SIN DATOS"
                                    End If
                                 End If
                                 data_inf.Recordset.Update
                                 data_cli.Recordset.MoveNext
                              Else
                                 data_cli.Recordset.MoveNext
                              End If
                           End If
                           If Wopsconvd = "AMBULATORIO" Then
                              If data_conv.Recordset("cnv_colrec") = "R" Then
                                 data_inf.Recordset.AddNew
                                 data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                 data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                 If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                    If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                       data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                    End If
                                 End If
                                 data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                 data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                                 data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                                 data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                                 data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                 data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                 data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                                 data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                                 data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                                 data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                                 data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                                 data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                                 data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                                 data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                                 data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                                 data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                 data_inf.Recordset("cl_nomconv") = Mid(data_conv.Recordset("cnv_desc"), 1, 30)
                                 data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                                 data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                                 data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                 data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                                 If data_cli.Recordset("cl_sexo") = 2 Then
                                    data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                    data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                 Else
                                    If data_cli.Recordset("cl_sexo") = 1 Then
                                       data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                       data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                    Else
                                       data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                       data_inf.Recordset("cl_sexo") = 3
                                    End If
                                 End If
                                 If Check3.Value = 1 Then
                                    data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo")
                                    data_abm.Refresh
                                    If data_abm.Recordset.RecordCount > 0 Then
                                       data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                    Else
                                       data_inf.Recordset("info_debit") = "SIN DATOS"
                                    End If
                                 End If
                                 data_inf.Recordset.Update
                                 data_cli.Recordset.MoveNext
                              Else
                                 data_cli.Recordset.MoveNext
                              End If
                           End If
                           If Wopsconvd = "EMERGENCIA" Then
                              If data_conv.Recordset("cnv_codigo") = "EMERN" Or _
                                 data_conv.Recordset("cnv_codigo") = "EMERC" Or _
                                 data_conv.Recordset("cnv_codigo") = "EMERF" Or _
                                 data_conv.Recordset("cnv_codigo") = "EMERG" Or _
                                 data_conv.Recordset("cnv_codigo") = "EMERJ" Or _
                                 data_conv.Recordset("cnv_codigo") = "EMERNE" Or _
                                 data_conv.Recordset("cnv_codigo") = "EMERNT" Or _
                                 data_conv.Recordset("cnv_codigo") = "EMERSA" Or _
                                 data_conv.Recordset("cnv_codigo") = "CASA1" Or _
                                 data_conv.Recordset("cnv_codigo") = "CASA6" Then
                                 data_inf.Recordset.AddNew
                                 data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                 data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                 If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                    If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                       data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                    End If
                                 End If
                                 data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                 data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                                 data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                                 data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                                 data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                 data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                 data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                                 data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                                 data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                                 data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                                 data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                                 data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                                 data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                                 data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                                 data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                                 data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                 data_inf.Recordset("cl_nomconv") = Mid(data_conv.Recordset("cnv_desc"), 1, 30)
                                 data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                                 data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                                 data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                 data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                                 If data_cli.Recordset("cl_sexo") = 2 Then
                                    data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                    data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                 Else
                                    If data_cli.Recordset("cl_sexo") = 1 Then
                                       data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                       data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                    Else
                                       data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                       data_inf.Recordset("cl_sexo") = 3
                                    End If
                                 End If
                                 If Check3.Value = 1 Then
'                                    data_abm.Recordset.FindFirst "cl_codigo =" & data_cli.Recordset("cl_codigo")
                                    data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo")
                                    data_abm.Refresh
                                    If data_abm.Recordset.RecordCount > 0 Then
                                       data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                    Else
                                       data_inf.Recordset("info_debit") = "SIN DATOS"
                                    End If
                                 End If
                                 data_inf.Recordset.Update
                                 data_cli.Recordset.MoveNext
                              Else
                                 data_cli.Recordset.MoveNext
                              End If
                           End If
                           If Wopsconvd = "AREAS P." Then
                              If data_conv.Recordset("cnv_colrec") = "M" And _
                                 data_conv.Recordset("cnv_cant_r") <> 2 And _
                                 data_cli.Recordset("cl_fecing") <= data_conv.Recordset("cnv_desde") Then
                                 data_inf.Recordset.AddNew
                                 data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                 data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                 If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                    If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                       data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                    End If
                                 End If
                                 data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                 data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                                 data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                                 data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                                 data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                 data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                 data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                                 data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                                 data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                                 data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                                 data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                                 data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                                 data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                                 data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                                 data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                                 data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                 data_inf.Recordset("cl_nomconv") = Mid(data_conv.Recordset("cnv_desc"), 1, 30)
                                 data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                                 data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                                 data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                 data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                                 If data_cli.Recordset("cl_sexo") = 2 Then
                                    data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                    data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                 Else
                                    If data_cli.Recordset("cl_sexo") = 1 Then
                                       data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                       data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                    Else
                                       data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                       data_inf.Recordset("cl_sexo") = 3
                                    End If
                                 End If
                                 If Check3.Value = 1 Then
                                    data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo")
                                    data_abm.Refresh
                                    If data_abm.Recordset.RecordCount > 0 Then
                                       data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                    Else
                                       data_inf.Recordset("info_debit") = "SIN DATOS"
                                    End If
                                 End If
                                 data_inf.Recordset.Update
                                 data_cli.Recordset.MoveNext
                              Else
                                 data_cli.Recordset.MoveNext
                              End If
                           End If
                           If Wopsconvd = "PARCIAL" Then
                              If data_conv.Recordset("cnv_colrec") = "A" And _
                                 data_conv.Recordset("cnv_grupo") = "" Then
                                 data_inf.Recordset.AddNew
                                 data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                 data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                 If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                    If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                        data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                    End If
                                 End If
                                 data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                 data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                                 data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                                 data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                                 data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                 data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                 data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                                 data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                                 data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                                 data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                                 data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                                 data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                                 data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                                 data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                                 data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                                 data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                 data_inf.Recordset("cl_nomconv") = Mid(data_conv.Recordset("cnv_desc"), 1, 30)
                                 data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                                 data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                                 data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                 data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                                 If data_cli.Recordset("cl_sexo") = 2 Then
                                    data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                    data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                 Else
                                    If data_cli.Recordset("cl_sexo") = 1 Then
                                       data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                       data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                    Else
                                       data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                       data_inf.Recordset("cl_sexo") = 3
                                    End If
                                 End If
                                 If Check3.Value = 1 Then
                                    data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo")
                                    data_abm.Refresh
                                    If data_abm.Recordset.RecordCount > 0 Then
                                       data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                    Else
                                       data_inf.Recordset("info_debit") = "SIN DATOS"
                                    End If
                                 End If
                                 data_inf.Recordset.Update
                                 data_cli.Recordset.MoveNext
                              Else
                                 data_cli.Recordset.MoveNext
                              End If
                           End If
                        Else
                           data_cli.Recordset.MoveNext
                        End If
                     Else
                        If Wopsconv = 4 Then 'Selección
                           If data_cli.Recordset("cl_codconv") = Wopsconvd Then
                              data_inf.Recordset.AddNew
                              data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                              data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                              If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                 If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                    data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                 End If
                              End If
                              data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                              data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                              data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                              data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                              data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                              data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                              data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                              data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                              data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                              data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                              data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                              data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                              data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                              data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                              data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                              data_inf.Recordset("cl_nomconv") = Mid(data_conv.Recordset("cnv_desc"), 1, 30)
                              data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                              data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                              data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                              data_inf.Recordset("estado") = data_cli.Recordset("estado")
                              data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                              If data_cli.Recordset("cl_sexo") = 2 Then
                                 data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                 data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                              Else
                                 If data_cli.Recordset("cl_sexo") = 1 Then
                                    data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                    data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                 Else
                                    data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                    data_inf.Recordset("cl_sexo") = 3
                                 End If
                              End If
                              If Check3.Value = 1 Then
                                 data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo")
                                 data_abm.Refresh
                                 If data_abm.Recordset.RecordCount > 0 Then
                                    data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                 Else
                                    data_inf.Recordset("info_debit") = "SIN DATOS"
                                 End If
                              End If
                              data_inf.Recordset.Update
                              data_cli.Recordset.MoveNext
                           Else
                              data_cli.Recordset.MoveNext
                           End If
                        Else
                           If Wopsconv = 9 Then 'Grupos de sapp sin complemento
'                              data_conv.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                              data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                              data_conv.Refresh
                              If data_conv.Recordset.RecordCount > 0 Then
                                 If Wopsconvd = "AMBULATORIO" Then
                                    If data_conv.Recordset("cnv_colrec") = "R" And data_conv.Recordset("cnv_grupo") = "" Then
                                       data_inf.Recordset.AddNew
                                       data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                       data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                       If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                          If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                             data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                          End If
                                       End If
                                       data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                       data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                                       data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                                       data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                                       data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                       data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                       data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                                       data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                                       data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                                       data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                                       data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                                       data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                                       data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                                       data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                                       data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                       data_inf.Recordset("cl_nomconv") = Mid(data_conv.Recordset("cnv_desc"), 1, 30)
                                       data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                                       data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                                       data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                                       data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                       data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                                       If data_cli.Recordset("cl_sexo") = 2 Then
                                          data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                          data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                       Else
                                          If data_cli.Recordset("cl_sexo") = 1 Then
                                             data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                             data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                          Else
                                             data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                             data_inf.Recordset("cl_sexo") = 3
                                          End If
                                       End If
                                       If Check3.Value = 1 Then
                                          data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo")
                                          data_abm.Refresh
                                          If data_abm.Recordset.RecordCount > 0 Then
                                             data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                          Else
                                             data_inf.Recordset("info_debit") = "SIN DATOS"
                                          End If
                                       End If
                                       data_inf.Recordset.Update
                                       data_cli.Recordset.MoveNext
                                    Else
                                       data_cli.Recordset.MoveNext
                                    End If
                                 Else
                                    data_cli.Recordset.MoveNext
                                 End If
                              Else
                                 data_cli.Recordset.MoveNext
                              End If
                           Else
                              If Combo1.Text = "MUTUALISTA" Then
                                    data_inf.Recordset.AddNew
                                    data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                    data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                    If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                       If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                          data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                       End If
                                    End If
                                    data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                    data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                                    data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                                    data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                                    data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                    data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                    data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                                    data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                                    data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                                    data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                                    data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                                    data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                                    data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                                    data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                                    data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                    data_inf.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
                                    data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                                    data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                                    data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                                    data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                    data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                                    If data_cli.Recordset("cl_sexo") = 2 Then
                                       data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                       data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                    Else
                                       If data_cli.Recordset("cl_sexo") = 1 Then
                                          data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                          data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                       Else
                                          data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                          data_inf.Recordset("cl_sexo") = 3
                                       End If
                                    End If
                                    data_inf.Recordset.Update
                                    data_cli.Recordset.MoveNext
                              
                              Else
                                 data_cli.Recordset.MoveNext
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
      barr.Value = barr.Value + 1
   Loop
End If
Command4_Click
barr.Visible = False

End Sub

Private Sub Command6_Click()
barr.Visible = True
barr.Value = 0

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\inftab.mdb")
MiBaseact.Execute "Delete * from convenio"

data_convloc.RecordSource = "convenio"
data_convloc.Refresh
data_conv.RecordSource = "convenio"
data_conv.Refresh

data_conv.Recordset.MoveFirst
If Wopsconv = 2 Or Wopsconv = 5 Or Wopsconv = 6 Then
    Do While Not data_conv.Recordset.EOF
       If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
          If data_conv.Recordset("cnv_grupo") = Wopsconvd Then
             data_convloc.Recordset.AddNew
             data_convloc.Recordset("cnv_codigo") = data_conv.Recordset("cnv_codigo")
             data_convloc.Recordset("cnv_desc") = data_conv.Recordset("cnv_desc")
             data_convloc.Recordset("cnv_desde") = data_conv.Recordset("cnv_desde")
             data_convloc.Recordset("cnv_hasta") = data_conv.Recordset("cnv_hasta")
             data_convloc.Recordset("cnv_colrec") = data_conv.Recordset("cnv_colrec")
             data_convloc.Recordset("cnv_precio") = data_conv.Recordset("cnv_precio")
             data_convloc.Recordset("cnv_emite") = data_conv.Recordset("cnv_emite")
             data_convloc.Recordset("cnv_alta") = data_conv.Recordset("cnv_alta")
             data_convloc.Recordset("cnv_cant_r") = data_conv.Recordset("cnv_cant_r")
             data_convloc.Recordset("cnv_grupo") = data_conv.Recordset("cnv_grupo")
             data_convloc.Recordset.Update
          End If
       End If
       data_conv.Recordset.MoveNext
   Loop
Else
    Do While Not data_conv.Recordset.EOF
       If IsNull(data_conv.Recordset("cnv_colrec")) = False Then
             data_convloc.Recordset.AddNew
             data_convloc.Recordset("cnv_codigo") = data_conv.Recordset("cnv_codigo")
             data_convloc.Recordset("cnv_desc") = data_conv.Recordset("cnv_desc")
             data_convloc.Recordset("cnv_desde") = data_conv.Recordset("cnv_desde")
             data_convloc.Recordset("cnv_hasta") = data_conv.Recordset("cnv_hasta")
             data_convloc.Recordset("cnv_colrec") = data_conv.Recordset("cnv_colrec")
             data_convloc.Recordset("cnv_precio") = data_conv.Recordset("cnv_precio")
             data_convloc.Recordset("cnv_emite") = data_conv.Recordset("cnv_emite")
             data_convloc.Recordset("cnv_alta") = data_conv.Recordset("cnv_alta")
             data_convloc.Recordset("cnv_cant_r") = data_conv.Recordset("cnv_cant_r")
             data_convloc.Recordset("cnv_grupo") = data_conv.Recordset("cnv_grupo")
             data_convloc.Recordset.Update
       End If
       data_conv.Recordset.MoveNext
   Loop
End If
If data_cli.Recordset.RecordCount > 0 Then
   data_cli.Recordset.MoveLast
   barr.Max = data_cli.Recordset.RecordCount
   barr.Value = 0
   data_cli.Recordset.MoveFirst
   Do While Not data_cli.Recordset.EOF
        If Xop1 = 1 Then
           If Wopsconv = 1 Then 'Todos los convenios
              data_inf.Recordset.AddNew
              data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
              data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
              If IsNull(data_cli.Recordset("cl_codced")) = False Then
                 If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                    data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                 End If
              End If
              data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
              data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
              data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
              data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
              data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
              data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
              data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
              data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
              data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
              data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
              data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
              data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
              data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
              data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
              data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
              data_inf.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
              data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
              data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
              data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
              data_inf.Recordset("estado") = data_cli.Recordset("estado")
              data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
              If data_cli.Recordset("cl_sexo") = 2 Then
                 data_inf.Recordset("cl_diacobr") = "FEMENINO"
                 data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
              Else
                 If data_cli.Recordset("cl_sexo") = 1 Then
                    data_inf.Recordset("cl_diacobr") = "MASCULINO"
                    data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                 Else
                    data_inf.Recordset("cl_diacobr") = "SIN DATO"
                    data_inf.Recordset("cl_sexo") = 3
                 End If
              End If
              If Check3.Value = 1 Then
'                 data_abm.Recordset.FindFirst "cl_codigo =" & data_cli.Recordset("cl_codigo") & " And fecha ='" & Format(data_cli.Recordset("fecha_baja"), "yyyy-mm-dd") & "'"
                 data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo") & " And fecha ='" & Format(data_cli.Recordset("fecha_baja"), "yyyy-mm-dd") & "'"
                 data_abm.Refresh
                 If data_abm.Recordset.RecordCount > 0 Then
                    data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                 Else
                    data_inf.Recordset("info_debit") = "SIN DATOS"
                 End If
              End If
              data_inf.Recordset.Update
              data_cli.Recordset.MoveNext
           Else
              If Wopsconv = 2 Then 'Mutuales todos
'                 data_convloc.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                 data_convloc.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                 data_convloc.Refresh
                 If data_convloc.Recordset.RecordCount > 0 Then
                    If data_convloc.Recordset("cnv_grupo") = Wopsconvd Then
                       data_inf.Recordset.AddNew
                       data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                       data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                       If IsNull(data_cli.Recordset("cl_codced")) = False Then
                          If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                             data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                          End If
                       End If
                       data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                       data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                       data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                       data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                       data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                       data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                       data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                       data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                       data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                       data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                       data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                       data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                       data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                       data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                       data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                       data_inf.Recordset("cl_nomconv") = Mid(data_convloc.Recordset("cnv_desc"), 1, 30)
                       data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                       data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                       data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                       data_inf.Recordset("estado") = data_cli.Recordset("estado")
                       data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                       If data_cli.Recordset("cl_sexo") = 2 Then
                          data_inf.Recordset("cl_diacobr") = "FEMENINO"
                          data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                       Else
                          If data_cli.Recordset("cl_sexo") = 1 Then
                             data_inf.Recordset("cl_diacobr") = "MASCULINO"
                             data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                          Else
                             data_inf.Recordset("cl_diacobr") = "SIN DATO"
                             data_inf.Recordset("cl_sexo") = 3
                          End If
                       End If
                       If Check3.Value = 1 Then
                          data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo") & " And fecha ='" & Format(data_cli.Recordset("fecha_baja"), "yyyy-mm-dd") & "'"
                          data_abm.Refresh
                          If data_abm.Recordset.RecordCount > 0 Then
                             data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                          Else
                             data_inf.Recordset("info_debit") = "SIN DATOS"
                          End If
                       End If
                       data_inf.Recordset.Update
                       data_cli.Recordset.MoveNext
                    Else
                       data_cli.Recordset.MoveNext
                    End If
                 Else
                    data_cli.Recordset.MoveNext
                 End If
              Else
                 If Wopsconv = 5 Then 'Mutuales con complementos
'                    data_convloc.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                    data_convloc.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                    data_convloc.Refresh
                    If data_convloc.Recordset.RecordCount > 0 Then
                       If data_convloc.Recordset("cnv_grupo") = Wopsconvd And data_convloc.Recordset("cnv_precio") <> 0 Then
                          data_inf.Recordset.AddNew
                          data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                          data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                          If IsNull(data_cli.Recordset("cl_codced")) = False Then
                             If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                             End If
                          End If
                          data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                          data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                          data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                          data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                          data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                          data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                          data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                          data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                          data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                          data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                          data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                          data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                          data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                          data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                          data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                          data_inf.Recordset("cl_nomconv") = Mid(data_convloc.Recordset("cnv_desc"), 1, 30)
                          data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                          data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                          data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                          data_inf.Recordset("estado") = data_cli.Recordset("estado")
                          data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                          If data_cli.Recordset("cl_sexo") = 2 Then
                             data_inf.Recordset("cl_diacobr") = "FEMENINO"
                             data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                          Else
                             If data_cli.Recordset("cl_sexo") = 1 Then
                                data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                             Else
                                data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                data_inf.Recordset("cl_sexo") = 3
                             End If
                          End If
                          If Check3.Value = 1 Then
                             data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo") & " And fecha ='" & Format(data_cli.Recordset("fecha_baja"), "yyyy-mm-dd") & "'"
                             data_abm.Refresh
                             If data_abm.Recordset.RecordCount > 0 Then
                                data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                             Else
                                data_inf.Recordset("info_debit") = "SIN DATOS"
                             End If
                          End If
                          data_inf.Recordset.Update
                          data_cli.Recordset.MoveNext
                       Else
                          data_cli.Recordset.MoveNext
                       End If
                    Else
                       data_cli.Recordset.MoveNext
                    End If
                 Else
                    If Wopsconv = 6 Then 'Mutuales sin complemento
'                       data_convloc.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                       data_convloc.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                       data_convloc.Refresh
                       If data_convloc.Recordset.RecordCount > 0 Then
                          If data_convloc.Recordset("cnv_grupo") = Wopsconvd And data_convloc.Recordset("cnv_precio") = 0 Then
                             data_inf.Recordset.AddNew
                             data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                             data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                             If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                    data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                End If
                             End If
                             data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                             data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                             data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                             data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                             data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                             data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                             data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                             data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                             data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                             data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                             data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                             data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                             data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                             data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                             data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                             data_inf.Recordset("cl_nomconv") = Mid(data_convloc.Recordset("cnv_desc"), 1, 30)
                             data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                             data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                             data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                             data_inf.Recordset("estado") = data_cli.Recordset("estado")
                             data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                             If data_cli.Recordset("cl_sexo") = 2 Then
                                data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                             Else
                                If data_cli.Recordset("cl_sexo") = 1 Then
                                   data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                   data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                Else
                                   data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                   data_inf.Recordset("cl_sexo") = 3
                                End If
                             End If
                             If Check3.Value = 1 Then
                                data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo") & " And fecha ='" & Format(data_cli.Recordset("fecha_baja"), "yyyy-mm-dd") & "'"
                                data_abm.Refresh
                                If data_abm.Recordset.RecordCount > 0 Then
                                   data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                Else
                                   data_inf.Recordset("info_debit") = "SIN DATOS"
                                End If
                             End If
                             data_inf.Recordset.Update
                             data_cli.Recordset.MoveNext
                          Else
                             data_cli.Recordset.MoveNext
                          End If
                       Else
                          data_cli.Recordset.MoveNext
                       End If
                    Else
                       If Wopsconv = 3 Then 'Grupos de sapp
                          data_convloc.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                          data_convloc.Refresh
                          If data_convloc.Recordset.RecordCount > 0 Then
                             If Wopsconvd = "TODOS" Then
                                If data_convloc.Recordset("cnv_colrec") = "M" Or data_convloc.Recordset("cnv_colrec") = "V" Or data_convloc.Recordset("cnv_colrec") = "R" Or data_convloc.Recordset("cnv_colrec") = "A" Then
                                   If data_convloc.Recordset("cnv_emite") = "SI" Then
                                       data_inf.Recordset.AddNew
                                       data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                       data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                       If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                          If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                             data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                          End If
                                       End If
                                       data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                       data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                                       data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                                       data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                                       data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                       data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                       data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                                       data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                                       data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                                       data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                                       data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                                       data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                                       data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                                       data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                                       data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                       data_inf.Recordset("cl_nomconv") = Mid(data_convloc.Recordset("cnv_desc"), 1, 30)
                                       data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                                       data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                                       data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                                       data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                       data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                                       If data_cli.Recordset("cl_sexo") = 2 Then
                                          data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                          data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                       Else
                                          If data_cli.Recordset("cl_sexo") = 1 Then
                                             data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                             data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                          Else
                                             data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                             data_inf.Recordset("cl_sexo") = 3
                                          End If
                                       End If
                                       If Check3.Value = 1 Then
                                          data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo") & " And fecha ='" & Format(data_cli.Recordset("fecha_baja"), "yyyy-mm-dd") & "'"
                                          data_abm.Refresh
                                          If data_abm.Recordset.RecordCount > 0 Then
                                             data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                          Else
                                             data_inf.Recordset("info_debit") = "SIN DATOS"
                                          End If
                                       End If
                                       data_inf.Recordset.Update
                                       data_cli.Recordset.MoveNext
                                    Else
                                       data_cli.Recordset.MoveNext
                                    End If
                                 Else
                                    data_cli.Recordset.MoveNext
                                 End If
                             End If
                             If Wopsconvd = "AMBULATORIO" Then
                                If data_convloc.Recordset("cnv_colrec") = "R" Then
                                   data_inf.Recordset.AddNew
                                   data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                   data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                   If Len(data_cli.Recordset("cl_codced")) = 1 Then
                                      data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                   End If
                                   data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                   data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                                   data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                                   data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                                   data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                   data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                   data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                                   data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                                   data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                                   data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                                   data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                                   data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                                   data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                                   data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                                   data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                   data_inf.Recordset("cl_nomconv") = Mid(data_convloc.Recordset("cnv_desc"), 1, 30)
                                   data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                                   data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                                   data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                                   data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                   data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                                   If data_cli.Recordset("cl_sexo") = 2 Then
                                      data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                      data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                   Else
                                      If data_cli.Recordset("cl_sexo") = 1 Then
                                         data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                         data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                      Else
                                         data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                         data_inf.Recordset("cl_sexo") = 3
                                      End If
                                   End If
                                   If Check3.Value = 1 Then
                                      data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo") & " And fecha ='" & Format(data_cli.Recordset("fecha_baja"), "yyyy-mm-dd") & "'"
                                      data_abm.Refresh
                                      If data_abm.Recordset.RecordCount > 0 Then
                                         data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                      Else
                                         data_inf.Recordset("info_debit") = "SIN DATOS"
                                      End If
                                   End If
                                   data_inf.Recordset.Update
                                   data_cli.Recordset.MoveNext
                                Else
                                   data_cli.Recordset.MoveNext
                                End If
                             End If
                             If Wopsconvd = "EMERGENCIA" Then
                                If data_convloc.Recordset("cnv_codigo") = "EMERN" Or _
                                   data_convloc.Recordset("cnv_codigo") = "EMERC" Or _
                                   data_convloc.Recordset("cnv_codigo") = "EMERF" Or _
                                   data_convloc.Recordset("cnv_codigo") = "EMERG" Or _
                                   data_convloc.Recordset("cnv_codigo") = "EMERJ" Or _
                                   data_convloc.Recordset("cnv_codigo") = "EMERNE" Or _
                                   data_convloc.Recordset("cnv_codigo") = "EMERNT" Or _
                                   data_convloc.Recordset("cnv_codigo") = "EMERSA" Or _
                                   data_convloc.Recordset("cnv_codigo") = "CASA1" Or _
                                   data_convloc.Recordset("cnv_codigo") = "CASA6" Then
                                   data_inf.Recordset.AddNew
                                   data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                   data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                   If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                      If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                         data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                      End If
                                   End If
                                   data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                   data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                                   data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                                   data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                                   data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                   data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                   data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                                   data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                                   data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                                   data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                                   data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                                   data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                                   data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                                   data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                                   data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                   data_inf.Recordset("cl_nomconv") = Mid(data_convloc.Recordset("cnv_desc"), 1, 30)
                                   data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                                   data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                                   data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                                   data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                   data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                                   If data_cli.Recordset("cl_sexo") = 2 Then
                                      data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                      data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                   Else
                                      If data_cli.Recordset("cl_sexo") = 1 Then
                                         data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                         data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                      Else
                                         data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                         data_inf.Recordset("cl_sexo") = 3
                                      End If
                                   End If
                                   If Check3.Value = 1 Then
                                      data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo") & " And fecha ='" & Format(data_cli.Recordset("fecha_baja"), "yyyy-mm-dd") & "'"
                                      data_abm.Refresh
                                      If data_abm.Recordset.RecordCount > 0 Then
                                         data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                      Else
                                         data_inf.Recordset("info_debit") = "SIN DATOS"
                                      End If
                                   End If
                                   data_inf.Recordset.Update
                                   data_cli.Recordset.MoveNext
                                Else
                                   data_cli.Recordset.MoveNext
                                End If
                             End If
                             If Wopsconvd = "AREAS P." Then
                                If data_convloc.Recordset("cnv_colrec") = "M" And _
                                   data_convloc.Recordset("cnv_cant_r") <> 2 And _
                                   data_cli.Recordset("cl_fecing") <= data_convloc.Recordset("cnv_desde") Then
                                   data_inf.Recordset.AddNew
                                   data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                   data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                   If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                      If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                         data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                      End If
                                   End If
                                   data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                   data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                                   data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                                   data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                                   data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                   data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                   data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                                   data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                                   data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                                   data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                                   data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                                   data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                                   data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                                   data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                                   data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                   data_inf.Recordset("cl_nomconv") = Mid(data_convloc.Recordset("cnv_desc"), 1, 30)
                                   data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                                   data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                                   data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                                   data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                   data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                                   If data_cli.Recordset("cl_sexo") = 2 Then
                                      data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                      data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                   Else
                                      If data_cli.Recordset("cl_sexo") = 1 Then
                                         data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                         data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                      Else
                                         data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                         data_inf.Recordset("cl_sexo") = 3
                                      End If
                                   End If
                                   If Check3.Value = 1 Then
                                      data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo") & " And fecha ='" & Format(data_cli.Recordset("fecha_baja"), "yyyy-mm-dd") & "'"
                                      data_abm.Refresh
                                      If data_abm.Recordset.RecordCount > 0 Then
                                         data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                      Else
                                         data_inf.Recordset("info_debit") = "SIN DATOS"
                                      End If
                                   End If
                                   data_inf.Recordset.Update
                                   data_cli.Recordset.MoveNext
                                Else
                                   data_cli.Recordset.MoveNext
                                End If
                             End If
                             If Wopsconvd = "PARCIAL" Then
                                If data_convloc.Recordset("cnv_colrec") = "A" And _
                                   data_convloc.Recordset("cnv_grupo") = "" Then
                                   data_inf.Recordset.AddNew
                                   data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                   data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                   If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                      If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                         data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                      End If
                                   End If
                                   data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                   data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                                   data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                                   data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                                   data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                   data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                   data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                                   data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                                   data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                                   data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                                   data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                                   data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                                   data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                                   data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                                   data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                   data_inf.Recordset("cl_nomconv") = Mid(data_convloc.Recordset("cnv_desc"), 1, 30)
                                   data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                                   data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                                   data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                                   data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                   data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                                   If data_cli.Recordset("cl_sexo") = 2 Then
                                      data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                      data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                   Else
                                      If data_cli.Recordset("cl_sexo") = 1 Then
                                         data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                         data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                      Else
                                         data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                         data_inf.Recordset("cl_sexo") = 3
                                      End If
                                   End If
                                   If Check3.Value = 1 Then
                                      data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo") & " And fecha ='" & Format(data_cli.Recordset("fecha_baja"), "yyyy-mm-dd") & "'"
                                      data_abm.Refresh
                                      If data_abm.Recordset.RecordCount > 0 Then
                                         data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                      Else
                                         data_inf.Recordset("info_debit") = "SIN DATOS"
                                      End If
                                   End If
                                   data_inf.Recordset.Update
                                   data_cli.Recordset.MoveNext
                                Else
                                   data_cli.Recordset.MoveNext
                                End If
                             End If
                          
                          Else
                             data_cli.Recordset.MoveNext
                          End If
                       Else
                          If Wopsconv = 4 Then 'Selección
'                             data_convloc.Recordset.FindFirst "cnv_codigo ='" & Wopsconvd & "'"
                             data_convloc.RecordSource = "Select * from convenio where cnv_codigo ='" & Wopsconvd & "'"
                             data_convloc.Refresh
                             If data_cli.Recordset("cl_codconv") = Wopsconvd Then
                                data_inf.Recordset.AddNew
                                data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                   If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                      data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                   End If
                                End If
                                data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                                data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                                data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                                data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                                data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                                data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                                data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                                data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                                data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                                data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                                data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                                data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                data_inf.Recordset("cl_nomconv") = Mid(data_convloc.Recordset("cnv_desc"), 1, 30)
                                data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                                data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                                data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                                data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                                If data_cli.Recordset("cl_sexo") = 2 Then
                                   data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                   data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                Else
                                   If data_cli.Recordset("cl_sexo") = 1 Then
                                      data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                      data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                   Else
                                      data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                      data_inf.Recordset("cl_sexo") = 3
                                   End If
                                End If
                                If Check3.Value = 1 Then
                                   data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo") & " And fecha ='" & Format(data_cli.Recordset("fecha_baja"), "yyyy-mm-dd") & "'"
                                   data_abm.Refresh
                                   If data_abm.Recordset.RecordCount > 0 Then
                                      data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                   Else
                                      data_inf.Recordset("info_debit") = "SIN DATOS"
                                   End If
                                End If
                                data_inf.Recordset.Update
                                data_cli.Recordset.MoveNext
                             Else
                                data_cli.Recordset.MoveNext
                             End If
                          Else
                             If Wopsconv = 9 Then 'Grupos de sapp sin complemento
'                                data_convloc.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                                data_convloc.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                                data_convloc.Refresh
                                If data_convloc.Recordset.RecordCount > 0 Then
                                   If Wopsconvd = "AMBULATORIO" Then
                                      If data_convloc.Recordset("cnv_colrec") = "R" And data_convloc.Recordset("cnv_grupo") = "" Then
                                         data_inf.Recordset.AddNew
                                         data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                         data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                         If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                            If data_cli.Recordset("cl_codced") >= 0 And data_cli.Recordset("cl_codced") <= 9 Then
                                               data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                                            End If
                                         End If
                                         data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                         data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                                         data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                                         data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                                         data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                         data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                         data_inf.Recordset("cl_edad") = data_cli.Recordset("cl_edad")
                                         data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                                         data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
                                         data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
                                         data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
                                         data_inf.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
                                         data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                                         data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
                                         data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                         data_inf.Recordset("cl_nomconv") = Mid(data_convloc.Recordset("cnv_desc"), 1, 30)
                                         data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                                         data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
                                         data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
                                         data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                         data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
                                         If data_cli.Recordset("cl_sexo") = 2 Then
                                            data_inf.Recordset("cl_diacobr") = "FEMENINO"
                                            data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                         Else
                                            If data_cli.Recordset("cl_sexo") = 1 Then
                                               data_inf.Recordset("cl_diacobr") = "MASCULINO"
                                               data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                            Else
                                               data_inf.Recordset("cl_diacobr") = "SIN DATO"
                                               data_inf.Recordset("cl_sexo") = 3
                                            End If
                                         End If
                                         If Check3.Value = 1 Then
                                            data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo") & " And fecha ='" & Format(data_cli.Recordset("fecha_baja"), "yyyy-mm-dd") & "'"
                                            data_abm.Refresh
                                            If data_abm.Recordset.RecordCount > 0 Then
                                               data_inf.Recordset("info_debit") = data_abm.Recordset("cl_motivo")
                                            Else
                                               data_inf.Recordset("info_debit") = "SIN DATOS"
                                            End If
                                         End If
                                         data_inf.Recordset.Update
                                         data_cli.Recordset.MoveNext
                                      Else
                                         data_cli.Recordset.MoveNext
                                      End If
                                   Else
                                      data_cli.Recordset.MoveNext
                                   End If
                                Else
                                   data_cli.Recordset.MoveNext
                                End If
                             Else
                                data_cli.Recordset.MoveNext
                             End If
                          End If
                       End If
                    End If
                 End If
              End If
           End If
        Else
           data_cli.Recordset.MoveNext
        End If
        barr.Value = barr.Value + 1
   Loop
End If
barr.Visible = False

End Sub

Private Sub Command7_Click()
Dim Xcuen As Long
Xcuen = 0

data_convloc.RecordSource = "convenio"
data_convloc.Refresh

If data_convloc.Recordset.RecordCount > 0 Then
   data_convloc.Recordset.MoveFirst
   Do While Not data_convloc.Recordset.EOF
      data_convloc.Recordset.Delete
      data_convloc.Recordset.MoveNext
   Loop
End If
data_inf.DatabaseName = App.path & "\infcli.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If

data_conv.Recordset.MoveFirst
Do While Not data_conv.Recordset.EOF
   If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
      If data_conv.Recordset("cnv_grupo") = Wopsconvd Then
         data_convloc.Recordset.AddNew
         data_convloc.Recordset("cnv_codigo") = data_conv.Recordset("cnv_codigo")
         data_convloc.Recordset("cnv_desc") = data_conv.Recordset("cnv_desc")
         data_convloc.Recordset("cnv_desde") = data_conv.Recordset("cnv_desde")
         data_convloc.Recordset("cnv_hasta") = data_conv.Recordset("cnv_hasta")
         data_convloc.Recordset("cnv_colrec") = data_conv.Recordset("cnv_colrec")
         data_convloc.Recordset("cnv_precio") = data_conv.Recordset("cnv_precio")
         data_convloc.Recordset("cnv_emite") = data_conv.Recordset("cnv_emite")
         data_convloc.Recordset("cnv_alta") = data_conv.Recordset("cnv_alta")
         data_convloc.Recordset("cnv_cant_r") = data_conv.Recordset("cnv_cant_r")
         data_convloc.Recordset("cnv_grupo") = data_conv.Recordset("cnv_grupo")
         data_convloc.Recordset.Update
      End If
   End If
   data_conv.Recordset.MoveNext
Loop
Wopsconvd = "H.EVANGELICO"
'data_convloc.RecordsetType = 0
'data_convloc.Recordset.Index = "cnv_codigo"
Dim Xldes, Xlhas As String
Xldes = "E"
Xlhas = "H"
data_cli.RecordSource = "select * from clientes where estado <>" & 2 & " And cl_codconv >='E' And cl_codconv <='H'"
data_cli.Refresh
data_cli.Recordset.MoveLast
barr.Max = data_cli.Recordset.RecordCount

data_cli.Recordset.MoveFirst
DoEvents
Do While Not data_cli.Recordset.EOF
   data_convloc.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
   data_cli.Recordset.MoveNext
   barr.Value = barr.Value + 1
Loop
MsgBox "Terminado " & Xcuen

End Sub

Private Sub Form_Load()
'data_cli.DatabaseName = App.Path & "\sapp.mdb"
data_cli.ConnectionString = "dsn=" & Xconexrmt
data_conv.ConnectionString = "dsn=" & Xconexrmt
'data_conv.RecordSource = "convenio"
'data_conv.Refresh
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.RecordSource = "infcli"
'data_inf.Refresh
data_abm.ConnectionString = "dsn=" & Xconexrmt
data_convloc.DatabaseName = App.path & "\inftab.mdb"

Wopsconv = 0
Wopsconvd = ""


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = 0
    .Width = 0
End With

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub md_LostFocus()
If md.Text <> "__/__/____" Then
   If IsDate(md.Text) = True Then
   Else
      MsgBox "Verifique la fecha"
   End If
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.ListIndex = 0
   Combo1.SetFocus
End If

End Sub

Private Sub mh_LostFocus()
If mh.Text <> "__/__/____" Then
   If IsDate(mh.Text) = True Then
   Else
      MsgBox "Verifique Fecha"
   End If
End If

End Sub


Public Sub Carga_mutuales()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from ca_adm order by ca_nom"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      Combo2.AddItem Xrecclii("ca_nom")
      Xrecclii.MoveNext
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub
