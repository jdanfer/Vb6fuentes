VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_vtasflia 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas por Familia"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6360
   Icon            =   "frm_vtasflia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3720
      TabIndex        =   21
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar barr 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin Crystal.CrystalReport crf 
      Left            =   1080
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton b_canc 
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
      MouseIcon       =   "frm_vtasflia.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "frm_vtasflia.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salida"
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton b_acep 
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
      MouseIcon       =   "frm_vtasflia.frx":0CD6
      MousePointer    =   99  'Custom
      Picture         =   "frm_vtasflia.frx":0FE0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Procesar"
      Top             =   5160
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Datos para informe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C00000&
         Caption         =   "Con edades"
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
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Solo Fertilab"
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
         Left            =   3000
         TabIndex        =   25
         Top             =   3840
         Width           =   2535
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   240
         TabIndex        =   20
         Top             =   2400
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1800
         Visible         =   0   'False
         Width           =   2535
      End
      Begin MSAdodcLib.Adodc data_conv 
         Height          =   375
         Left            =   3360
         Top             =   3360
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSAdodcLib.Adodc data_cli 
         Height          =   375
         Left            =   3960
         Top             =   3720
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   375
         Left            =   1440
         Top             =   3240
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
         Caption         =   "data_lin"
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
      Begin VB.CheckBox Check2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Desde historial"
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
         TabIndex        =   19
         Top             =   3840
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Top             =   1560
         Width           =   1575
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
         ItemData        =   "frm_vtasflia.frx":156A
         Left            =   2640
         List            =   "frm_vtasflia.frx":1580
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   14
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Informe sin detalle"
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
         Left            =   2760
         TabIndex        =   11
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox tm 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Data data_med 
         Caption         =   "data_med"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "familias"
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txt_b 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   7
         Top             =   2160
         Width           =   735
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "frm_vtasflia.frx":15C0
         Height          =   360
         Left            =   1560
         TabIndex        =   5
         Top             =   960
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648384
         ForeColor       =   0
         ListField       =   "FAM_NOMBRE"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   2040
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Código de servicio:"
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
         TabIndex        =   17
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "GRUPO MUTUAL:"
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
         TabIndex        =   15
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "CÓDIGO DE CONVENIO:"
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
         TabIndex        =   13
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "BASE: (99=TODAS)"
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
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Familia"
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
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
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
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Label labd 
      Height          =   255
      Left            =   1800
      TabIndex        =   24
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label labm 
      Height          =   255
      Left            =   960
      TabIndex        =   23
      Top             =   5520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label laba 
      Height          =   255
      Left            =   1080
      TabIndex        =   22
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   2280
      Picture         =   "frm_vtasflia.frx":15D7
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1815
   End
End
Attribute VB_Name = "frm_vtasflia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_acep_Click()
b_acep.Enabled = False
b_canc.Enabled = False
Dim Xcualedad As Long
Dim Xlaedades As String
Dim Siono As Integer
Dim DifAnio As Double

Dim xcuenta, Xcodprod As Long
xcuenta = 0
DifAnio = 0

Siono = 0
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"
data_inf.RecordSource = "infvtas"
data_inf.Refresh

Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet

Dim XCol, Xlin, Xnrocan, Xcolfija, Xcantsrv, Xcanttot As Long
Dim Xarchtex As String
Dim Xlabrir As New Excel.Application

'Dim Xlin As String
'      Open "c:\debitos\PruebaC.csv" For Output As #1
'      Xlin = "1000" & vbTab & "1020010240001" & vbTab & "JUAN PEREZ" & vbTab & "085801009005"
'      Print #1, Xlin
'      Xlin = "1000" & vbTab & "1020010240001" & vbTab & "JORGE DANIEL FERNANDEZ SOSA" & vbTab & "085801009005"
'      Print #1, Xlin
'      Close #1

XCol = 1
Xlin = 1
Xnrocan = 1
List1.Clear
Xcanttot = 0
Xcantsrv = 0
If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      If tm.Text <> "" Then
         If txt_b.Text <> "" Then
            If txt_b.Text = 99 Then
               If Text1.Text = "" Then
                  If Text2.Text = "" Then
                     If Check2.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " and tipo <>'" & "NOTA CR" & "' order by cod_prod"
                        data_lin.Refresh
                     Else
                        If tm.Text = 1 Then
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod in (10001,10002,10003,10004,10005,10006,10007,10008) order by cod_prod"
                           data_lin.Refresh
                        Else
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " order by cod_prod"
                           data_lin.Refresh
                        End If
                     End If
                  Else
                     If Check2.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " and cod_prod =" & Text2.Text & " and tipo <>'" & "NOTA CR" & "' order by cod_prod"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " and cod_prod =" & Text2.Text & " order by cod_prod"
                        data_lin.Refresh
                     End If
                  End If
               Else
                  
                  If Check2.Value = 1 Then
                     data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " and convenio ='" & Text1.Text & "' and tipo <>'" & "NOTA CR" & "' order by cod_prod"
                     data_lin.Refresh
                  Else
                     data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " and convenio ='" & Text1.Text & "' order by cod_prod"
                     data_lin.Refresh
                  End If
               End If
            Else
               If Text1.Text = "" Then
                  If Text2.Text = "" Then
                     If Check2.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " And base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "' order by cod_prod"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " And base =" & txt_b.Text & " order by cod_prod"
                        data_lin.Refresh
                     End If
                  Else
                     If Check2.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " And base =" & txt_b.Text & " and cod_prod =" & Text2.Text & " and tipo <>'" & "NOTA CR" & "' order by cod_prod"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " And base =" & txt_b.Text & " and cod_prod =" & Text2.Text & " order by cod_prod"
                        data_lin.Refresh
                     End If
                  End If
               Else
                  If Check2.Value = 1 Then
                     data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " And base =" & txt_b.Text & " And convenio ='" & Text1.Text & "' and tipo <>'" & "NOTA CR" & "' order by cod_prod"
                     data_lin.Refresh
                  Else
                     data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " And base =" & txt_b.Text & " And convenio ='" & Text1.Text & " order by cod_prod"
                     data_lin.Refresh
                  End If
               End If
            End If
            If Check3.Value = 1 Then
               data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia in (3) and esfertilab in (1) order by cod_prod"
               data_lin.Refresh
            End If
            If data_lin.Recordset.RecordCount > 0 Then
               Set Xobjexel = New Excel.Application
               Set Xlibexel = Xobjexel.Workbooks.Add
               Set Xarchexel = Xlibexel.Worksheets.Add
               Xarchexel.Name = "VENTAS POR FAMILIA"
               Xlibexel.SaveAs ("C:\planillas\" & "Infvtas" & ".xls")
               Xarchtex = "C:\planillas\" & "Infvtas" & ".xls"
               frm_vtasflia.MousePointer = 11
               barr.Visible = True
               data_lin.Recordset.MoveLast
               barr.Max = data_lin.Recordset.RecordCount + 100
               barr.Value = 0
               data_lin.Recordset.MoveFirst
               DoEvents
                Xarchexel.Cells(Xlin, XCol) = "SAPP - CÓMPUTOS"
                Xlin = Xlin + 1
                XCol = XCol + 1
                Xarchexel.Range("A1", "C3").Font.Size = 16
                Xarchexel.Cells(Xlin, XCol) = "INFORME VENTAS POR FAMILIA DESDE: " & md.Text & " HASTA: " & mh.Text
                Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
                XCol = 1
                Xlin = Xlin + 2
                Xnrocan = Xnrocan + Xlin
                Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
                Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
                Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
                Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
                Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
                Xarchexel.Range("A" & Trim(str(Xlin)), "AD" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
                Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
                Xarchexel.Cells(Xlin, XCol) = "FECHA"
                XCol = XCol + 1
                Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
                XCol = XCol + 1
                Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
                Xarchexel.Cells(Xlin, XCol) = "NOMBRE"
                XCol = XCol + 1
                Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "CEDULA"
                XCol = XCol + 1
                Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "COD.PROD"
                XCol = XCol + 1
                Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 35
                Xarchexel.Cells(Xlin, XCol) = "SERVICIO"
                XCol = XCol + 1
                Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
                XCol = XCol + 1
                Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "CUOTA"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "FLIA."
                XCol = XCol + 1
                Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "COSTO"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "BASE"
                XCol = XCol + 1
                Xarchexel.Range("L" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "HORA"
                XCol = XCol + 1
                Xarchexel.Range("M" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "OPERADOR"
                XCol = XCol + 1
                Xarchexel.Range("N" & Trim(str(Xlin))).ColumnWidth = 9
                Xarchexel.Cells(Xlin, XCol) = "EDAD"
                
                Dim XimpCuota As Double
                XimpCuota = 0
                Xlin = Xlin + 1
                XCol = 1
                Xcodprod = data_lin.Recordset("cod_prod")
               Do While Not data_lin.Recordset.EOF
                   If tm.Text = 19 Then
                      If Combo1.ListIndex > 0 Then
                         data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
                         data_conv.Refresh
                         If data_conv.Recordset.RecordCount > 0 Then
                            If IsNull(data_conv.Recordset("cnv_grupo")) = True Then
                               Siono = 9
                            Else
                               If Combo1.Text = data_conv.Recordset("cnv_grupo") Then
                                  Siono = 0
                               Else
                                  Siono = 9
                               End If
                            End If
                         End If
                      Else
                         Siono = 0
                      End If
                   Else
                      If Combo1.ListIndex > 0 Then
                         data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
                         data_conv.Refresh
                         If data_conv.Recordset.RecordCount > 0 Then
                            XimpCuota = data_conv.Recordset("cnv_precio")
                            If IsNull(data_conv.Recordset("cnv_grupo")) = True Then
                               Siono = 9
                            Else
                               If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                                  Siono = 0
                               Else
                                  Siono = 9
                               End If
                            End If
                         End If
                      Else
                         Siono = 0
                      End If
                   End If
                   If Siono = 9 Then
                   Else
                      If IsNull(data_lin.Recordset("cod_prod")) = False Then
                        If Xcodprod = data_lin.Recordset("cod_prod") Then
                           Xcantsrv = Xcantsrv + 1
                           Xcanttot = Xcanttot + 1
                        Else
                           data_lin.Recordset.MovePrevious
                           List1.AddItem data_lin.Recordset("nom_prod") & ": " & Xcantsrv
                           data_lin.Recordset.MoveNext
                           Xcantsrv = 1
                           Xcanttot = Xcanttot + 1
                        End If
                     End If
                      Xarchexel.Cells(Xlin, XCol) = "'" & Format(data_lin.Recordset("fecha"), "dd/mm/yyyy")
                      XCol = XCol + 1
                      If IsNull(data_lin.Recordset("cod_cli")) = False Then
                         Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("cod_cli")
                      End If
                      XCol = XCol + 1
                      If IsNull(data_lin.Recordset("nom_cli")) = False Then
                         Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("nom_cli")
                      End If
                      XCol = XCol + 1
                      If IsNull(data_lin.Recordset("ced_socio")) = False Then
                         If IsNull(data_lin.Recordset("fact")) = False Then
                            Xarchexel.Cells(Xlin, XCol) = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                         End If
                      End If
                      XCol = XCol + 1
                      If IsNull(data_lin.Recordset("cod_prod")) = False Then
                         Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("cod_prod")
                      End If
                      XCol = XCol + 1
                      If IsNull(data_lin.Recordset("nom_prod")) = False Then
                         Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("nom_prod")
                      End If
                      XCol = XCol + 1
                      If IsNull(data_lin.Recordset("convenio")) = False Then
                         Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("convenio")
                      End If
                      XCol = XCol + 1
                      Xarchexel.Cells(Xlin, XCol) = XimpCuota
                      
                      XCol = XCol + 1
                      If IsNull(data_lin.Recordset("nro_flia")) = False Then
                         Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("nro_flia")
                      End If
                      XCol = XCol + 1
                      If IsNull(data_lin.Recordset("tot_lin")) = False Then
                         Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("tot_lin")
                      End If
                      XCol = XCol + 1
                      If IsNull(data_lin.Recordset("base")) = False Then
                         Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("base")
                      End If
                      XCol = XCol + 1
                      If IsNull(data_lin.Recordset("hora")) = False Then
                         Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("hora")
                      End If
                      XCol = XCol + 1
                      If IsNull(data_lin.Recordset("operador")) = False Then
                         Xarchexel.Cells(Xlin, XCol) = data_lin.Recordset("operador")
                      End If
                      If Check4.Value = 1 Then
                         data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                         data_cli.Refresh
                         If data_cli.Recordset.RecordCount > 0 Then
                            If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                               DifAnio = Int(DateDiff("y", data_cli.Recordset("cl_fnac"), Date) / 365)
                               DifAnio = Int(DifAnio)
                            Else
                               DifAnio = 0
                            End If
                         Else
                            DifAnio = 0
                         End If
                      Else
                         DifAnio = 0
                      End If
                      XCol = XCol + 1
                      Xarchexel.Cells(Xlin, XCol) = Int(DifAnio)
                      
                      Xlin = Xlin + 1
                      XCol = 1
                      
                      'data_inf.Recordset.Update
                  End If
                  If IsNull(data_lin.Recordset("cod_prod")) = False Then
                     Xcodprod = data_lin.Recordset("cod_prod")
                  End If
                  data_lin.Recordset.MoveNext
                  barr.Value = barr.Value + 1
               Loop
               DoEvents
               data_lin.Recordset.MovePrevious
               If IsNull(data_lin.Recordset("nom_prod")) = False Then
                  List1.AddItem data_lin.Recordset("nom_prod") & ": " & Xcantsrv
               End If
               data_lin.Recordset.MoveNext
               Xcantsrv = 1
               Xcanttot = Xcanttot '+1 ??
               
               Xlin = Xlin + 1
               XCol = 2
               Dim i As Integer
               i = 0
               Xarchexel.Cells(Xlin, XCol) = "TOTALES POR SERVICIO: "
               If Text2.Text = "" Then
                  List1.ListIndex = 0
                  Do While i <= List1.ListCount - 1
                     List1.ListIndex = i
                     Xarchexel.Cells(Xlin, XCol) = List1.List(List1.ListIndex)
                     i = i + 1
                     Xlin = Xlin + 1
                  Loop
               End If
               Xlin = Xlin + 1
               Xarchexel.Cells(Xlin, XCol) = "TOTAL GENERAL: " & Xcanttot
                
                Xlibexel.Save
                Xlibexel.Close
                Xobjexel.Quit
                Xlabrir.Workbooks.Open Xarchtex, , False
                Xlabrir.Visible = True
                Xlabrir.WindowState = xlMaximized
               
               data_inf.RecordSource = "Select * from infvtas order by nro_flia"
               data_inf.Refresh
               If Text1.Text = "SMIN" Then
                  If data_inf.Recordset.RecordCount > 0 Then
                     data_inf.Recordset.MoveLast
                     barr.Max = barr.Max + data_inf.Recordset.RecordCount - 100
                     data_inf.Recordset.MoveFirst
                     Do While Not data_inf.Recordset.EOF
                        data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_inf.Recordset("cod_cli")
                        data_cli.Refresh
                        If data_cli.Recordset.RecordCount > 0 Then
                           If IsNull(data_cli.Recordset("cl_cedula")) = False Then
                              If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                 data_inf.Recordset.Edit
                                 data_inf.Recordset("zona") = Trim(str(data_cli.Recordset("cl_cedula"))) & "-" & Trim(str(data_cli.Recordset("cl_codced")))
                                 data_inf.Recordset.Update
                              Else
                                 data_inf.Recordset.Edit
                                 data_inf.Recordset("zona") = Trim(str(data_cli.Recordset("cl_cedula"))) & "-0"
                                 data_inf.Recordset.Update
                              End If
                           Else
                              data_inf.Recordset.Edit
                              data_inf.Recordset("zona") = "0"
                              data_inf.Recordset.Update
                           End If
                           data_inf.Recordset.Edit
                           If IsNull(data_cli.Recordset("cl_dpto")) = False Then
                              data_inf.Recordset("nom_superv") = data_cli.Recordset("cl_dpto")
                           Else
                              If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                                 data_inf.Recordset("nom_superv") = data_cli.Recordset("cl_telefon")
                              Else
                                 data_inf.Recordset("nom_superv") = "0"
                              End If
                           End If
                           data_inf.Recordset("nom_medic") = Mid(data_cli.Recordset("cl_direcci"), 1, 50)
                           data_inf.Recordset("nom_med_s") = Mid(data_cli.Recordset("cl_zona"), 1, 40)
                           data_inf.Recordset.Update
                        End If
                        data_inf.Recordset.MoveNext
                     Loop
                  End If
               End If
               barr.Value = 0
               barr.Visible = False
               frm_vtasflia.MousePointer = 0
               data_inf.RecordSource = "Select * from infvtas"
               data_inf.Refresh
               
'               If Check1.value = 1 Then
'                  If tm.Text = 19 Then
'                     crf.ReportFileName = App.Path & "\infvta19d.rpt"
'                     If txt_b.Text = 99 Then
'                        crf.ReportTitle = "INFORME DE METAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " == TODAS LAS BASES =="
'                     Else
'                        crf.ReportTitle = "INFORME DE METAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " == BASE: " & txt_b.Text
'                     End If
'                     crf.Action = 1
'
'                  Else
'                     crf.ReportFileName = App.Path & "\infvtasxflin.rpt"
'                     If txt_b.Text = 99 Then
'                        crf.ReportTitle = "INFORME DE VENTAS POR FAMILIA FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " == TODAS LAS BASES =="
'                     Else
'                        crf.ReportTitle = "INFORME DE VENTAS POR FAMILIA FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " == BASE: " & txt_b.Text
'                     End If
'                     crf.Action = 1
'                  End If
'               Else
'                  If tm.Text = 19 Then
'                     crf.ReportFileName = App.Path & "\infvta19d.rpt"
'                     If txt_b.Text = 99 Then
'                        crf.ReportTitle = "INFORME DE METAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " == TODAS LAS BASES =="
'                     Else
'                        crf.ReportTitle = "INFORME DE METAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " == BASE: " & txt_b.Text
'                     End If
'                     crf.Action = 1
'                  Else
'                     crf.ReportFileName = App.Path & "\infvtasxfli.rpt"
'                     If txt_b.Text = 99 Then
'                        crf.ReportTitle = "INFORME DE VENTAS POR FAMILIA FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " == TODAS LAS BASES =="
'                     Else
'                        crf.ReportTitle = "INFORME DE VENTAS POR FAMILIA FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " == BASE: " & txt_b.Text
'                     End If
'                     crf.Action = 1
'                  End If
'               End If
            
            Else
               MsgBox "No existen registros con esta selección", vbInformation, "Mensaje"
            End If
         Else
            MsgBox "Ingrese Base", vbInformation, "Mensaje"
            txt_b.SetFocus
         End If
      Else
         MsgBox "Número de familia incorrecto", vbInformation, "Mensaje"
         DBCombo1.SetFocus
      End If
   Else
      MsgBox "Ingrese Fecha", vbInformation, "Mensaje"
      mh.SetFocus
   End If
Else
   MsgBox "Ingrese fecha", vbInformation, "Mensaje"
   md.SetFocus
End If
b_acep.Enabled = True
b_canc.Enabled = True

End Sub

Private Sub b_canc_Click()
Unload Me

End Sub


Private Sub Command1_Click()
Dim Xcualedad As Long
Dim Xlaedades As String
Dim Siono As Integer
Dim xcuenta, Xcodprod As Long
xcuenta = 0
Siono = 0
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"
data_inf.RecordSource = "infvtas"
data_inf.Refresh

Dim XCol, Xnrocan, Xcolfija, Xcantsrv, Xcanttot As Long
Dim Xarchtex As String

Dim Xlin As String

XCol = 1
Xlin = 1
Xnrocan = 1
List1.Clear
Xcanttot = 0
Xcantsrv = 0
If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      If tm.Text <> "" Then
         If txt_b.Text <> "" Then
            If txt_b.Text = 99 Then
               If Text1.Text = "" Then
                  If Text2.Text = "" Then
                     If Check2.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " and tipo <>'" & "NOTA CR" & "' order by cod_prod"
                        data_lin.Refresh
                     Else
                        If tm.Text = 1 Then
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod in (10001,10002,10003,10004,10005,10006,10007,10008) order by cod_prod"
                           data_lin.Refresh
                        Else
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " order by cod_prod"
                           data_lin.Refresh
                        End If
                     End If
                  Else
                     If Check2.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " and cod_prod =" & Text2.Text & " and tipo <>'" & "NOTA CR" & "' order by cod_prod"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " and cod_prod =" & Text2.Text & " order by cod_prod"
                        data_lin.Refresh
                     End If
                  End If
               Else
                  
                  If Check2.Value = 1 Then
                     data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " and convenio ='" & Text1.Text & "' and tipo <>'" & "NOTA CR" & "' order by cod_prod"
                     data_lin.Refresh
                  Else
                     data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " and convenio ='" & Text1.Text & "' order by cod_prod"
                     data_lin.Refresh
                  End If
               End If
            Else
               If Text1.Text = "" Then
                  If Text2.Text = "" Then
                     If Check2.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " And base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "' order by cod_prod"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " And base =" & txt_b.Text & " order by cod_prod"
                        data_lin.Refresh
                     End If
                  Else
                     If Check2.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " And base =" & txt_b.Text & " and cod_prod =" & Text2.Text & " and tipo <>'" & "NOTA CR" & "' order by cod_prod"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " And base =" & txt_b.Text & " and cod_prod =" & Text2.Text & " order by cod_prod"
                        data_lin.Refresh
                     End If
                  End If
               Else
                  If Check2.Value = 1 Then
                     data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " And base =" & txt_b.Text & " And convenio ='" & Text1.Text & "' and tipo <>'" & "NOTA CR" & "' order by cod_prod"
                     data_lin.Refresh
                  Else
                     data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_flia =" & tm.Text & " And base =" & txt_b.Text & " And convenio ='" & Text1.Text & " order by cod_prod"
                     data_lin.Refresh
                  End If
               End If
            End If
            If data_lin.Recordset.RecordCount > 0 Then
               Open "C:\planillas\" & "Infvtas" & ".csv" For Output As #1
'      Print #1, Xlin
'      Xlin = "1000" & vbTab & "1020010240001" & vbTab & "JORGE DANIEL FERNANDEZ SOSA" & vbTab & "085801009005"
'      Print #1, Xlin
'      Close #1
               
               barr.Visible = True
               data_lin.Recordset.MoveLast
               barr.Max = data_lin.Recordset.RecordCount + 100
               barr.Value = 0
               data_lin.Recordset.MoveFirst
               DoEvents
               Xlin = "SAPP - CÓMPUTOS" & vbTab & "FECHA:" & vbTab & Format(Date, "dd/mm/yyyy")
               Print #1, Xlin
               Xlin = "INFORME VENTAS POR FAMILIA DESDE: " & md.Text & " HASTA: " & mh.Text
               Print #1, Xlin
               Xlin = "FECHA" & vbTab & "MATRICULA" & vbTab & "NOMBRE" & vbTab & "CEDULA" & vbTab & "COD.PROD" & _
               vbTab & "SERVICIO" & vbTab & "CONVENIO" & vbTab & "FLIA." & vbTab & "COSTO" & vbTab & "BASE"
               Print #1, Xlin
               Xcodprod = data_lin.Recordset("cod_prod")
               Do While Not data_lin.Recordset.EOF
                   If tm.Text = 19 Then
                      If Combo1.ListIndex > 0 Then
                         data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
                         data_conv.Refresh
                         If data_conv.Recordset.RecordCount > 0 Then
                            If IsNull(data_conv.Recordset("cnv_grupo")) = True Then
                               Siono = 9
                            Else
                               If Combo1.Text = data_conv.Recordset("cnv_grupo") Then
                                  Siono = 0
                               Else
                                  Siono = 9
                               End If
                            End If
                         End If
                      Else
                         Siono = 0
                      End If
                   Else
                      If Combo1.ListIndex > 0 Then
                         data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
                         data_conv.Refresh
                         If data_conv.Recordset.RecordCount > 0 Then
                            If IsNull(data_conv.Recordset("cnv_grupo")) = True Then
                               Siono = 9
                            Else
                               If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                                  Siono = 0
                               Else
                                  Siono = 9
                               End If
                            End If
                         End If
                      Else
                         Siono = 0
                      End If
                   End If
                   If Siono = 9 Then
                   Else
                      If IsNull(data_lin.Recordset("cod_prod")) = False Then
                        If Xcodprod = data_lin.Recordset("cod_prod") Then
                           Xcantsrv = Xcantsrv + 1
                           Xcanttot = Xcanttot + 1
                        Else
                           data_lin.Recordset.MovePrevious
                           List1.AddItem data_lin.Recordset("nom_prod") & ": " & Xcantsrv
                           data_lin.Recordset.MoveNext
                           Xcantsrv = 1
                           Xcanttot = Xcanttot + 1
                        End If
                     End If
                     Xlin = Format(data_lin.Recordset("fecha"), "dd/mm/yyyy") & vbTab
                     If IsNull(data_lin.Recordset("cod_cli")) = False Then
                        Xlin = Xlin & data_lin.Recordset("cod_cli") & vbTab
                     End If
                     If IsNull(data_lin.Recordset("nom_cli")) = False Then
                        Xlin = Xlin & data_lin.Recordset("nom_cli") & vbTab
                     End If
                     If IsNull(data_lin.Recordset("ced_socio")) = False Then
                        If IsNull(data_lin.Recordset("fact")) = False Then
                           Xlin = Xlin & Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact"))) & vbTab
                        End If
                     End If
                     If IsNull(data_lin.Recordset("cod_prod")) = False Then
                        Xlin = Xlin & data_lin.Recordset("cod_prod") & vbTab
                     End If
                     If IsNull(data_lin.Recordset("nom_prod")) = False Then
                        Xlin = Xlin & data_lin.Recordset("nom_prod") & vbTab
                     End If
                     If IsNull(data_lin.Recordset("convenio")) = False Then
                        Xlin = Xlin & data_lin.Recordset("convenio") & vbTab
                     End If
                     If IsNull(data_lin.Recordset("nro_flia")) = False Then
                        Xlin = Xlin & data_lin.Recordset("nro_flia") & vbTab
                     End If
                     If IsNull(data_lin.Recordset("tot_lin")) = False Then
                        Xlin = Xlin & data_lin.Recordset("tot_lin") & vbTab
                     End If
                     If IsNull(data_lin.Recordset("base")) = False Then
                        Xlin = Xlin & data_lin.Recordset("base") & vbTab
                     End If
                      
                      If tm.Text = 199 Then
                         data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                         data_cli.Refresh
                         If data_cli.Recordset.RecordCount > 0 Then
                            data_inf.Recordset.AddNew
                            data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                            data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                            data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                            
                            If IsNull(data_cli.Recordset("cl_sexo")) = True Then
                               data_inf.Recordset("operador") = "MASC"
                            Else
                               If data_cli.Recordset("cl_sexo") = 2 Then
                                  data_inf.Recordset("operador") = "FEM"
                               Else
                                  data_inf.Recordset("operador") = "MASC"
                               End If
                            End If
                            If IsNull(data_cli.Recordset("cl_fnac")) = True Then
                               If IsNull(data_cli.Recordset("cl_fultpag")) = True Then
                                  data_inf.Recordset("cantidad") = 999
                                  data_inf.Recordset("nro_superv") = 999
                               Else
                                  Xcualedad = data_inf.Recordset("fecha") - data_cli.Recordset("cl_fultpag")
                                  If Xcualedad < 365 Then
                                     If Xcualedad < 31 Then
                                        Xlaedades = "D"
                                     Else
                                        Xcualedad = Xcualedad / 30
                                        Xlaedades = "M"
                                     End If
                                  Else
                                     Xcualedad = Xcualedad / 365
                                     Xlaedades = "A"
                                  End If
                                  data_inf.Recordset("realizada") = data_cli.Recordset("cl_fultpag")
                                  data_inf.Recordset("cantidad") = Int(Xcualedad)
                                  data_inf.Recordset("tipo_mov") = Xlaedades
        ' Edad actual
                                  Xcualedad = Date - data_cli.Recordset("cl_fultpag")
                                  If Xcualedad < 365 Then
                                     If Xcualedad < 31 Then
                                        Xlaedades = "D"
                                     Else
                                        Xcualedad = Xcualedad / 30
                                        Xlaedades = "M"
                                     End If
                                  Else
                                     Xcualedad = Xcualedad / 365
                                     Xlaedades = "A"
                                  End If
                                  data_inf.Recordset("nro_superv") = Int(Xcualedad)
                                  data_inf.Recordset("moneda") = Xlaedades
                               End If
                            Else
                               Xcualedad = data_inf.Recordset("fecha") - data_cli.Recordset("cl_fnac")
                               If Xcualedad < 365 Then
                                  If Xcualedad < 31 Then
                                     Xlaedades = "D"
                                  Else
                                     Xcualedad = Xcualedad / 30
                                     Xlaedades = "M"
                                  End If
                               Else
                                  Xcualedad = Xcualedad / 365
                                  Xlaedades = "A"
                               End If
                               data_inf.Recordset("realizada") = data_cli.Recordset("cl_fnac")
                               data_inf.Recordset("cantidad") = Int(Xcualedad)
                               data_inf.Recordset("tipo_mov") = Xlaedades
        ' Edad actual
                               Xcualedad = Date - data_cli.Recordset("cl_fnac")
                               If Xcualedad < 365 Then
                                  If Xcualedad < 31 Then
                                     Xlaedades = "D"
                                  Else
                                     Xcualedad = Xcualedad / 30
                                     Xlaedades = "M"
                                  End If
                               Else
                                  Xcualedad = Xcualedad / 365
                                  Xlaedades = "A"
                               End If
                               data_inf.Recordset("nro_superv") = Int(Xcualedad)
                               data_inf.Recordset("moneda") = Xlaedades
                            End If
                         End If
                      End If
                      'data_inf.Recordset.Update
                  End If
                  If IsNull(data_lin.Recordset("cod_prod")) = False Then
                     Xcodprod = data_lin.Recordset("cod_prod")
                  End If
                  Print #1, Xlin
                  data_lin.Recordset.MoveNext
                  barr.Value = barr.Value + 1
               Loop
               DoEvents
               data_lin.Recordset.MovePrevious
               If IsNull(data_lin.Recordset("nom_prod")) = False Then
                  List1.AddItem data_lin.Recordset("nom_prod") & ": " & Xcantsrv
               End If
               data_lin.Recordset.MoveNext
               Xcantsrv = 1
               Xcanttot = Xcanttot + 1
               
               Dim i As Integer
               i = 0
               Xlin = "==============================================================================="
               Print #1, Xlin
               Xlin = "TOTALES POR SERVICIO: "
               Print #1, Xlin
               If Text2.Text = "" Then
                  List1.ListIndex = 1
                  Do While i <= List1.ListCount - 1
                     List1.ListIndex = i
                     Xlin = List1.List(List1.ListIndex)
                     i = i + 1
                     Print #1, Xlin
                  Loop
               End If
               Xlin = "TOTAL GENERAL: " & vbTab & Xcanttot
               Print #1, Xlin
               Close #1
               data_inf.RecordSource = "Select * from infvtas order by nro_flia"
               data_inf.Refresh
               If Text1.Text = "SMIN" Then
                  If data_inf.Recordset.RecordCount > 0 Then
                     data_inf.Recordset.MoveLast
                     barr.Max = barr.Max + data_inf.Recordset.RecordCount - 100
                     data_inf.Recordset.MoveFirst
                     Do While Not data_inf.Recordset.EOF
                        data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_inf.Recordset("cod_cli")
                        data_cli.Refresh
                        If data_cli.Recordset.RecordCount > 0 Then
                           If IsNull(data_cli.Recordset("cl_cedula")) = False Then
                              If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                 data_inf.Recordset.Edit
                                 data_inf.Recordset("zona") = Trim(str(data_cli.Recordset("cl_cedula"))) & "-" & Trim(str(data_cli.Recordset("cl_codced")))
                                 data_inf.Recordset.Update
                              Else
                                 data_inf.Recordset.Edit
                                 data_inf.Recordset("zona") = Trim(str(data_cli.Recordset("cl_cedula"))) & "-0"
                                 data_inf.Recordset.Update
                              End If
                           Else
                              data_inf.Recordset.Edit
                              data_inf.Recordset("zona") = "0"
                              data_inf.Recordset.Update
                           End If
                           data_inf.Recordset.Edit
                           If IsNull(data_cli.Recordset("cl_dpto")) = False Then
                              data_inf.Recordset("nom_superv") = data_cli.Recordset("cl_dpto")
                           Else
                              If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                                 data_inf.Recordset("nom_superv") = data_cli.Recordset("cl_telefon")
                              Else
                                 data_inf.Recordset("nom_superv") = "0"
                              End If
                           End If
                           data_inf.Recordset("nom_medic") = Mid(data_cli.Recordset("cl_direcci"), 1, 50)
                           data_inf.Recordset("nom_med_s") = Mid(data_cli.Recordset("cl_zona"), 1, 40)
                           data_inf.Recordset.Update
                        End If
                        data_inf.Recordset.MoveNext
                     Loop
                  End If
               End If
               barr.Value = 0
               barr.Visible = False
               frm_vtasflia.MousePointer = 0
               data_inf.RecordSource = "Select * from infvtas"
               data_inf.Refresh
               
            
            Else
               MsgBox "No existen registros con esta selección", vbInformation, "Mensaje"
            End If
         Else
            MsgBox "Ingrese Base", vbInformation, "Mensaje"
            txt_b.SetFocus
         End If
      Else
         MsgBox "Número de familia incorrecto", vbInformation, "Mensaje"
         DBCombo1.SetFocus
      End If
   Else
      MsgBox "Ingrese Fecha", vbInformation, "Mensaje"
      mh.SetFocus
   End If
Else
   MsgBox "Ingrese fecha", vbInformation, "Mensaje"
   md.SetFocus
End If
b_acep.Enabled = True
b_canc.Enabled = True

End Sub

Private Sub DBCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_b.SetFocus
End If

End Sub

Private Sub DBCombo1_LostFocus()

If IsNumeric(DBCombo1.Text) = True Then
   data_med.Recordset.FindFirst "fam_numero =" & DBCombo1.Text
   If Not data_med.Recordset.NoMatch Then
      tm.Text = data_med.Recordset("fam_numero")
      DBCombo1.Text = data_med.Recordset("fam_nombre")
   Else
      MsgBox "No encontrado, consulte por nombre", vbInformation, "Mensaje"
      DBCombo1.SetFocus
   End If
Else
   If DBCombo1.Text <> "" Then
      data_med.Recordset.FindFirst "fam_nombre ='" & DBCombo1.Text & "'"
      If Not data_med.Recordset.NoMatch Then
         tm.Text = data_med.Recordset("fam_numero")
         DBCombo1.Text = data_med.Recordset("fam_nombre")
      Else
         MsgBox "No encontrado, consulte por nombre", vbInformation, "Mensaje"
         DBCombo1.SetFocus
      End If
   End If
End If


End Sub

Private Sub Form_Load()
data_med.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_lin.ConnectionString = "dsn=" & Xconexrmt
data_cli.ConnectionString = "dsn=" & Xconexrmt
data_conv.ConnectionString = "dsn=" & Xconexrmt

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
   DBCombo1.SetFocus
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_b.SetFocus
End If

End Sub

Private Sub txt_b_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_acep.SetFocus
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
Anios = DateDiff("yyyy", CDate(Format(FNaci, "dd/MM/yyyy")), CDate(FAct))
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
   labedad.Caption = Anios
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
   labunie.Caption = Meses
   labdias.Caption = Dias
Else
   MsgBox "Fecha Inválida"
   labedad.Caption = 0
   labunie.Caption = 0
   labdias.Caption = 0
End If

End Sub


