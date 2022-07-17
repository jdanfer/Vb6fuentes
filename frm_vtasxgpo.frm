VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_vtasxgpo 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas por grupo"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "frm_vtasxgpo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   5280
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Data data_buscnv 
      Caption         =   "data_buscnv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   2895
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   6240
      Top             =   5400
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
      Height          =   495
      Left            =   5640
      Picture         =   "frm_vtasxgpo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   5760
      Width           =   615
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
      Height          =   495
      Left            =   360
      Picture         =   "frm_vtasxgpo.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Procesar"
      Top             =   5760
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Opciones de listado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.CommandButton Command6 
         Caption         =   "Evang"
         Height          =   375
         Left            =   480
         TabIndex        =   24
         Top             =   3240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         Caption         =   "smi"
         Height          =   375
         Left            =   1440
         TabIndex        =   23
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Crear informe mutual"
         Height          =   495
         Left            =   3840
         TabIndex        =   22
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Data data_infccou 
         Caption         =   "data_infccou"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ccou"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1440
         Visible         =   0   'False
         Width           =   2775
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   3600
         Top             =   3240
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
         Caption         =   "Adodc1"
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
         Left            =   2640
         Top             =   3840
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
      Begin MSAdodcLib.Adodc data_conv 
         Height          =   375
         Left            =   600
         Top             =   2640
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
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   375
         Left            =   3480
         Top             =   3120
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
      Begin VB.Data data_inf2 
         Caption         =   "data_inf2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FF0000&
         Caption         =   "Cantidad de consultas por cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0000FFFF&
         Caption         =   "Informe por rango de edad"
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
         Left            =   120
         TabIndex        =   17
         Top             =   4560
         Width           =   5655
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Informes desde respaldos"
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
         Left            =   120
         TabIndex        =   16
         Top             =   4080
         Width           =   5655
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF0000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   3600
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ver Servicios"
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
         Left            =   3840
         TabIndex        =   12
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txt_b 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox t_cod 
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
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frm_vtasxgpo.frx":0F56
         Left            =   2040
         List            =   "frm_vtasxgpo.frx":0F7E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   3735
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
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
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "BASE (99=TODAS)"
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
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "COD. SERV.(99= TODOS)"
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
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Gpo.Mutual:"
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
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "FECHAS:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   2880
      Picture         =   "frm_vtasxgpo.frx":0FF0
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1695
   End
End
Attribute VB_Name = "frm_vtasxgpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1.Text = "SELECCION" Then
   Xwesvtas = 9
   frm_buscondesp.Show vbModal
Else
   Xwesvtas = 0
   Text1.Text = ""
   Text2.Text = ""
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_cod.SetFocus
End If

End Sub

Private Sub Command1_Click()
Dim Xlaedaqtiene As Long
Dim Xlamatmut As Long
Dim Xlacantmut As Integer

Xlamatmut = 0
Xlacantmut = 0

frm_vtasxgpo.MousePointer = 11

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\inftab.mdb")

MiBaseact.Execute "Delete * from convenio"
data_buscnv.RecordSource = "convenio"
data_buscnv.Refresh

Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"
data_inf2.RecordSource = "infcli"
data_inf2.Refresh

MiBaseact.Execute "Delete * from infvtas"
data_inf.RecordSource = "infvtas"
data_inf.Refresh

'24 coch A destino Balizas

data_conv.RecordSource = "Select * from convenio where cnv_grupo ='" & Combo1.Text & "'"
data_conv.Refresh
If data_conv.Recordset.RecordCount > 0 Then
    data_conv.Recordset.MoveFirst
    Do While Not data_conv.Recordset.EOF
       If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
          data_buscnv.Recordset.AddNew
          data_buscnv.Recordset("cnv_codigo") = data_conv.Recordset("cnv_codigo")
          data_buscnv.Recordset("cnv_desc") = data_conv.Recordset("cnv_desc")
          data_buscnv.Recordset("cnv_grupo") = data_conv.Recordset("cnv_grupo")
          data_buscnv.Recordset.Update
       End If
       data_conv.Recordset.MoveNext
    Loop
End If
If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
         If Combo1.Text <> "" Then
            If t_cod.Text = 99 Then
               If txt_b.Text = 99 Then
                  If Check1.Value = 1 Then
                     If Text1.Text = "" Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and tipo <>'" & "NOTA CR" & "'"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and tipo <>'" & "NOTA CR" & "' and convenio ='" & Text1.Text & "'"
                        data_lin.Refresh
                     End If
                  Else
                     If Text1.Text = "" Then
                        If Format(mh.Text, "yyyy/mm/dd") <= Format("01/10/2016", "yyyy/mm/dd") Then
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
                        Else
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and pendiente in ('F','T','X','Z')"
                        End If
                        data_lin.Refresh
                     Else
                        If Format(mh.Text, "yyyy-mm-dd") <= Format("01/10/2016") Then
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio ='" & Text1.Text & "'"
                        Else
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and pendiente in ('F','T','X','Z') and convenio ='" & Text1.Text & "'"
                        End If
                        data_lin.Refresh
                     End If
                  End If
               Else
                  If Check1.Value = 1 Then
                     If Text1.Text = "" Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "'"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "' and convenio ='" & Text1.Text & "'"
                        data_lin.Refresh
                     End If
                  Else
                     If Text1.Text = "" Then
                        If Format(mh.Text, "yyyy/mm/dd") <= Format("01/10/2016", "yyyy/mm/dd") Then
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and base =" & txt_b.Text
                        Else
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and base =" & txt_b.Text & " and pendiente in ('F','T','X','Z')"
                        End If
                        data_lin.Refresh
                     Else
                        If Format(mh.Text, "yyyy/mm/dd") <= Format("01/10/2016", "yyyy/mm/dd") Then
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and base =" & txt_b.Text & " and convenio ='" & Text1.Text & "'"
                        Else
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and base =" & txt_b.Text & " and pendiente in ('F','T','X','Z') and convenio ='" & Text1.Text & "'"
                        End If
                        data_lin.Refresh
                     End If
                  End If
               End If
            Else
               If txt_b.Text = 99 Then
                  If Check1.Value = 1 Then
                     If Text1.Text = "" Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cod_prod =" & t_cod.Text & " and tipo <>'" & "NOTA CR" & "'"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cod_prod =" & t_cod.Text & " and tipo <>'" & "NOTA CR" & "' and convenio ='" & Text1.Text & "'"
                        data_lin.Refresh
                     End If
                  Else
                     If Text1.Text = "" Then
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cod_prod =" & t_cod.Text & " and pendiente in ('F','T','X','Z')"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cod_prod =" & t_cod.Text & " and pendiente in ('F','T','X','Z') & " ' and convenio ='" & Text1.Text & "'"
                        data_lin.Refresh
                     End If
                  End If
               Else
                  If Check1.Value = 1 Then
                     If Text1.Text = "" Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cod_prod =" & t_cod.Text & " And base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "'"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cod_prod =" & t_cod.Text & " And base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "' and convenio ='" & Text1.Text & "'"
                        data_lin.Refresh
                     End If
                  Else
                     If Text1.Text = "" Then
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cod_prod =" & t_cod.Text & " And base =" & txt_b.Text & " and pendiente in ('F','T','X','Z')"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And cod_prod =" & t_cod.Text & " And base =" & txt_b.Text & " and pendiente in ('F','T','X','Z')" & "' and convenio ='" & Text1.Text & "'"
                        data_lin.Refresh
                     End If
                  End If
               End If
            End If
            If data_lin.Recordset.RecordCount > 0 Then
               data_lin.Recordset.MoveLast
               pb.Max = data_lin.Recordset.RecordCount + 1000
               pb.Value = 0
               data_lin.Recordset.MoveFirst
               DoEvents
               Do While Not data_lin.Recordset.EOF
                  If Text1.Text = "" Then
'                     data_buscnv.Recordset.FindFirst "cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
                     data_buscnv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
                     data_buscnv.Refresh
                     If data_buscnv.Recordset.RecordCount > 0 Then
                        If IsNull(data_buscnv.Recordset("cnv_grupo")) = False Then
                           If data_buscnv.Recordset("cnv_grupo") <> "" Then
                              If data_buscnv.Recordset("cnv_grupo") = Combo1.Text Then
                                 If data_lin.Recordset("cod_prod") = 999 Or _
                                    data_lin.Recordset("cod_prod") = 998 Or _
                                    data_lin.Recordset("cod_prod") = 993 Or _
                                    data_lin.Recordset("cod_prod") = 994 Or _
                                    data_lin.Recordset("cod_prod") = 995 Or _
                                    data_lin.Recordset("cod_prod") = 991 Or _
                                    data_lin.Recordset("cod_prod") = 61340 Or _
                                    data_lin.Recordset("cod_prod") = 997 Then
                                 Else
                                    If data_lin.Recordset("cod_prod") >= 13000 And _
                                       data_lin.Recordset("cod_prod") <= 13888 Then
                                    Else
                                       data_inf.Recordset.AddNew
                                       data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                                       data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                                       data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                                       data_inf.Recordset("nro_flia") = data_lin.Recordset("nro_flia")
                                       data_inf.Recordset("nro_med_a") = data_lin.Recordset("nro_med_a")
                                       If data_lin.Recordset("cod_prod") >= 60100 And data_lin.Recordset("cod_prod") <= 60109 Then
                                          data_inf.Recordset("cod_prod") = 6
                                          data_inf.Recordset("nom_prod") = "MEDICACION"
                                       Else
                                          data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                                          data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                                       End If
                                       data_inf.Recordset("convenio") = data_lin.Recordset("convenio")
                                       data_inf.Recordset("base") = data_lin.Recordset("base")
                                       data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                       If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                          data_inf.Recordset("nom_medic") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                       Else
                                          data_inf.Recordset("nom_medic") = "0"
                                       End If
                                       data_inf.Recordset.Update
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                  Else
                     If data_lin.Recordset("cod_prod") = 999 Or _
                        data_lin.Recordset("cod_prod") = 998 Or _
                        data_lin.Recordset("cod_prod") = 993 Or _
                        data_lin.Recordset("cod_prod") = 994 Or _
                        data_lin.Recordset("cod_prod") = 995 Or _
                        data_lin.Recordset("cod_prod") = 991 Or _
                        data_lin.Recordset("cod_prod") = 61340 Or _
                        data_lin.Recordset("cod_prod") = 997 Then
                     Else
                        If data_lin.Recordset("cod_prod") >= 13000 And _
                           data_lin.Recordset("cod_prod") <= 13888 Then
                        Else
                           data_inf.Recordset.AddNew
                           data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                           data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                           data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                           data_inf.Recordset("nro_flia") = data_lin.Recordset("nro_flia")
                           data_inf.Recordset("nro_med_a") = data_lin.Recordset("nro_med_a")
                           If data_lin.Recordset("cod_prod") >= 60100 And data_lin.Recordset("cod_prod") <= 60109 Then
                              data_inf.Recordset("cod_prod") = 6
                              data_inf.Recordset("nom_prod") = "MEDICACION"
                           Else
                              data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                              data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                           End If
                           data_inf.Recordset("convenio") = data_lin.Recordset("convenio")
                           data_inf.Recordset("base") = data_lin.Recordset("base")
                           data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                           If IsNull(data_lin.Recordset("ced_socio")) = False Then
                              data_inf.Recordset("nom_medic") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                           Else
                              data_inf.Recordset("nom_medic") = "0"
                           End If
                           data_inf.Recordset.Update
                        End If
                     End If
                  End If
                  data_lin.Recordset.MoveNext
                  pb.Value = pb.Value + 1
               Loop
'               MsgBox "Proceso finalizado"
               If Check3.Value = 1 Then
                  data_inf.RecordSource = "Select * from infvtas order by cod_cli"
                  data_inf.Refresh
                  If data_inf.Recordset.RecordCount > 0 Then
                     data_inf.Recordset.MoveFirst
                     Xlamatmut = data_inf.Recordset("cod_cli")
                     Do While Not data_inf.Recordset.EOF
                        If data_inf.Recordset("cod_cli") = Xlamatmut Then
                           Xlacantmut = Xlacantmut + 1
                           Xlamatmut = data_inf.Recordset("cod_cli")
                           data_inf.Recordset.MoveNext
                        Else
                           data_inf.Recordset.MovePrevious
                           data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_inf.Recordset("cod_cli")
                           data_cli.Refresh
                           If data_cli.Recordset.RecordCount > 0 Then
                              data_inf2.Recordset.AddNew
                              data_inf2.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                              data_inf2.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                              data_inf2.Recordset("cl_codconv") = data_inf.Recordset("convenio")
'                              data_conv.Recordset.FindFirst "cnv_codigo ='" & data_inf.Recordset("convenio") & "'"
                              data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_inf.Recordset("convenio") & "'"
                              data_conv.Refresh
                              If data_conv.Recordset.RecordCount > 0 Then
                                 data_inf2.Recordset("cl_nomconv") = data_conv.Recordset("cnv_desc")
                              Else
                                 data_inf2.Recordset("cl_nomconv") = data_inf.Recordset("convenio")
                              End If
                              
                              data_inf2.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                              data_inf2.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                              data_inf2.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                              data_inf2.Recordset("cl_nrocobr") = Xlacantmut
                              data_inf2.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                              data_inf2.Recordset.Update
                           End If
                           Xlacantmut = 0
                           data_inf.Recordset.MoveNext
                           Xlamatmut = data_inf.Recordset("cod_cli")
                        End If
'                        data_inf.Recordset.MoveNext

                     Loop
                     data_inf.Recordset.MovePrevious
                     data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_inf.Recordset("cod_cli")
                     data_cli.Refresh
                     If data_cli.Recordset.RecordCount > 0 Then
                        data_inf2.Recordset.AddNew
                        data_inf2.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf2.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_inf2.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_inf2.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
                        data_inf2.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                        data_inf2.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                        data_inf2.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                        data_inf2.Recordset("cl_nrocobr") = Xlacantmut
                        data_inf2.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                        data_inf2.Recordset.Update
                     End If
                     Xlacantmut = 0
                  
                  End If
               Else
                   If data_inf.Recordset.RecordCount > 0 Then
                      data_inf.Recordset.MoveLast
                      pb.Max = pb.Max + data_inf.Recordset.RecordCount - 1000
                      data_inf.Recordset.MoveFirst
                      DoEvents
                      Do While Not data_inf.Recordset.EOF
                         data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_inf.Recordset("cod_cli")
                         data_cli.Refresh
    '                     data_cli.Recordset.FindFirst "cl_codigo =" & data_inf.Recordset("cod_cli")
    '                     If Not data_cli.Recordset.NoMatch Then
                         If data_cli.Recordset.RecordCount > 0 Then
                            data_inf.Recordset.Edit
                            If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                               If data_cli.Recordset("cl_sexo") = 2 Then
                                  data_inf.Recordset("nom_medic") = "FEMENINO"
                               Else
                                  data_inf.Recordset("nom_medic") = "MASCULINO"
                               End If
                            Else
                               data_inf.Recordset("nom_medic") = "FEMENINO"
                            End If
                            If IsNull(data_cli.Recordset("cl_fnac")) = True Then
                               data_inf.Recordset("cod_medic") = 999
                               data_inf.Recordset("mes_paga") = 999
                               data_inf.Recordset("ano_paga") = 999
                               data_inf.Recordset("nom_superv") = "SIN DATOS"
                            Else
    ' Años = DateDiff("yyyy", Fecha_Nacimiento, Now)
                               Xlaedaqtiene = Date - data_cli.Recordset("cl_fnac")
                               Xlaedaqtiene = Xlaedaqtiene / 365
                               If Xlaedaqtiene < 0 Then
                                  Xlaedaqtiene = 0
                               End If
                               data_inf.Recordset("cod_medic") = Int(Xlaedaqtiene)
                               Xlaedaqtiene = data_inf.Recordset("fecha") - data_cli.Recordset("cl_fnac")
                               Xlaedaqtiene = Xlaedaqtiene / 365
                               data_inf.Recordset("mes_paga") = Int(Xlaedaqtiene) ' La edad que TENIA
                               If Xlaedaqtiene <= 0 Then
                                  data_inf.Recordset("ano_paga") = 0
                                  data_inf.Recordset("nom_superv") = "0 AÑOS"
                               Else
                                  If Xlaedaqtiene >= 1 And Xlaedaqtiene < 5 Then
                                     data_inf.Recordset("ano_paga") = 1
                                     data_inf.Recordset("nom_superv") = "1 A 4 AÑOS"
                                  Else
                                     If Xlaedaqtiene >= 5 And Xlaedaqtiene < 15 Then
                                        data_inf.Recordset("ano_paga") = 2
                                        data_inf.Recordset("nom_superv") = "5 a 14 AÑOS"
                                     Else
                                        If Xlaedaqtiene >= 15 And Xlaedaqtiene < 20 Then
                                           data_inf.Recordset("ano_paga") = 3
                                           data_inf.Recordset("nom_superv") = "15 a 19 AÑOS"
                                        Else
                                           If Xlaedaqtiene >= 20 And Xlaedaqtiene < 45 Then
                                              data_inf.Recordset("ano_paga") = 4
                                              data_inf.Recordset("nom_superv") = "20 a 44 AÑOS"
                                           Else
                                              If Xlaedaqtiene >= 45 And Xlaedaqtiene < 65 Then
                                                 data_inf.Recordset("ano_paga") = 5
                                                 data_inf.Recordset("nom_superv") = "45 a 64 AÑOS"
                                              Else
                                                 If Xlaedaqtiene >= 65 And Xlaedaqtiene < 75 Then
                                                    data_inf.Recordset("ano_paga") = 6
                                                    data_inf.Recordset("nom_superv") = "65 a 74 AÑOS"
                                                 Else
                                                    data_inf.Recordset("ano_paga") = 7
                                                    data_inf.Recordset("nom_superv") = "Mayor de 75"
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
                               If IsNull(data_cli.Recordset("cl_codced")) = False Then
                                  data_inf.Recordset("ruc") = Trim(str(data_cli.Recordset("cl_cedula"))) & "-" & Trim(str(data_cli.Recordset("cl_codced")))
                               Else
                                  data_inf.Recordset("ruc") = Trim(str(data_cli.Recordset("cl_cedula"))) & "-0"
                               End If
                            Else
                               data_inf.Recordset("ruc") = "0"
                            End If
    '                        data_inf.Recordset("nro_flia") = data_lin.Recordset("nro_flia")
                            data_inf.Recordset.Update
                         Else
                            data_inf.Recordset.Edit
                            data_inf.Recordset("cod_medic") = 0
                            data_inf.Recordset("nom_medic") = "FEMENINO"
                            data_inf.Recordset("mes_paga") = 999
                            data_inf.Recordset("ano_paga") = 999
                            data_inf.Recordset("nom_superv") = "SIN DATOS"
                            data_inf.Recordset("ruc") = 0
                            data_inf.Recordset.Update
                         End If
                         data_inf.Recordset.MoveNext
                         pb.Value = pb.Value + 1
                      Loop
                   End If
                   data_inf.RecordSource = "select * from infvtas order by cod_cli"
                   data_inf.Refresh
                   Dim Xlama, Xcanm As Integer
                   If data_inf.Recordset.RecordCount > 0 Then
                      data_inf.Recordset.MoveFirst
                      Xlama = 0
                      Xcanm = 0
                      Do While Not data_inf.Recordset.EOF
                         If Xlama = data_inf.Recordset("cod_cli") Then
                            Xcanm = Xcanm + 1
                            data_inf.Recordset.Edit
                            data_inf.Recordset("linea") = 0
                            data_inf.Recordset.Update
                         Else
                            If Xcanm >= 3 Then
                               data_inf.Recordset.Edit
                               data_inf.Recordset("linea") = Xcanm
                               data_inf.Recordset.Update
                            Else
                               data_inf.Recordset.Edit
                               data_inf.Recordset("linea") = 0
                               data_inf.Recordset.Update
                            End If
                            Xcanm = 1
                         End If
                         Xlama = data_inf.Recordset("cod_cli")
                         data_inf.Recordset.MoveNext
                      Loop
                   End If
               End If
               frm_vtasxgpo.MousePointer = 0
               If Check3.Value <> 1 Then
                  data_inf.RecordSource = "select * from infvtas order by fecha"
                  data_inf.Refresh
                  If Check2.Value = 1 Then
                     If Option2.Value = True Then
                        cr1.ReportFileName = App.path & "\infvtasgpocn.rpt"
                     Else
                        cr1.ReportFileName = App.path & "\infvtasgpoc.rpt"
                     End If
                  Else
                     cr1.ReportFileName = App.path & "\infvtasgpob.rpt"
                  End If
                  cr1.ReportTitle = "INFORME DE SERVICIOS A SOCIOS " & Combo1.Text & " FECHA: " & md.Text & " HASTA: " & mh.Text
                  cr1.Action = 1
               Else
                  cr1.ReportFileName = App.path & "\infservmut.rpt"
                  cr1.ReportTitle = "INFORME DE CANTIDAD DE SERVICIOS POR CLIENTE " & " FECHA: " & md.Text & " HASTA: " & mh.Text
                  cr1.Action = 1
               End If
            Else
               MsgBox "No existen registros"
            End If
         End If
   End If
End If
frm_vtasxgpo.MousePointer = 0
pb.Value = 0
If Check4.Value = 1 Then
   If Combo1.Text = "CCOU" Or Combo1.Text = "SMI" Or Combo1.Text = "H.EVANGELICO" Then
      data_infccou.DatabaseName = App.path & "\informess.mdb"
      data_infccou.RecordSource = "select * from infvtas"
      data_infccou.Refresh
      If data_infccou.Recordset.RecordCount > 0 Then
         data_infccou.Recordset.MoveFirst
         Do While Not data_infccou.Recordset.EOF
            data_infccou.Recordset.Delete
            data_infccou.Recordset.MoveNext
         Loop
      End If
      data_inf.RecordSource = "select * from infvtas where nro_flia <>" & 6
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         Do While Not data_inf.Recordset.EOF
            data_infccou.Recordset.AddNew
            data_infccou.Recordset("fecha") = data_inf.Recordset("fecha")
            data_infccou.Recordset("cod_cli") = data_inf.Recordset("cod_cli")
            data_infccou.Recordset("nom_cli") = data_inf.Recordset("nom_cli")
            data_infccou.Recordset("nro_flia") = data_inf.Recordset("nro_flia")
            data_infccou.Recordset("nro_med_a") = data_inf.Recordset("nro_med_a")
            data_infccou.Recordset("cod_prod") = data_inf.Recordset("cod_prod")
            data_infccou.Recordset("nom_prod") = data_inf.Recordset("nom_prod")
            data_infccou.Recordset("convenio") = data_inf.Recordset("convenio")
            data_infccou.Recordset("base") = data_inf.Recordset("base")
            data_infccou.Recordset("tot_lin") = data_inf.Recordset("tot_lin")
            data_infccou.Recordset("nom_medic") = data_inf.Recordset("nom_medic")
            data_infccou.Recordset("ano_paga") = data_inf.Recordset("ano_paga")
            data_infccou.Recordset("ruc") = data_inf.Recordset("ruc")

            data_infccou.Recordset.Update
            data_inf.Recordset.MoveNext
         Loop
         If Combo1.Text = "SMI" Then
            Command5_Click
         Else
            If Combo1.Text = "H.EVANGELICO" Then
               Command6_Click
            Else
               Command4_Click
            End If
         End If
      End If
   End If
End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
frm_verserv.Show vbModal

End Sub


Private Sub Command4_Click()

Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Textofecha As String
Dim Xtempo As Integer
Dim Cuentactrol As Integer
Dim Xlabrir3 As New Excel.Application

MsgBox "Se procesará informe para CCOU, aguarde..."

frm_vtasxgpo.MousePointer = 11
Cuentactrol = 0
Xlin = 1
XCol = 1
Xtotreg = 0
Xsub = 0
Set Xobjexel22 = New Excel.Application
Set Xlibexel22 = Xobjexel22.Workbooks.Add
Set Xarchexel22 = Xlibexel22.Worksheets.Add
Xarchexel22.Name = Trim("CCOU")
Xlibexel22.SaveAs ("C:\planillas\InfoCCOU.xls")
Xarchtex = "C:\planillas\InfoCCOU.xls"

Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
Xlin = Xlin + 1
XCol = XCol + 1
Xarchexel22.Range("A1", "C3").Font.Size = 16
Xarchexel22.Range("A" & Trim(str(Xlin)), "K" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)

Xarchexel22.Cells(Xlin, XCol) = "INFORMES SERVICIOS CCOU DESDE: " & md.Text & " HASTA: " & mh.Text
        
XCol = 1
Xlin = Xlin + 2
Xnrocan = Xnrocan + Xlin
        
Xarchexel22.Range("A" & Trim(str(Xlin)), "K" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS"
XCol = XCol + 1
Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 30
Xarchexel22.Cells(Xlin, XCol) = "MODALIDAD"
XCol = XCol + 1
Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 20
Xarchexel22.Cells(Xlin, XCol) = "LUGAR"
XCol = XCol + 1
Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 30
Xarchexel22.Cells(Xlin, XCol) = "TIPO DE ATENCION"
XCol = XCol + 1
Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "FECHA"
XCol = XCol + 1
Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
XCol = XCol + 1
Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 35
Xarchexel22.Cells(Xlin, XCol) = "NOMBRES"
XCol = XCol + 1
Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "SEXO"
XCol = XCol + 1
Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "GPO.EDAD"
XCol = XCol + 1
Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 35
Xarchexel22.Cells(Xlin, XCol) = "SERVICIO"
        
Xlin = Xlin + 1
XCol = 1
        
data_infccou.DatabaseName = App.path & "\informess.mdb"
data_infccou.RecordSource = "select * from infvtas where nro_flia =" & 6
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
      data_infccou.Recordset.Delete
      data_infccou.Recordset.MoveNext
   Loop
End If
   
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10001)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "MEDICINA GENERAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop
End If

data_infccou.RecordSource = "select * from infvtas where cod_prod in (14001)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "PEDIATRIA"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
data_infccou.RecordSource = "select * from infvtas where cod_prod in (2)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
      data_conv.RecordSource = "select * from medicos where med_cod =" & data_infccou.Recordset("nro_med_a")
      data_conv.Refresh
      If data_conv.Recordset.RecordCount > 0 Then
         If IsNull(data_conv.Recordset("med_esp")) = False Then
            data_infccou.Recordset.Edit
            data_infccou.Recordset("nom_med_a") = data_conv.Recordset("med_esp")
            data_infccou.Recordset.Update
         End If
      End If

      data_infccou.Recordset.MoveNext
   Loop
End If
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("nom_med_a")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_med_a")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "Sin Datos Esp"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10002)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "DOMICILIO"
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ano_paga")) = False Then
         If data_infccou.Recordset("ano_paga") > 14 Then
            Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO NO URGENTE ADULTOS"
         Else
            Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO NO URGENTE PEDIATRICOS"
         End If
      Else
         Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO NO URGENTE ADULTOS"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10003,10005)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ano_paga")) = False Then
         If data_infccou.Recordset("ano_paga") > 14 Then
            Xarchexel22.Cells(Xlin, XCol) = "URGENCIA CENTRALIZADA ADULTOS"
         Else
            Xarchexel22.Cells(Xlin, XCol) = "URGENCIA CENTRALIZADA PEDIATRIA"
         End If
      Else
         Xarchexel22.Cells(Xlin, XCol) = "URGENCIA CENTRALIZADA ADULTOS"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10004,10006)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "DOMICILIO"
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ano_paga")) = False Then
         If data_infccou.Recordset("ano_paga") > 14 Then
            Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO URGENTE ADULTOS"
         Else
            Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO URGENTE PEDIATRIA"
         End If
      Else
         Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO URGENTE ADULTOS"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10018,10050,14005,3)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA NO PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "DOMICILIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTA TELEFONICA"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
data_infccou.RecordSource = "select * from infvtas where nro_flia in (2)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "SERVICIOS ENFERMERIA"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "ENFERMERIA"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
   
Xarchexel22.Cells(Xlin, XCol) = "Total Registros: " & Trim(str(Xtotreg))
Xsub = 0
Xlin = Xlin + 1
XCol = 1
   
Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")

Xlibexel22.Save
Xlibexel22.Close
Xobjexel22.Quit
Xlabrir3.Workbooks.Open Xarchtex, , False
Xlabrir3.Visible = True
Xlabrir3.WindowState = xlMaximized

frm_vtasxgpo.MousePointer = 0

MsgBox "Proceso terminado"





End Sub

Private Sub Command5_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Textofecha As String
Dim Xtempo As Integer
Dim Cuentactrol As Integer
Dim Xlabrir3 As New Excel.Application

MsgBox "Se procesará informe para SMI, aguarde..."

frm_vtasxgpo.MousePointer = 11
Cuentactrol = 0
Xlin = 1
XCol = 1
Xtotreg = 0
Xsub = 0
Set Xobjexel22 = New Excel.Application
Set Xlibexel22 = Xobjexel22.Workbooks.Add
Set Xarchexel22 = Xlibexel22.Worksheets.Add
Xarchexel22.Name = Trim("SMI")
Xlibexel22.SaveAs ("C:\planillas\Infosmi.xls")
Xarchtex = "C:\planillas\Infosmi.xls"

Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
Xlin = Xlin + 1
XCol = XCol + 1
Xarchexel22.Range("A1", "C3").Font.Size = 16
Xarchexel22.Range("A" & Trim(str(Xlin)), "K" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)

Xarchexel22.Cells(Xlin, XCol) = "INFORMES SERVICIOS S.M.I. DESDE: " & md.Text & " HASTA: " & mh.Text
        
XCol = 1
Xlin = Xlin + 2
Xnrocan = Xnrocan + Xlin
        
Xarchexel22.Range("A" & Trim(str(Xlin)), "K" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS"
XCol = XCol + 1
Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 30
Xarchexel22.Cells(Xlin, XCol) = "MODALIDAD"
XCol = XCol + 1
Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 20
Xarchexel22.Cells(Xlin, XCol) = "LUGAR"
XCol = XCol + 1
Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 30
Xarchexel22.Cells(Xlin, XCol) = "TIPO DE ATENCION"
XCol = XCol + 1
Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "FECHA"
XCol = XCol + 1
Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
XCol = XCol + 1
Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 35
Xarchexel22.Cells(Xlin, XCol) = "NOMBRES"
XCol = XCol + 1
Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "SEXO"
XCol = XCol + 1
Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "GPO.EDAD"
XCol = XCol + 1
Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 35
Xarchexel22.Cells(Xlin, XCol) = "SERVICIO"
        
Xlin = Xlin + 1
XCol = 1
        
data_infccou.DatabaseName = App.path & "\informess.mdb"
data_infccou.RecordSource = "select * from infvtas where nro_flia =" & 6
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
      data_infccou.Recordset.Delete
      data_infccou.Recordset.MoveNext
   Loop
End If
   
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10001)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "MEDICINA GENERAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop
End If

data_infccou.RecordSource = "select * from infvtas where cod_prod in (14001)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "PEDIATRIA"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
data_infccou.RecordSource = "select * from infvtas where cod_prod in (2)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
      If IsNull(data_infccou.Recordset("nro_med_a")) = False Then
         data_conv.RecordSource = "select * from medicos where med_cod =" & data_infccou.Recordset("nro_med_a")
      Else
         data_conv.RecordSource = "select * from medicos where med_cod =" & 440
      End If
      data_conv.Refresh
      If data_conv.Recordset.RecordCount > 0 Then
         If IsNull(data_conv.Recordset("med_esp")) = False Then
            data_infccou.Recordset.Edit
            data_infccou.Recordset("nom_med_a") = data_conv.Recordset("med_esp")
            data_infccou.Recordset.Update
         End If
      End If
      data_infccou.Recordset.MoveNext
   Loop
End If
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("nom_med_a")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_med_a")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "Sin Datos Esp"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10002)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "DOMICILIO"
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ano_paga")) = False Then
         If data_infccou.Recordset("ano_paga") > 14 Then
            Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO NO URGENTE ADULTOS"
         Else
            Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO NO URGENTE PEDIATRICOS"
         End If
      Else
         Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO NO URGENTE ADULTOS"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10003,10005)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ano_paga")) = False Then
         If data_infccou.Recordset("ano_paga") > 14 Then
            Xarchexel22.Cells(Xlin, XCol) = "URGENCIA CENTRALIZADA ADULTOS"
         Else
            Xarchexel22.Cells(Xlin, XCol) = "URGENCIA CENTRALIZADA PEDIATRIA"
         End If
      Else
         Xarchexel22.Cells(Xlin, XCol) = "URGENCIA CENTRALIZADA ADULTOS"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10004,10006)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "DOMICILIO"
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ano_paga")) = False Then
         If data_infccou.Recordset("ano_paga") > 14 Then
            Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO URGENTE ADULTOS"
         Else
            Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO URGENTE PEDIATRIA"
         End If
      Else
         Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO URGENTE ADULTOS"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10018,10050,14005,3)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA NO PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "DOMICILIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTA TELEFONICA"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
data_infccou.RecordSource = "select * from infvtas where nro_flia in (2)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "SERVICIOS ENFERMERIA"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "ENFERMERIA"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
   
data_infccou.RecordSource = "select * from infvtas where nro_flia in (9)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "TRASLADOS"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "SERVICIOS TRASLADOS"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "TRASLADO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "TRASLADOS"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
data_infccou.RecordSource = "select * from infvtas where nro_flia in (3)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "SERVICIO DE LABORATORIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "LABORATORIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
data_infccou.RecordSource = "select * from infvtas where nro_flia in (5)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "SERVICIO ECOGRAFIAS"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "ECOGRAFIAS"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
Xarchexel22.Cells(Xlin, XCol) = "Total Registros: " & Trim(str(Xtotreg))
Xsub = 0
Xlin = Xlin + 1
XCol = 1
   
Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")

Xlibexel22.Save
Xlibexel22.Close
Xobjexel22.Quit
Xlabrir3.Workbooks.Open Xarchtex, , False
Xlabrir3.Visible = True
Xlabrir3.WindowState = xlMaximized

frm_vtasxgpo.MousePointer = 0

MsgBox "Proceso terminado"

End Sub

Private Sub Command6_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Textofecha As String
Dim Xtempo As Integer
Dim Cuentactrol As Integer
Dim Xlabrir3 As New Excel.Application

MsgBox "Se procesará informe para Evangélico, aguarde..."

frm_vtasxgpo.MousePointer = 11
Cuentactrol = 0
Xlin = 1
XCol = 1
Xtotreg = 0
Xsub = 0
Set Xobjexel22 = New Excel.Application
Set Xlibexel22 = Xobjexel22.Workbooks.Add
Set Xarchexel22 = Xlibexel22.Worksheets.Add
Xarchexel22.Name = Trim("HE")
Xlibexel22.SaveAs ("C:\planillas\InfoHE.xls")
Xarchtex = "C:\planillas\InfoHE.xls"

Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
Xlin = Xlin + 1
XCol = XCol + 1
Xarchexel22.Range("A1", "C3").Font.Size = 16
Xarchexel22.Range("A" & Trim(str(Xlin)), "K" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)

Xarchexel22.Cells(Xlin, XCol) = "INFORMES SERVICIOS H.Evangélico DESDE: " & md.Text & " HASTA: " & mh.Text
        
XCol = 1
Xlin = Xlin + 2
Xnrocan = Xnrocan + Xlin
        
Xarchexel22.Range("A" & Trim(str(Xlin)), "K" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS"
XCol = XCol + 1
Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 30
Xarchexel22.Cells(Xlin, XCol) = "MODALIDAD"
XCol = XCol + 1
Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 20
Xarchexel22.Cells(Xlin, XCol) = "LUGAR"
XCol = XCol + 1
Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 30
Xarchexel22.Cells(Xlin, XCol) = "TIPO DE ATENCION"
XCol = XCol + 1
Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "FECHA"
XCol = XCol + 1
Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
XCol = XCol + 1
Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 35
Xarchexel22.Cells(Xlin, XCol) = "NOMBRES"
XCol = XCol + 1
Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "SEXO"
XCol = XCol + 1
Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "GPO.EDAD"
XCol = XCol + 1
Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 35
Xarchexel22.Cells(Xlin, XCol) = "SERVICIO"
        
Xlin = Xlin + 1
XCol = 1
        
data_infccou.DatabaseName = App.path & "\informess.mdb"
data_infccou.RecordSource = "select * from infvtas where nro_flia =" & 6
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
      data_infccou.Recordset.Delete
      data_infccou.Recordset.MoveNext
   Loop
End If
   
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10001)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "MEDICINA GENERAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop
End If

data_infccou.RecordSource = "select * from infvtas where cod_prod in (14001)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "PEDIATRIA"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
data_infccou.RecordSource = "select * from infvtas where cod_prod in (2)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
      data_conv.RecordSource = "select * from medicos where med_cod =" & data_infccou.Recordset("nro_med_a")
      data_conv.Refresh
      If data_conv.Recordset.RecordCount > 0 Then
         If IsNull(data_conv.Recordset("med_esp")) = False Then
            data_infccou.Recordset.Edit
            data_infccou.Recordset("nom_med_a") = data_conv.Recordset("med_esp")
            data_infccou.Recordset.Update
         End If
      End If
'       data_infccou.Recordset.Edit
'       data_infccou.Recordset("nom_med_a") = "ESPECIALISTA"
'       data_infccou.Recordset.Update
      
      data_infccou.Recordset.MoveNext
   Loop
End If
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("nom_med_a")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_med_a")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "Sin Datos Esp"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10002)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "DOMICILIO"
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ano_paga")) = False Then
         If data_infccou.Recordset("ano_paga") > 14 Then
            Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO NO URGENTE ADULTOS"
         Else
            Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO NO URGENTE PEDIATRICOS"
         End If
      Else
         Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO NO URGENTE ADULTOS"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10003,10005)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ano_paga")) = False Then
         If data_infccou.Recordset("ano_paga") > 14 Then
            Xarchexel22.Cells(Xlin, XCol) = "URGENCIA CENTRALIZADA ADULTOS"
         Else
            Xarchexel22.Cells(Xlin, XCol) = "URGENCIA CENTRALIZADA PEDIATRIA"
         End If
      Else
         Xarchexel22.Cells(Xlin, XCol) = "URGENCIA CENTRALIZADA ADULTOS"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10004,10006)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "DOMICILIO"
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ano_paga")) = False Then
         If data_infccou.Recordset("ano_paga") > 14 Then
            Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO URGENTE ADULTOS"
         Else
            Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO URGENTE PEDIATRIA"
         End If
      Else
         Xarchexel22.Cells(Xlin, XCol) = "CONSULTA DOMICILIO URGENTE ADULTOS"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
data_infccou.RecordSource = "select * from infvtas where cod_prod in (10018,10050,14005,3)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA NO PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTAS NO URGENTES"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "DOMICILIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTA TELEFONICA"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
'data_infccou.RecordSource = "select * from infvtas where nro_flia in (2)"
'data_infccou.Refresh
'If data_infccou.Recordset.RecordCount > 0 Then
'   data_infccou.Recordset.MoveFirst
'   Do While Not data_infccou.Recordset.EOF
'
'      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = "SERVICIOS ENFERMERIA"
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = "ENFERMERIA"
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
'      XCol = XCol + 1
'      If IsNull(data_infccou.Recordset("ruc")) = False Then
'         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
'      Else
'         Xarchexel22.Cells(Xlin, XCol) = "0-0"
'      End If
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
'      Xlin = Xlin + 1
'      XCol = 1
'      Xtotreg = Xtotreg + 1
'      Xsub = Xsub + 1
'      data_infccou.Recordset.MoveNext
'   Loop

'End If
   
'data_infccou.RecordSource = "select * from infvtas where nro_flia in (2)"
'data_infccou.Refresh
'If data_infccou.Recordset.RecordCount > 0 Then
'   data_infccou.Recordset.MoveFirst
'   Do While Not data_infccou.Recordset.EOF
              
'      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = "SERVICIOS LABORATORIO"
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = "LABORATORIO"
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
'      XCol = XCol + 1
'      If IsNull(data_infccou.Recordset("ruc")) = False Then
'         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
'      Else
'         Xarchexel22.Cells(Xlin, XCol) = "0-0"
'      End If
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
'
'      Xlin = Xlin + 1
'      XCol = 1
'      Xtotreg = Xtotreg + 1
'      Xsub = Xsub + 1
'      data_infccou.Recordset.MoveNext
'   Loop

'End If
   
data_infccou.RecordSource = "select * from infvtas where nro_flia in (9)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "TRASLADOS"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "SERVICIOS TRASLADOS"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "TRASLADO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "TRASLADOS"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
'data_infccou.RecordSource = "select * from infvtas where nro_flia in (3)"
'data_infccou.Refresh
'If data_infccou.Recordset.RecordCount > 0 Then
'   data_infccou.Recordset.MoveFirst
'   Do While Not data_infccou.Recordset.EOF
'
'      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = "SERVICIO DE LABORATORIO"
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = "LABORATORIO"
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
'      XCol = XCol + 1
'      If IsNull(data_infccou.Recordset("ruc")) = False Then
'         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
'      Else
'         Xarchexel22.Cells(Xlin, XCol) = "0-0"
'      End If
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
'      XCol = XCol + 1
'      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
'
'      Xlin = Xlin + 1
'      XCol = 1
'      Xtotreg = Xtotreg + 1
'      Xsub = Xsub + 1
'      data_infccou.Recordset.MoveNext
'   Loop

'End If
   
data_infccou.RecordSource = "select * from infvtas where nro_flia in (5)"
data_infccou.Refresh
If data_infccou.Recordset.RecordCount > 0 Then
   data_infccou.Recordset.MoveFirst
   Do While Not data_infccou.Recordset.EOF
              
      Xarchexel22.Cells(Xlin, XCol) = "ATENCIÓN AMBULATORIA PRESENCIAL"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "SERVICIO ECOGRAFIAS"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "CONSULTORIO"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "ECOGRAFIAS"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_infccou.Recordset("fecha"), "dd/mm/yyyy"))
      XCol = XCol + 1
      If IsNull(data_infccou.Recordset("ruc")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ruc")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_medic")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("ano_paga")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_infccou.Recordset("nom_prod")
              
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Xsub = Xsub + 1
      data_infccou.Recordset.MoveNext
   Loop

End If
   
Xarchexel22.Cells(Xlin, XCol) = "Total Registros: " & Trim(str(Xtotreg))
Xsub = 0
Xlin = Xlin + 1
XCol = 1
   
Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")

Xlibexel22.Save
Xlibexel22.Close
Xobjexel22.Quit
Xlabrir3.Workbooks.Open Xarchtex, , False
Xlabrir3.Visible = True
Xlabrir3.WindowState = xlMaximized

frm_vtasxgpo.MousePointer = 0

MsgBox "Proceso terminado"

End Sub

Private Sub Form_Load()
data_lin.ConnectionString = "dsn=" & Xconexrmt
data_inf.DatabaseName = App.path & "\informes.mdb"
data_inf.RecordSource = "infvtas"
data_inf.Refresh
data_inf2.DatabaseName = App.path & "\informes.mdb"

data_conv.ConnectionString = "dsn=" & Xconexrmt
'data_conv.RecordSource = "convenio"
'data_conv.Refresh
data_cli.ConnectionString = "dsn=" & Xconexrmt
'data_cli.RecordSource = "clientes"
'data_cli.Refresh
data_buscnv.DatabaseName = App.path & "\inftab.mdb"
data_buscnv.RecordSource = "convenio"
data_buscnv.Refresh


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

Private Sub t_cod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_b.SetFocus
End If

End Sub

Private Sub txt_b_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub
