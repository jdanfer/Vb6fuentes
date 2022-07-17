VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infdemoras 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demoras en llamados"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6975
   Icon            =   "frm_infdemoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc data_llama 
      Height          =   330
      Left            =   3000
      Top             =   5280
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "data_llama"
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
      Left            =   2400
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   6240
      Picture         =   "frm_infdemoras.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   5040
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
      Left            =   240
      Picture         =   "frm_infdemoras.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Procesar"
      Top             =   5040
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos del informe"
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
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1200
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1560
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FF0000&
         Caption         =   "Generar promedio de demoras en domicilio"
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
         Left            =   120
         TabIndex        =   26
         Top             =   3480
         Width           =   4215
      End
      Begin VB.TextBox t_codmed 
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
         Height          =   375
         Left            =   4680
         TabIndex        =   25
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox t_mov 
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
         Height          =   405
         Left            =   4680
         TabIndex        =   23
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   375
         Left            =   4560
         TabIndex        =   21
         Top             =   3960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FF0000&
         Caption         =   "Ver demoras por médico"
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
         Left            =   120
         TabIndex        =   20
         Top             =   3000
         Width           =   3255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   375
         Left            =   4800
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0080FF80&
         Caption         =   "Generar planilla de totales"
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
         TabIndex        =   18
         Top             =   3960
         Width           =   3135
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frm_infdemoras.frx":0F56
         Left            =   1680
         List            =   "frm_infdemoras.frx":0F66
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF0000&
         Caption         =   "Sin llamados cancelados"
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
         Left            =   3600
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   2775
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
         Left            =   3360
         TabIndex        =   13
         Top             =   4440
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   4440
         Value           =   -1  'True
         Width           =   2175
      End
      Begin MSMask.MaskEdBox mhh 
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Top             =   960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mhd 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
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
         ItemData        =   "frm_infdemoras.frx":0F8B
         Left            =   1680
         List            =   "frm_infdemoras.frx":0F9E
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3720
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
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   1680
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
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         Caption         =   "Médico:"
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
         Left            =   3480
         TabIndex        =   24
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF0000&
         Caption         =   "Móvil:"
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
         Left            =   3480
         TabIndex        =   22
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Códigos:"
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
         TabIndex        =   16
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   0
         X2              =   6480
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Horario:"
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
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Zona:"
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
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Fechas:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.OLE OLE1 
      AutoActivate    =   1  'GetFocus
      DisplayType     =   1  'Icon
      Height          =   975
      Left            =   5880
      OleObjectBlob   =   "frm_infdemoras.frx":0FC6
      SourceDoc       =   "C:\sappmys\sappmysql\DEMORAS.txt"
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   3960
      Picture         =   "frm_infdemoras.frx":2BDE
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1695
   End
End
Attribute VB_Name = "frm_infdemoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check4_Click()
If Check4.Value = 1 Then
   If Check3.Value = 1 Then
      Check3.Value = 0
   End If
End If

End Sub

Private Sub Command1_Click()
Dim xhh, xmm As Integer
Dim Xhhh, Xmmh, Xcoddelmed As Integer
Dim Xelprommed As Double
Dim xdemh, xdemm, xcuenta As Integer
Dim Xnom, Xnommed, Xquetex, Xmedunavez As String
Dim Xnromov, Xmotcon, XCat, Xdescol As String
Dim Xzona, Xedad, Xcadllama As String
Dim Xtotal, Xtotalmins As Long
Dim Xtotgral, Xhasta30, Xdemmas30, Xdemmas1, Xdemmas2 As Long
Dim Xtotresz1, Xtotresz2, Xtotresz3 As Long
Dim Xminhs, Xtotporme, Xtothsporme As Long
Xminhs = 0
Command1.Enabled = False
Command2.Enabled = False

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from inflla"

data_inf.RecordSource = "inflla"
data_inf.Refresh

Data1.RecordSource = "inflla"
Data1.Refresh
Xtotalmins = 0
Xquetex = "50"
If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      frm_infdemoras.MousePointer = 11
      If Combo1.ListIndex = 0 Then
         If Combo2.ListIndex = 0 Then
            If t_mov.Text <> "" Then
               data_llama.RecordSource = "Select * from llamado where codmot ='" & "V" & "' and codzon =" & 1 & " And fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Xquetex & "' and movilpas =" & t_mov.Text & " and cancela is null order by fecha,codzon"
               data_llama.Refresh
            Else
               data_llama.RecordSource = "Select * from llamado where codmot ='" & "V" & "' and codzon =" & 1 & " And fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Xquetex & "' and cancela is null order by fecha,codzon"
               data_llama.Refresh
            End If
         Else
            data_llama.RecordSource = "Select * from llamado where codzon =" & 1 & " And fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Xquetex & "' and cancela is null order by fecha,codzon"
            data_llama.Refresh
         End If
      Else
         If Combo1.ListIndex = 1 Then
            If Combo2.ListIndex = 0 Then
               data_llama.RecordSource = "Select * from llamado where codmot ='" & "V" & "' and codzon =" & 2 & " And fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Xquetex & "' and cancela is null order by fecha,codzon"
               data_llama.Refresh
            Else
               data_llama.RecordSource = "Select * from llamado where codzon =" & 2 & " And fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Xquetex & "' and cancela is null order by fecha,codzon"
               data_llama.Refresh
            End If
         Else
            If Combo1.ListIndex = 2 Then
               If Combo2.ListIndex = 0 Then
                  data_llama.RecordSource = "Select * from llamado where codmot ='" & "V" & "' and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And codzon in (1,2,3,5) And categ <>'" & Xquetex & "' and cancela is null order by fecha,codzon"
                  data_llama.Refresh
               Else
                  data_llama.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And codzon in (1,2,3,5) And categ <>'" & Xquetex & "' and cancela is null order by fecha,codzon"
                  data_llama.Refresh
               End If
            Else
               If Combo1.ListIndex = 3 Then
                  If Combo2.ListIndex = 0 Then
                     data_llama.RecordSource = "Select * from llamado where codmot ='" & "V" & "' and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And codzon =" & 3 & " And categ <>'" & Xquetex & "' and cancela is null order by fecha,codzon"
                     data_llama.Refresh
                  Else
                     data_llama.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And codzon =" & 3 & " And categ <>'" & Xquetex & "' and cancela is null order by fecha,codzon"
                     data_llama.Refresh
                  End If
               Else
                  If Combo2.ListIndex = 0 Then 'aca
                     data_llama.RecordSource = "Select * from llamado where codmot ='" & "V" & "' and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Xquetex & "' and codzon in (1,2,3,5) and cancela is null and codmed <>" & 959 & " and categ not in ('SAMC','50','UDEMM','CERSEM','CERADU','CERDGI','CERSEM','CERHEV','CERCAS','CERMAT','CERKEV','CERIMP','CERSEV','CERVIS') and movilpas not in (2015,0) order by fecha,codzon"
                     data_llama.Refresh
                  Else
                     If Combo2.ListIndex = 1 Then
                        data_llama.RecordSource = "Select * from llamado where codmot ='" & "A" & "' and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Xquetex & "' and codzon in (1,2,3,5) and cancela is null order by fecha,codzon"
                        data_llama.Refresh
                     Else
                        If Combo2.ListIndex = 2 Then
                           data_llama.RecordSource = "Select * from llamado where codmot ='" & "R" & "' and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Xquetex & "' and codzon in (1,2,3,5) and cancela is null order by fecha,codzon"
                           data_llama.Refresh
                        Else
                           If t_mov.Text <> "" Then
                              data_llama.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Xquetex & "' and codzon in (1,2,3,5) and movilpas =" & t_mov.Text & " and codmed <>" & 959 & " and cancela is null order by fecha,codzon"
                              data_llama.Refresh
                           Else
                              If Combo2.Text = "Amarillos" Then
                                 data_llama.RecordSource = "Select * from llamado where codmot in ('A') and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codzon in (1,2,3,5) and codmed <>" & 959 & " and categ not in ('SAMC','50','UDEMM','CERSEM','CERADU','CERDGI','CERSEM','CERHEV','CERCAS','CERMAT','CERKEV','CERIMP','CERSEV','CERVIS') and cancela is null and movilpas not in (2015) order by fecha,codzon"
                                 data_llama.Refresh
                              Else
                                 data_llama.RecordSource = "Select * from llamado where codmot in ('V') and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codzon in (1,2,3,5) and codmed <>" & 959 & " and categ not in ('SAMC','50','UDEMM','CERSEM','CERADU','CERDGI','CERSEM','CERHEV','CERCAS','CERMAT','CERKEV','CERIMP','CERSEV','CERVIS') and cancela is null and movilpas not in (2015,0) order by fecha,codzon"
'                                data_llama.RecordSource = "Select * from llamado where codmot in ('V','Z') and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codzon in (1,2,3) and codmed <>" & 959 & " and categ not in ('SAMC','50','UDEMM','CERSEM','CERADU','CERDGI','CERSEM','CERHEV','CERCAS','CERMAT','CERKEV','CERIMP','CERSEV','CERVIS') and movilpas not in (99) order by fecha,codzon"
                                 data_llama.Refresh
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
      If t_codmed.Text <> "" Then
         data_llama.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " And categ <>'" & Xquetex & "' and codzon in (1,2,3,5) and codmed =" & t_codmed.Text & " and cancela is null order by fecha,codzon"
         data_llama.Refresh
      End If
      If data_llama.Recordset.RecordCount > 0 Then
         If Option1.Value = True Then
             data_llama.Recordset.MoveFirst
             Do While Not data_llama.Recordset.EOF
                If data_llama.Recordset("codzon") = 4 Then
                Else
                    data_inf.Recordset.AddNew
                    data_inf.Recordset("nro") = data_llama.Recordset("nro")
                    data_inf.Recordset("fecha") = data_llama.Recordset("fecha")
                    data_inf.Recordset("hora") = data_llama.Recordset("hora")
                    data_inf.Recordset("usuario") = data_llama.Recordset("usuario")
                    If data_llama.Recordset("matric") >= 999999999 Then
                       data_inf.Recordset("matric") = 0
                    Else
                       data_inf.Recordset("matric") = data_llama.Recordset("matric")
                    End If
                    data_inf.Recordset("nombre") = data_llama.Recordset("nombre")
                    data_inf.Recordset("edad") = data_llama.Recordset("edad")
                    data_inf.Recordset("unied") = data_llama.Recordset("unied")
                    data_inf.Recordset("categ") = data_llama.Recordset("categ")
                    data_inf.Recordset("nomcat") = data_llama.Recordset("nomcat")
                    If data_llama.Recordset("ci") >= 999999999 Then
                       data_inf.Recordset("ci") = 0
                    Else
                       data_inf.Recordset("ci") = data_llama.Recordset("ci")
                    End If
                    data_inf.Recordset("direcc") = data_llama.Recordset("direcc")
                    data_inf.Recordset("telef") = data_llama.Recordset("telef")
                    data_inf.Recordset("codzon") = data_llama.Recordset("codzon")
                    data_inf.Recordset("base") = data_llama.Recordset("base")
                    data_inf.Recordset("referen") = data_llama.Recordset("referen")
                    data_inf.Recordset("motcon") = data_llama.Recordset("motcon")
                    data_inf.Recordset("obsmot") = data_llama.Recordset("obsmot")
                    data_inf.Recordset("codmot") = data_llama.Recordset("codmot")
                    data_inf.Recordset("descol") = data_llama.Recordset("descol")
                    data_inf.Recordset("movilpas") = data_llama.Recordset("movilpas")
                    data_inf.Recordset("pend") = data_llama.Recordset("pend")
                    If IsNull(data_llama.Recordset("fec_rea")) = True Then
                       data_inf.Recordset("fec_rea") = data_llama.Recordset("fecpas")
                    Else
                       data_inf.Recordset("fec_rea") = data_llama.Recordset("fec_rea")
                    End If
                    If IsNull(data_llama.Recordset("hor_rea")) = True Then
                       data_inf.Recordset("hor_rea") = data_llama.Recordset("horpas")
                    Else
                       data_inf.Recordset("hor_rea") = data_llama.Recordset("hor_rea")
                    End If
                    data_inf.Recordset("fecpas") = data_llama.Recordset("fecpas")
                    data_inf.Recordset("horpas") = data_llama.Recordset("horpas")
                    data_inf.Recordset("fecsali") = data_llama.Recordset("fecsali")
                    data_inf.Recordset("horsali") = data_llama.Recordset("horsali")
                    If IsNull(data_llama.Recordset("fec_llega")) = True Then
                       data_inf.Recordset("fec_llega") = data_llama.Recordset("fecpas")
                    Else
                       data_inf.Recordset("fec_llega") = data_llama.Recordset("fec_llega")
                    End If
                    If IsNull(data_llama.Recordset("hor_llega")) = True Then
                       data_inf.Recordset("hor_llega") = data_llama.Recordset("horpas")
                    Else
                       data_inf.Recordset("hor_llega") = data_llama.Recordset("hor_llega")
                    End If
                    data_inf.Recordset("diag") = data_llama.Recordset("diag")
                    data_inf.Recordset("colormot") = data_llama.Recordset("colormot")
                    data_inf.Recordset("codmed") = data_llama.Recordset("codmed")
                    data_inf.Recordset("obs") = data_llama.Recordset("obs")
                    data_inf.Recordset("nommed") = data_llama.Recordset("nommed")
                    data_inf.Recordset("trasla") = data_llama.Recordset("trasla")
                    data_inf.Recordset("lugar") = data_llama.Recordset("lugar")
                    data_inf.Recordset("hsald") = data_llama.Recordset("hsald")
                    data_inf.Recordset("hllega") = data_llama.Recordset("hllega")
                    data_inf.Recordset("hzona") = data_llama.Recordset("hzona")
                    data_inf.Recordset("movil_rea") = data_llama.Recordset("movil_rea")
                    data_inf.Recordset("totdem") = data_llama.Recordset("totdem")
                    data_inf.Recordset("totend") = data_llama.Recordset("totend")
                    data_inf.Recordset("cancela") = data_llama.Recordset("cancela")
                    data_inf.Recordset.Update
                End If
                data_llama.Recordset.MoveNext
             Loop
             data_inf.RecordSource = "Select * from inflla"
             data_inf.Refresh
             If Check3.Value = 1 Then
                Command5_Click
             Else
                 If data_inf.Recordset.RecordCount > 0 Then
                    data_inf.Recordset.MoveFirst
                    Do While Not data_inf.Recordset.EOF
                       If IsNull(data_inf.Recordset("cancela")) = False Then
                          If data_inf.Recordset("cancela") = 1 Then
                             data_inf.Recordset.Delete
                          End If
                       End If
                       data_inf.Recordset.MoveNext
                    Loop
                 End If
                 data_inf.Recordset.MoveFirst
                 Xcadllama = "INFORME DEMORAS EN LLAMADOS A PARTIR DE 30 MINUTOS"
                 Open App.path & "\DEMORAS.txt" For Output As #1
                 If Check4.Value <> 1 Then
                     
                     Print #1, Xcadllama
                     Print #1, "===================================================="
            '         Print #1, "FECHA/HORA REC.   N O M B R E S                                   EDAD  H.REA. DEMORA   MEDICO               MOVIL  MOTIVO CONSULTA                         CATEG.  COLOR     ZONA"
            '         Print #1, "===================================================================================================================================================================================="
                     Xcadllama = ""
                     Do While Not data_inf.Recordset.EOF
                        If IsNull(data_inf.Recordset("hor_llega")) = True Then
                           data_inf.Recordset.MoveNext
                        Else
                            If IsNull(data_inf.Recordset("hora")) = False Then
                               xhh = Val(Mid(data_inf.Recordset("hora"), 1, 2))
                               xmm = Val(Mid(data_inf.Recordset("hora"), 4, 2))
                            End If
                            If IsNull(data_inf.Recordset("hor_llega")) = False Then
                               Xhhh = Val(Mid(data_inf.Recordset("hor_llega"), 1, 2))
                               Xmmh = Val(Mid(data_inf.Recordset("hor_llega"), 4, 2))
                            End If
                            xdemh = Xhhh - xhh
                            xdemm = Xmmh - xmm
                            If data_inf.Recordset("fecha") < data_inf.Recordset("fec_llega") Then
                               xdemh = Xhhh - xhh
                               xdemh = xdemh + 24
                            End If
                            If xdemh > 0 Then
                               If xdemm < 0 Then
                                  xdemm = xdemm + 60
                                  xdemh = xdemh - 1
                               End If
                            Else
                               If xdemm < 0 Then
                                  xdemm = xdemm + 60
                               End If
                            End If
                            data_inf.Recordset.Edit
                            If xdemh > 9 Then
                               If xdemm > 9 Then
                                  data_inf.Recordset("totdem") = Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                               Else
                                  data_inf.Recordset("totdem") = Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                               End If
                            Else
                               If xdemm > 9 Then
                                  If xdemh < 0 Then
                                     xdemh = 0
                                  End If
                                  data_inf.Recordset("totdem") = "0" + Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                               Else
                                  If xdemh < 0 Then
                                     xdemh = 0
                                  End If
                                  data_inf.Recordset("totdem") = "0" + Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                               End If
                            End If
                            If xdemh > 0 Then
                               Xminhs = xdemh * 60
                            Else
                               Xminhs = 0
                            End If
    '                        data_inf.Recordset("mes") = Xminhs + xdemm
                               
                            data_inf.Recordset.Update
                            If Len(data_inf.Recordset("movilpas")) = 1 Then
                               Xnromov = "  " + str(data_inf.Recordset("movilpas"))
                            End If
                            If Len(data_inf.Recordset("movilpas")) = 2 Then
                               Xnromov = " " + str(data_inf.Recordset("movilpas"))
                            End If
                            If Len(data_inf.Recordset("movilpas")) = 3 Then
                               Xnromov = str(data_inf.Recordset("movilpas"))
                            End If
                            If Len(data_inf.Recordset("edad")) = 1 Then
                               Xedad = "  " + str(data_inf.Recordset("edad"))
                            End If
                            If Len(data_inf.Recordset("edad")) = 2 Then
                               Xedad = " " + str(data_inf.Recordset("edad"))
                            End If
                            If Len(data_inf.Recordset("edad")) = 3 Then
                               Xedad = str(data_inf.Recordset("edad"))
                            End If
                            Xmotcon = data_inf.Recordset("motcon")
                            XCat = data_inf.Recordset("categ")
                            Xdescol = data_inf.Recordset("descol")
                            Xnom = data_inf.Recordset("nombre")
                            If IsNull(data_inf.Recordset("nommed")) = False Then
                               Xnommed = data_inf.Recordset("nommed")
                            Else
                               Xnommed = ""
                            End If
                            If Xnom <> "" Then
                               Xnom = Mid(Xnom, 1, 50)
                               xcuenta = Len(Xnom)
                               xcuenta = xcuenta + 1
                               For xcuenta = xcuenta To 50
                                   Xnom = Xnom + " "
                               Next
                            Else
                               Xnom = "                                                  "
                            End If
                            If Xnommed <> "" Then
                               Xnommed = Mid(Xnommed, 1, 25)
                               xcuenta = Len(Xnommed)
                               xcuenta = xcuenta + 1
                               For xcuenta = xcuenta To 25
                                   Xnommed = Xnommed + " "
                               Next
                            Else
                               Xnommed = "                         "
                            End If
                            If Xmotcon <> "" Then
                               Xmotcon = Mid(Xmotcon, 1, 40)
                               xcuenta = Len(Xmotcon)
                               xcuenta = xcuenta + 1
                               For xcuenta = xcuenta To 40
                                   Xmotcon = Xmotcon + " "
                               Next
                            Else
                               Xmotcon = "                                        "
                            End If
                            If XCat <> "" Then
                               XCat = Mid(XCat, 1, 6)
                               xcuenta = Len(XCat)
                               xcuenta = xcuenta + 1
                               For xcuenta = xcuenta To 6
                                   XCat = XCat + " "
                               Next
                            Else
                               XCat = "      "
                            End If
                            If Xdescol <> "" Then
                               Xdescol = Mid(Xdescol, 1, 9)
                               xcuenta = Len(Xdescol)
                               xcuenta = xcuenta + 1
                               For xcuenta = xcuenta To 9
                                   Xdescol = Xdescol + " "
                               Next
                            Else
                               Xdescol = "         "
                            End If
                            If data_inf.Recordset("codzon") = 1 Then
                               Xzona = "Z.COSTA"
                            Else
                               Xzona = "Z.NORTE"
                            End If
    '                        If data_inf.Recordset("totdem") >= "00:00" And data_inf.Recordset("totdem") <= "00:30" Then
    '                           Xtotal = Xtotal + 1
    '                            Print #1, CStr(data_inf.Recordset("fecha")) + " " + data_inf.Recordset("hora") + " " + Xnom + " " + Xedad _
    '                            ; " " + data_inf.Recordset("hor_llega") + " " + data_inf.Recordset("totdem") + " " + Xnommed _
    '                            + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama + " " + Trim(Str(data_inf.Recordset("mes")))
    '                           Xtotgral = Xtotgral + 1
    '                        End If
                            data_inf.Recordset.MoveNext
                        End If
                     Loop
                     Print #1, "================================================="
                     Print #1, "TOTAL......:" + str(Xtotal)
                     Print #1, "================================================="
                     Xhasta30 = Xtotal
                     Xtotal = 0
                     data_inf.Recordset.MoveFirst
                     Xcadllama = "DEMORAS EN LLAMADOS MAS DE 30 MINUTOS A 1 HORA"
                     Print #1, Xcadllama
                     Print #1, "================================================="
                     Print #1, "FECHA/HORA REC.   N O M B R E S                                   EDAD  H.REA. DEMORA   MEDICO               MOVIL  MOTIVO CONSULTA                         CATEG.  COLOR     ZONA"
                     Print #1, "===================================================================================================================================================================================="
                     Xcadllama = ""
                     Do While Not data_inf.Recordset.EOF
                        If IsNull(data_inf.Recordset("hor_llega")) = True Then
                           data_inf.Recordset.MoveNext
                        Else
                            If IsNull(data_inf.Recordset("hora")) = False Then
                               xhh = Val(Mid(data_inf.Recordset("hora"), 1, 2))
                               xmm = Val(Mid(data_inf.Recordset("hora"), 4, 2))
                            End If
                            If IsNull(data_inf.Recordset("hor_llega")) = False Then
                               Xhhh = Val(Mid(data_inf.Recordset("hor_llega"), 1, 2))
                               Xmmh = Val(Mid(data_inf.Recordset("hor_llega"), 4, 2))
                            End If
                            xdemh = Xhhh - xhh
                            xdemm = Xmmh - xmm
                            If data_inf.Recordset("fecha") < data_inf.Recordset("fec_llega") Then
                               xdemh = Xhhh - xhh
                               xdemh = xdemh + 24
                            End If
                            If xdemh > 0 Then
                               If xdemm < 0 Then
                                  xdemm = xdemm + 60
                                  xdemh = xdemh - 1
                               End If
                            Else
                               If xdemm < 0 Then
                                  xdemm = xdemm + 60
                               End If
                            End If
                            data_inf.Recordset.Edit
                            If xdemh > 9 Then
                               If xdemm > 9 Then
                                  data_inf.Recordset("totdem") = Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                               Else
                                  data_inf.Recordset("totdem") = Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                               End If
                            Else
                               If xdemm > 9 Then
                                  If xdemh < 0 Then
                                     xdemh = 0
                                  End If
                                  data_inf.Recordset("totdem") = "0" + Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                               Else
                                  If xdemh < 0 Then
                                     xdemh = 0
                                  End If
                                  data_inf.Recordset("totdem") = "0" + Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                               End If
                            End If
                            If xdemh > 0 Then
                               Xminhs = xdemh * 60
                            Else
                               Xminhs = 0
                            End If
    '                        data_inf.Recordset("mes") = Xminhs + xdemm
                            
                            data_inf.Recordset.Update
                            If Len(data_inf.Recordset("movilpas")) = 1 Then
                               Xnromov = "  " + str(data_inf.Recordset("movilpas"))
                            End If
                            If Len(data_inf.Recordset("movilpas")) = 2 Then
                               Xnromov = " " + str(data_inf.Recordset("movilpas"))
                            End If
                            If Len(data_inf.Recordset("movilpas")) = 3 Then
                               Xnromov = str(data_inf.Recordset("movilpas"))
                            End If
                            If Len(data_inf.Recordset("edad")) = 1 Then
                               Xedad = "  " + str(data_inf.Recordset("edad"))
                            End If
                            If Len(data_inf.Recordset("edad")) = 2 Then
                               Xedad = " " + str(data_inf.Recordset("edad"))
                            End If
                            If Len(data_inf.Recordset("edad")) = 3 Then
                               Xedad = str(data_inf.Recordset("edad"))
                            End If
                            Xmotcon = data_inf.Recordset("motcon")
                            XCat = data_inf.Recordset("categ")
                            Xdescol = data_inf.Recordset("descol")
                            Xnom = data_inf.Recordset("nombre")
                            If IsNull(data_inf.Recordset("nommed")) = False Then
                               Xnommed = data_inf.Recordset("nommed")
                            Else
                               Xnommed = ""
                            End If
                            If Xnom <> "" Then
                               Xnom = Mid(Xnom, 1, 50)
                               xcuenta = Len(Xnom)
                               xcuenta = xcuenta + 1
                               For xcuenta = xcuenta To 50
                                   Xnom = Xnom + " "
                               Next
                            Else
                               Xnom = "                                                  "
                            End If
                            If Xnommed <> "" Then
                               Xnommed = Mid(Xnommed, 1, 25)
                               xcuenta = Len(Xnommed)
                               xcuenta = xcuenta + 1
                               For xcuenta = xcuenta To 25
                                   Xnommed = Xnommed + " "
                               Next
                            Else
                               Xnommed = "                         "
                            End If
                            If Xmotcon <> "" Then
                               Xmotcon = Mid(Xmotcon, 1, 40)
                               xcuenta = Len(Xmotcon)
                               xcuenta = xcuenta + 1
                               For xcuenta = xcuenta To 40
                                   Xmotcon = Xmotcon + " "
                               Next
                            Else
                               Xmotcon = "                                        "
                            End If
                            If XCat <> "" Then
                               XCat = Mid(XCat, 1, 6)
                               xcuenta = Len(XCat)
                               xcuenta = xcuenta + 1
                               For xcuenta = xcuenta To 6
                                   XCat = XCat + " "
                               Next
                            Else
                               XCat = "      "
                            End If
                            If Xdescol <> "" Then
                               Xdescol = Mid(Xdescol, 1, 9)
                               xcuenta = Len(Xdescol)
                               xcuenta = xcuenta + 1
                               For xcuenta = xcuenta To 9
                                   Xdescol = Xdescol + " "
                               Next
                            Else
                               Xdescol = "         "
                            End If
                            If data_inf.Recordset("codzon") = 1 Then
                               Xzona = "Z.COSTA"
                            Else
                               Xzona = "Z.NORTE"
                            End If
                            If data_inf.Recordset("totdem") >= "00:31" And data_inf.Recordset("totdem") <= "01:00" Then
                               Xtotal = Xtotal + 1
                                Print #1, CStr(data_inf.Recordset("fecha")) + " " + data_inf.Recordset("hora") + " " + Xnom + " " + Xedad _
                                ; " " + data_inf.Recordset("hor_llega") + " " + data_inf.Recordset("totdem") + " " + Xnommed _
                                + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama
                               Xtotgral = Xtotgral + 1
                            End If
                            data_inf.Recordset.MoveNext
                        End If
                     Loop
                     Print #1, "================================================="
                     Print #1, "TOTAL......:" + str(Xtotal)
                     Print #1, "================================================="
                     Xdemmas30 = Xtotal
                     Xtotal = 0
                     data_inf.Recordset.MoveFirst
                     Xcadllama = "DEMORAS EN LLAMADOS MAS DE 1 HORA HASTA 2 HORAS"
                     Print #1, ""
                     Print #1, Xcadllama
                     Print #1, "================================================="
                     Xcadllama = ""
                     Do While Not data_inf.Recordset.EOF
                        If Len(data_inf.Recordset("movilpas")) = 1 Then
                           Xnromov = "  " + str(data_inf.Recordset("movilpas"))
                        End If
                        If Len(data_inf.Recordset("movilpas")) = 2 Then
                           Xnromov = " " + str(data_inf.Recordset("movilpas"))
                        End If
                        If Len(data_inf.Recordset("movilpas")) = 3 Then
                           Xnromov = str(data_inf.Recordset("movilpas"))
                        End If
                        If Len(data_inf.Recordset("edad")) = 1 Then
                           Xedad = "  " + str(data_inf.Recordset("edad"))
                        End If
                        If Len(data_inf.Recordset("edad")) = 2 Then
                           Xedad = " " + str(data_inf.Recordset("edad"))
                        End If
                        If Len(data_inf.Recordset("edad")) = 3 Then
                           Xedad = str(data_inf.Recordset("edad"))
                        End If
                        Xmotcon = data_inf.Recordset("motcon")
                        XCat = data_inf.Recordset("categ")
                        Xdescol = data_inf.Recordset("descol")
                        Xnom = data_inf.Recordset("nombre")
                        If IsNull(data_inf.Recordset("nommed")) = False Then
                           Xnommed = data_inf.Recordset("nommed")
                        Else
                           Xnommed = ""
                        End If
                        If Xnom <> "" Then
                           Xnom = Mid(Xnom, 1, 50)
                           xcuenta = Len(Xnom)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 50
                               Xnom = Xnom + " "
                           Next
                        Else
                           Xnom = "                                                  "
                        End If
                        If Xnommed <> "" Then
                           Xnommed = Mid(Xnommed, 1, 25)
                           xcuenta = Len(Xnommed)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 25
                               Xnommed = Xnommed + " "
                           Next
                        Else
                           Xnommed = "                         "
                        End If
                        If Xmotcon <> "" Then
                           Xmotcon = Mid(Xmotcon, 1, 40)
                           xcuenta = Len(Xmotcon)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 40
                               Xmotcon = Xmotcon + " "
                           Next
                        Else
                           Xmotcon = "                                        "
                        End If
                        If XCat <> "" Then
                           XCat = Mid(XCat, 1, 6)
                           xcuenta = Len(XCat)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 6
                               XCat = XCat + " "
                           Next
                        Else
                           XCat = "      "
                        End If
                        If Xdescol <> "" Then
                           Xdescol = Mid(Xdescol, 1, 9)
                           xcuenta = Len(Xdescol)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 9
                               Xdescol = Xdescol + " "
                           Next
                        Else
                           Xdescol = "         "
                        End If
                        If data_inf.Recordset("codzon") = 1 Then
                           Xzona = "Z.COSTA"
                        Else
                           Xzona = "Z.NORTE"
                        End If
                        If data_inf.Recordset("totdem") > "01:00" And data_inf.Recordset("totdem") <= "02:00" Then
                           Xtotal = Xtotal + 1
                            Print #1, CStr(data_inf.Recordset("fecha")) + " " + data_inf.Recordset("hora") + " " + Xnom + " " + Xedad _
                            ; " " + data_inf.Recordset("hor_llega") + " " + data_inf.Recordset("totdem") + " " + Xnommed _
                            + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama
                           Xtotgral = Xtotgral + 1
                        End If
                        data_inf.Recordset.MoveNext
                     Loop
                     Print #1, "================================================="
                     Print #1, "TOTAL......:" + str(Xtotal)
                     Print #1, "================================================="
                     Xdemmas1 = Xtotal
                     Xtotal = 0
                     data_inf.Recordset.MoveFirst
                     Xcadllama = "DEMORAS EN LLAMADOS MAS DE 2 HORAS"
                     Print #1, ""
                     Print #1, Xcadllama
                     Print #1, "================================================="
                     Xcadllama = ""
                     Do While Not data_inf.Recordset.EOF
                        If Len(data_inf.Recordset("movilpas")) = 1 Then
                           Xnromov = "  " + str(data_inf.Recordset("movilpas"))
                        End If
                        If Len(data_inf.Recordset("movilpas")) = 2 Then
                           Xnromov = " " + str(data_inf.Recordset("movilpas"))
                        End If
                        If Len(data_inf.Recordset("movilpas")) = 3 Then
                           Xnromov = str(data_inf.Recordset("movilpas"))
                        End If
                        If Len(data_inf.Recordset("edad")) = 1 Then
                           Xedad = "  " + str(data_inf.Recordset("edad"))
                        End If
                        If Len(data_inf.Recordset("edad")) = 2 Then
                           Xedad = " " + str(data_inf.Recordset("edad"))
                        End If
                        If Len(data_inf.Recordset("edad")) = 3 Then
                           Xedad = str(data_inf.Recordset("edad"))
                        End If
                        Xmotcon = data_inf.Recordset("motcon")
                        XCat = data_inf.Recordset("categ")
                        Xdescol = data_inf.Recordset("descol")
                        Xnom = data_inf.Recordset("nombre")
                        If IsNull(data_inf.Recordset("nommed")) = False Then
                           Xnommed = data_inf.Recordset("nommed")
                        Else
                           Xnommed = ""
                        End If
                        If Xnom <> "" Then
                           Xnom = Mid(Xnom, 1, 50)
                           xcuenta = Len(Xnom)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 50
                               Xnom = Xnom + " "
                           Next
                        Else
                           Xnom = "                                                  "
                        End If
                        If Xnommed <> "" Then
                           Xnommed = Mid(Xnommed, 1, 25)
                           xcuenta = Len(Xnommed)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 25
                               Xnommed = Xnommed + " "
                           Next
                        Else
                           Xnommed = "                         "
                        End If
                        If Xmotcon <> "" Then
                           Xmotcon = Mid(Xmotcon, 1, 40)
                           xcuenta = Len(Xmotcon)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 40
                               Xmotcon = Xmotcon + " "
                           Next
                        Else
                           Xmotcon = "                                        "
                        End If
                        If XCat <> "" Then
                           XCat = Mid(XCat, 1, 6)
                           xcuenta = Len(XCat)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 6
                               XCat = XCat + " "
                           Next
                        Else
                           XCat = "      "
                        End If
                        If Xdescol <> "" Then
                           Xdescol = Mid(Xdescol, 1, 9)
                           xcuenta = Len(Xdescol)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 9
                               Xdescol = Xdescol + " "
                           Next
                        Else
                           Xdescol = "         "
                        End If
                        If data_inf.Recordset("codzon") = 1 Then
                           Xzona = "Z.COSTA"
                        Else
                           Xzona = "Z.NORTE"
                        End If
                        If data_inf.Recordset("totdem") > "02:00" Then
                           Xtotal = Xtotal + 1
                            Print #1, CStr(data_inf.Recordset("fecha")) + " " + data_inf.Recordset("hora") + " " + Xnom + " " + Xedad _
                            ; " " + data_inf.Recordset("hor_llega") + " " + data_inf.Recordset("totdem") + " " + Xnommed _
                            + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama
                           Xtotgral = Xtotgral + 1
                        End If
                        data_inf.Recordset.MoveNext
                     Loop
                     Print #1, "================================================="
                     Print #1, "TOTAL......:" + str(Xtotal)
                     Print #1, "================================================="
                     Xdemmas2 = Xtotal
                 Else
                 
                     Xtotal = 0
                     Xcadllama = "INFORME DEMORAS PROMEDIO DEL MEDICO EN DOMICILIO VERDES"
                     Print #1, Xcadllama
                     Print #1, "================================================="
                     Print #1, "FECHA/HORA REC.   N O M B R E S                                   EDAD  H.REA. DEMORA   MEDICO               MOVIL  MOTIVO CONSULTA                         CATEG.  COLOR     ZONA"
                     Print #1, "===================================================================================================================================================================================="
                     Xcadllama = ""
                     data_inf.RecordSource = "Select * from inflla order by codmed"
                     data_inf.Refresh
                     
                     data_inf.Recordset.MoveFirst
                     Xcoddelmed = data_inf.Recordset("codmed")
                     Do While Not data_inf.Recordset.EOF
                        If data_inf.Recordset("codmot") = "V" Or data_inf.Recordset("codmot") = "Z" Then
                            If IsNull(data_inf.Recordset("hor_llega")) = True Then
                               data_inf.Recordset.MoveNext
                            Else
                                If IsNull(data_inf.Recordset("hor_llega")) = False Then
                                   xhh = Val(Mid(data_inf.Recordset("hor_llega"), 1, 2))
                '                   If xhh = 0 Then
                '                      xhh = 24
                '                   End If
                                   xmm = Val(Mid(data_inf.Recordset("hor_llega"), 4, 2))
                                End If
                                If IsNull(data_inf.Recordset("hor_rea")) = False Then
                                   Xhhh = Val(Mid(data_inf.Recordset("hor_rea"), 1, 2))
                '                   If Xhhh = 0 Then
                '                      Xhhh = 24
                '                   End If
                                   Xmmh = Val(Mid(data_inf.Recordset("hor_rea"), 4, 2))
                                End If
                                xdemh = Xhhh - xhh
                                xdemm = Xmmh - xmm
                                If data_inf.Recordset("fecha") < data_inf.Recordset("fec_llega") Then
                                   If xdemh < 0 Then
                                      xdemh = Xhhh - xhh
                                      xdemh = xdemh + 24
                                   End If
                                Else
                                   If IsNull(data_inf.Recordset("fec_llega")) = True Then
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
                                   If xdemm < 0 Then
                                      xdemm = xdemm + 60
                                   End If
                                End If
                                data_inf.Recordset.Edit
                                If xdemh > 9 Then
                                   If xdemm > 9 Then
                                      data_inf.Recordset("totend") = Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                                   Else
                                      data_inf.Recordset("totend") = Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                                   End If
                                Else
                                   If xdemm > 9 Then
                                      If xdemh < 0 Then
                                         xdemh = 0
                                      End If
                                      data_inf.Recordset("totend") = "0" + Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                                   Else
                                      If xdemh < 0 Then
                                         xdemh = 0
                                      End If
                                      If xdemh < 0 Then
                                         xdemh = 0
                                      End If
                                      If xdemm < 0 Then
                                         xdemm = 0
                                      End If
                                      data_inf.Recordset("totend") = "0" + Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                                   End If
                                End If
                                If xdemh > 0 Then
                                   Xminhs = xdemh * 60
                                Else
                                   Xminhs = 0
                                End If
                                data_inf.Recordset("mes") = Xminhs + xdemm
                                
                                data_inf.Recordset.Update
                                If Len(data_inf.Recordset("movilpas")) = 1 Then
                                   Xnromov = "  " + str(data_inf.Recordset("movilpas"))
                                End If
                                If Len(data_inf.Recordset("movilpas")) = 2 Then
                                   Xnromov = " " + str(data_inf.Recordset("movilpas"))
                                End If
                                If Len(data_inf.Recordset("movilpas")) = 3 Then
                                   Xnromov = str(data_inf.Recordset("movilpas"))
                                End If
                                If Len(data_inf.Recordset("edad")) = 1 Then
                                   Xedad = "  " + str(data_inf.Recordset("edad"))
                                End If
                                If Len(data_inf.Recordset("edad")) = 2 Then
                                   Xedad = " " + str(data_inf.Recordset("edad"))
                                End If
                                If Len(data_inf.Recordset("edad")) = 3 Then
                                   Xedad = str(data_inf.Recordset("edad"))
                                End If
                                Xmotcon = data_inf.Recordset("motcon")
                                XCat = data_inf.Recordset("categ")
                                Xdescol = data_inf.Recordset("descol")
                                Xnom = data_inf.Recordset("nombre")
                                If IsNull(data_inf.Recordset("nommed")) = False Then
                                   Xnommed = data_inf.Recordset("nommed")
                                Else
                                   Xnommed = ""
                                End If
                                If Xnom <> "" Then
                                   Xnom = Mid(Xnom, 1, 50)
                                   xcuenta = Len(Xnom)
                                   xcuenta = xcuenta + 1
                                   For xcuenta = xcuenta To 50
                                       Xnom = Xnom + " "
                                   Next
                                Else
                                   Xnom = "                                                  "
                                End If
                                If Xnommed <> "" Then
                                   Xnommed = Mid(Xnommed, 1, 25)
                                   xcuenta = Len(Xnommed)
                                   xcuenta = xcuenta + 1
                                   For xcuenta = xcuenta To 25
                                       Xnommed = Xnommed + " "
                                   Next
                                Else
                                   Xnommed = "                         "
                                End If
                                If Xmotcon <> "" Then
                                   Xmotcon = Mid(Xmotcon, 1, 40)
                                   xcuenta = Len(Xmotcon)
                                   xcuenta = xcuenta + 1
                                   For xcuenta = xcuenta To 40
                                       Xmotcon = Xmotcon + " "
                                   Next
                                Else
                                   Xmotcon = "                                        "
                                End If
                                If XCat <> "" Then
                                   XCat = Mid(XCat, 1, 6)
                                   xcuenta = Len(XCat)
                                   xcuenta = xcuenta + 1
                                   For xcuenta = xcuenta To 6
                                       XCat = XCat + " "
                                   Next
                                Else
                                   XCat = "      "
                                End If
                                If Xdescol <> "" Then
                                   Xdescol = Mid(Xdescol, 1, 9)
                                   xcuenta = Len(Xdescol)
                                   xcuenta = xcuenta + 1
                                   For xcuenta = xcuenta To 9
                                       Xdescol = Xdescol + " "
                                   Next
                                Else
                                   Xdescol = "         "
                                End If
                                If data_inf.Recordset("codzon") = 1 Then
                                   Xzona = "Z.COSTA"
                                Else
                                   Xzona = "Z.NORTE"
                                End If
                                If data_inf.Recordset("totend") > "00:00" Then
                                   If data_inf.Recordset("codmed") = Xcoddelmed Then
                                        Xtotal = Xtotal + 1
                                         Print #1, CStr(data_inf.Recordset("fecha")) + " " + data_inf.Recordset("hor_llega") + " " + Xnom + " " + Xedad _
                                         ; " " + data_inf.Recordset("hor_rea") + " " + data_inf.Recordset("totend") + " " + Xnommed _
                                         + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama + " " + Trim(str(data_inf.Recordset("mes")))
                                         Xtotalmins = Xtotalmins + data_inf.Recordset("mes")
                                        Xtotporme = Xtotporme + 1
                                        Xtothsporme = Xtothsporme + data_inf.Recordset("mes")
                                        Xmedunavez = data_inf.Recordset("nommed")
                                   Else
                                         If Xtothsporme <= 0 Then
                                            Xelprommed = 0
                                            Xtotporme = 0
                                            Xtothsporme = 0
                                         Else
                                            Xelprommed = Xtothsporme / Xtotporme
                                         End If
                                         Print #1, "***PROMEDIO POR MEDICO....:  " & Trim(str(Format(Xelprommed, "Standard"))) & "  -------> " & Xmedunavez
                                         Xtotporme = 0
                                         Xtothsporme = 0
                                         Print #1, CStr(data_inf.Recordset("fecha")) + " " + data_inf.Recordset("hor_llega") + " " + Xnom + " " + Xedad _
                                         ; " " + data_inf.Recordset("hor_rea") + " " + data_inf.Recordset("totend") + " " + Xnommed _
                                         + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama + " " + Trim(str(data_inf.Recordset("mes")))
                                         Xtotalmins = Xtotalmins + data_inf.Recordset("mes")
                                        Xtotporme = Xtotporme + 1
                                        Xtothsporme = data_inf.Recordset("mes")
                                         
                                   End If
                                End If
                                Xcoddelmed = data_inf.Recordset("codmed")
                                data_inf.Recordset.MoveNext
                            End If
                        Else
                            data_inf.Recordset.MoveNext
                        End If
                     Loop
                     data_inf.Recordset.MovePrevious
                     If Xtothsporme <= 0 Then
                        Xelprommed = 0
                        Xtotporme = 0
                        Xtothsporme = 0
                     Else
                        Xelprommed = Xtothsporme / Xtotporme
                     End If
                     Print #1, "***PROMEDIO POR MEDICO....:  " & Trim(str(Format(Xelprommed, "Standard"))) & "  -------> " & Xmedunavez
                     Xtotporme = 0
                     Xtothsporme = 0
                     Xtotalmins = Xtotalmins + data_inf.Recordset("mes")
                     Xtotporme = Xtotporme + 1
                     Xtothsporme = data_inf.Recordset("mes")
                     data_inf.Recordset.MoveNext
                     Dim Xelprom As Double
                     Xelprom = Xtotalmins / Xtotal
                     Print #1, "================================================="
                     Print #1, "TOTAL......:" + str(Xtotal) & " MINUTOS: " & Trim(str(Xtotalmins)) & " PROMEDIO: " & Format(Xelprom, "Standard")
                     Print #1, "================================================="
                     Xtotal = 0
                     Print #1, ""
                 End If
                 
                 Close #1
                 If Check2.Value = 1 Then
                    Command4_Click
                 End If
                 MsgBox "Proceso Terminado..."
                 OLE1.SourceDoc = App.path & "\DEMORAS.txt"
                 OLE1.Action = 1
                 OLE1.DoVerb (-1)
             End If
         Else
             Command3.Visible = True
             Command3_Click
             Command3.Visible = False
         
         End If
      End If
   End If
End If
''Timer1.Enabled = True
frm_infdemoras.MousePointer = 0
Command1.Enabled = True
Command2.Enabled = True


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
    data_llama.Recordset.MoveFirst
    Do While Not data_llama.Recordset.EOF
       data_inf.Recordset.AddNew
       data_inf.Recordset("nro") = data_llama.Recordset("nro")
       data_inf.Recordset("fecha") = data_llama.Recordset("fecha")
       data_inf.Recordset("hora") = data_llama.Recordset("hora")
       data_inf.Recordset("usuario") = data_llama.Recordset("usuario")
       If data_llama.Recordset("matric") >= 999999999 Then
          data_inf.Recordset("matric") = 0
       Else
          data_inf.Recordset("matric") = data_llama.Recordset("matric")
       End If
       data_inf.Recordset("nombre") = data_llama.Recordset("nombre")
       data_inf.Recordset("edad") = data_llama.Recordset("edad")
       data_inf.Recordset("unied") = data_llama.Recordset("unied")
       data_inf.Recordset("categ") = data_llama.Recordset("categ")
       data_inf.Recordset("nomcat") = data_llama.Recordset("nomcat")
       If data_llama.Recordset("ci") >= 999999999 Then
          data_inf.Recordset("ci") = 0
       Else
          data_inf.Recordset("ci") = data_llama.Recordset("ci")
       End If
       data_inf.Recordset("direcc") = data_llama.Recordset("direcc")
       data_inf.Recordset("telef") = data_llama.Recordset("telef")
       data_inf.Recordset("codzon") = data_llama.Recordset("codzon")
       data_inf.Recordset("base") = data_llama.Recordset("base")
       data_inf.Recordset("referen") = data_llama.Recordset("referen")
       data_inf.Recordset("motcon") = data_llama.Recordset("motcon")
       data_inf.Recordset("obsmot") = data_llama.Recordset("obsmot")
       data_inf.Recordset("codmot") = data_llama.Recordset("codmot")
       data_inf.Recordset("descol") = data_llama.Recordset("descol")
       data_inf.Recordset("movilpas") = data_llama.Recordset("movilpas")
       data_inf.Recordset("pend") = data_llama.Recordset("pend")
       If IsNull(data_llama.Recordset("fec_rea")) = True Then
          data_inf.Recordset("fec_rea") = data_llama.Recordset("fecpas")
       Else
          data_inf.Recordset("fec_rea") = data_llama.Recordset("fec_rea")
       End If
       If IsNull(data_llama.Recordset("hor_rea")) = True Then
          data_inf.Recordset("hor_rea") = data_llama.Recordset("horpas")
       Else
          data_inf.Recordset("hor_rea") = data_llama.Recordset("hor_rea")
       End If
       data_inf.Recordset("fecpas") = data_llama.Recordset("fecpas")
       data_inf.Recordset("horpas") = data_llama.Recordset("horpas")
       data_inf.Recordset("fecsali") = data_llama.Recordset("fecsali")
       data_inf.Recordset("horsali") = data_llama.Recordset("horsali")
       If IsNull(data_llama.Recordset("fec_llega")) = True Then
          data_inf.Recordset("fec_llega") = data_llama.Recordset("fecpas")
       Else
          data_inf.Recordset("fec_llega") = data_llama.Recordset("fec_llega")
       End If
       If IsNull(data_llama.Recordset("hor_llega")) = True Then
          data_inf.Recordset("hor_llega") = data_llama.Recordset("horpas")
       Else
          data_inf.Recordset("hor_llega") = data_llama.Recordset("hor_llega")
       End If
       data_inf.Recordset("diag") = data_llama.Recordset("diag")
       data_inf.Recordset("colormot") = data_llama.Recordset("colormot")
       data_inf.Recordset("codmed") = data_llama.Recordset("codmed")
       data_inf.Recordset("obs") = data_llama.Recordset("obs")
       data_inf.Recordset("nommed") = data_llama.Recordset("nommed")
       data_inf.Recordset("trasla") = data_llama.Recordset("trasla")
       data_inf.Recordset("lugar") = data_llama.Recordset("lugar")
       data_inf.Recordset("hsald") = data_llama.Recordset("hsald")
       data_inf.Recordset("hllega") = data_llama.Recordset("hllega")
       data_inf.Recordset("hzona") = data_llama.Recordset("hzona")
       data_inf.Recordset("movil_rea") = data_llama.Recordset("movil_rea")
       data_inf.Recordset("totdem") = data_llama.Recordset("totdem")
       data_inf.Recordset("totend") = data_llama.Recordset("totend")
       data_inf.Recordset("cancela") = data_llama.Recordset("cancela")
       data_inf.Recordset.Update
       data_llama.Recordset.MoveNext
    Loop
    data_inf.RecordSource = "Select * from inflla"
    data_inf.Refresh
    If data_inf.Recordset.RecordCount > 0 Then
       If Check1.Value = 1 Then
          Dim MiBaseact As Database
          Dim Unasesact As Workspace
          Set Unasesact = Workspaces(0)
          Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
          MiBaseact.Execute "Delete * from inflla where cancela =" & 1

       End If
       data_inf.Refresh
    End If
    
    Xcadllama = "RESUMEN DE LLAMADOS DEMORAS MAS DE 30 MINUTOS"
    Open "c:\sappmys\sappmysql\DEMORAS.txt" For Output As #1
    Print #1, Xcadllama
    Print #1, "================================================="
'         Print #1, "FECHA/HORA REC.   N O M B R E S                                   EDAD  H.REA. DEMORA   MEDICO               MOVIL  MOTIVO CONSULTA                         CATEG.  COLOR     ZONA"
'         Print #1, "===================================================================================================================================================================================="
    Xcadllama = ""
    Do While Not data_inf.Recordset.EOF
       If IsNull(data_inf.Recordset("hor_llega")) = True Then
          data_inf.Recordset.MoveNext
       Else
           If IsNull(data_inf.Recordset("hora")) = False Then
              xhh = Val(Mid(data_inf.Recordset("hora"), 1, 2))
              xmm = Val(Mid(data_inf.Recordset("hora"), 4, 2))
           End If
           If IsNull(data_inf.Recordset("hor_llega")) = False Then
              Xhhh = Val(Mid(data_inf.Recordset("hor_llega"), 1, 2))
              Xmmh = Val(Mid(data_inf.Recordset("hor_llega"), 4, 2))
           End If
           xdemh = Xhhh - xhh
           xdemm = Xmmh - xmm
           If data_inf.Recordset("fecha") < data_inf.Recordset("fec_llega") Then
              xdemh = Xhhh - xhh
              xdemh = xdemh + 24
           End If
           If xdemh > 0 Then
              If xdemm < 0 Then
                 xdemm = xdemm + 60
                 xdemh = xdemh - 1
              End If
           Else
              If xdemm < 0 Then
                 xdemm = xdemm + 60
              End If
           End If
           data_inf.Recordset.Edit
           If xdemh > 9 Then
              If xdemm > 9 Then
                 data_inf.Recordset("totdem") = Trim(str(xdemh)) + ":" + Trim(str(xdemm))
              Else
                 data_inf.Recordset("totdem") = Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
              End If
           Else
              If xdemm > 9 Then
                 If xdemh < 0 Then
                    xdemh = 0
                 End If
                 data_inf.Recordset("totdem") = "0" + Trim(str(xdemh)) + ":" + Trim(str(xdemm))
              Else
                 If xdemh < 0 Then
                    xdemh = 0
                 End If
                 data_inf.Recordset("totdem") = "0" + Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
              End If
           End If
           data_inf.Recordset.Update
           If Len(data_inf.Recordset("movilpas")) = 1 Then
              Xnromov = "  " + str(data_inf.Recordset("movilpas"))
           End If
           If Len(data_inf.Recordset("movilpas")) = 2 Then
              Xnromov = " " + str(data_inf.Recordset("movilpas"))
           End If
           If Len(data_inf.Recordset("movilpas")) = 3 Then
              Xnromov = str(data_inf.Recordset("movilpas"))
           End If
           If Len(data_inf.Recordset("edad")) = 1 Then
              Xedad = "  " + str(data_inf.Recordset("edad"))
           End If
           If Len(data_inf.Recordset("edad")) = 2 Then
              Xedad = " " + str(data_inf.Recordset("edad"))
           End If
           If Len(data_inf.Recordset("edad")) = 3 Then
              Xedad = str(data_inf.Recordset("edad"))
           End If
           Xmotcon = data_inf.Recordset("motcon")
           XCat = data_inf.Recordset("categ")
           Xdescol = data_inf.Recordset("descol")
           Xnom = data_inf.Recordset("nombre")
           If IsNull(data_inf.Recordset("nommed")) = False Then
              Xnommed = data_inf.Recordset("nommed")
           Else
              Xnommed = ""
           End If
           If Xnom <> "" Then
              Xnom = Mid(Xnom, 1, 50)
              xcuenta = Len(Xnom)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 50
                  Xnom = Xnom + " "
              Next
           Else
              Xnom = "                                                  "
           End If
           If Xnommed <> "" Then
              Xnommed = Mid(Xnommed, 1, 25)
              xcuenta = Len(Xnommed)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 25
                  Xnommed = Xnommed + " "
              Next
           Else
              Xnommed = "                         "
           End If
           If Xmotcon <> "" Then
              Xmotcon = Mid(Xmotcon, 1, 40)
              xcuenta = Len(Xmotcon)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 40
                  Xmotcon = Xmotcon + " "
              Next
           Else
              Xmotcon = "                                        "
           End If
           If XCat <> "" Then
              XCat = Mid(XCat, 1, 6)
              xcuenta = Len(XCat)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 6
                  XCat = XCat + " "
              Next
           Else
              XCat = "      "
           End If
           If Xdescol <> "" Then
              Xdescol = Mid(Xdescol, 1, 9)
              xcuenta = Len(Xdescol)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 9
                  Xdescol = Xdescol + " "
              Next
           Else
              Xdescol = "         "
           End If
           If data_inf.Recordset("codzon") = 1 Then
              Xzona = "Z.COSTA"
           Else
              Xzona = "Z.NORTE"
           End If
           If data_inf.Recordset("totdem") >= "00:00" And data_inf.Recordset("totdem") <= "00:30" Then
              Xtotal = Xtotal + 1
'                    Print #1, CStr(data_inf.Recordset("fecha")) + " " + data_inf.Recordset("hora") + " " + Xnom + " " + Xedad _
'                    ; " " + data_inf.Recordset("hor_llega") + " " + data_inf.Recordset("totdem") + " " + Xnommed _
'                    + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama
              Xtotgral = Xtotgral + 1
           End If
           data_inf.Recordset.MoveNext
       End If
    Loop
'         Print #1, "================================================="
'         Print #1, "TOTAL......:" + Str(Xtotal)
'         Print #1, "================================================="
'         Xhasta30 = Xtotal
    Xtotal = 0
    data_inf.Recordset.MoveFirst
    Xcadllama = "DEMORAS EN LLAMADOS MAS DE 30 MINUTOS A 1 HORA"
    Print #1, "================================================="
    Print #1, Xcadllama
    Print #1, "================================================="
    Xcadllama = ""
    Do While Not data_inf.Recordset.EOF
       If IsNull(data_inf.Recordset("hor_llega")) = True Then
          data_inf.Recordset.MoveNext
       Else
           If IsNull(data_inf.Recordset("hora")) = False Then
              xhh = Val(Mid(data_inf.Recordset("hora"), 1, 2))
              xmm = Val(Mid(data_inf.Recordset("hora"), 4, 2))
           End If
           If IsNull(data_inf.Recordset("hor_llega")) = False Then
              Xhhh = Val(Mid(data_inf.Recordset("hor_llega"), 1, 2))
              Xmmh = Val(Mid(data_inf.Recordset("hor_llega"), 4, 2))
           End If
           xdemh = Xhhh - xhh
           xdemm = Xmmh - xmm
           If data_inf.Recordset("fecha") < data_inf.Recordset("fec_llega") Then
              xdemh = Xhhh - xhh
              xdemh = xdemh + 24
           End If
           If xdemh > 0 Then
              If xdemm < 0 Then
                 xdemm = xdemm + 60
                 xdemh = xdemh - 1
              End If
           Else
              If xdemm < 0 Then
                 xdemm = xdemm + 60
              End If
           End If
           data_inf.Recordset.Edit
           If xdemh > 9 Then
              If xdemm > 9 Then
                 data_inf.Recordset("totdem") = Trim(str(xdemh)) + ":" + Trim(str(xdemm))
              Else
                 data_inf.Recordset("totdem") = Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
              End If
           Else
              If xdemm > 9 Then
                 If xdemh < 0 Then
                    xdemh = 0
                 End If
                 data_inf.Recordset("totdem") = "0" + Trim(str(xdemh)) + ":" + Trim(str(xdemm))
              Else
                 If xdemh < 0 Then
                    xdemh = 0
                 End If
                 data_inf.Recordset("totdem") = "0" + Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
              End If
           End If
           data_inf.Recordset.Update
           If Len(data_inf.Recordset("movilpas")) = 1 Then
              Xnromov = "  " + str(data_inf.Recordset("movilpas"))
           End If
           If Len(data_inf.Recordset("movilpas")) = 2 Then
              Xnromov = " " + str(data_inf.Recordset("movilpas"))
           End If
           If Len(data_inf.Recordset("movilpas")) = 3 Then
              Xnromov = str(data_inf.Recordset("movilpas"))
           End If
           If Len(data_inf.Recordset("edad")) = 1 Then
              Xedad = "  " + str(data_inf.Recordset("edad"))
           End If
           If Len(data_inf.Recordset("edad")) = 2 Then
              Xedad = " " + str(data_inf.Recordset("edad"))
           End If
           If Len(data_inf.Recordset("edad")) = 3 Then
              Xedad = str(data_inf.Recordset("edad"))
           End If
           Xmotcon = data_inf.Recordset("motcon")
           XCat = data_inf.Recordset("categ")
           Xdescol = data_inf.Recordset("descol")
           Xnom = data_inf.Recordset("nombre")
           If IsNull(data_inf.Recordset("nommed")) = False Then
              Xnommed = data_inf.Recordset("nommed")
           Else
              Xnommed = ""
           End If
           If Xnom <> "" Then
              Xnom = Mid(Xnom, 1, 50)
              xcuenta = Len(Xnom)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 50
                  Xnom = Xnom + " "
              Next
           Else
              Xnom = "                                                  "
           End If
           If Xnommed <> "" Then
              Xnommed = Mid(Xnommed, 1, 25)
              xcuenta = Len(Xnommed)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 25
                  Xnommed = Xnommed + " "
              Next
           Else
              Xnommed = "                         "
           End If
           If Xmotcon <> "" Then
              Xmotcon = Mid(Xmotcon, 1, 40)
              xcuenta = Len(Xmotcon)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 40
                  Xmotcon = Xmotcon + " "
              Next
           Else
              Xmotcon = "                                        "
           End If
           If XCat <> "" Then
              XCat = Mid(XCat, 1, 6)
              xcuenta = Len(XCat)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 6
                  XCat = XCat + " "
              Next
           Else
              XCat = "      "
           End If
           If Xdescol <> "" Then
              Xdescol = Mid(Xdescol, 1, 9)
              xcuenta = Len(Xdescol)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 9
                  Xdescol = Xdescol + " "
              Next
           Else
              Xdescol = "         "
           End If
           If data_inf.Recordset("codzon") = 1 Then
              Xzona = "Z.COSTA"
           Else
              Xzona = "Z.NORTE"
           End If
           If data_inf.Recordset("totdem") >= "00:31" And data_inf.Recordset("totdem") <= "01:00" Then
              Xtotal = Xtotal + 1
'               Print #1, CStr(data_inf.Recordset("fecha")) + " " + data_inf.Recordset("hora") + " " + Xnom + " " + Xedad _
'               ; " " + data_inf.Recordset("hor_llega") + " " + data_inf.Recordset("totdem") + " " + Xnommed _
'               + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama
              Xtotgral = Xtotgral + 1
           End If
           data_inf.Recordset.MoveNext
       End If
    Loop
    Print #1, "================================================="
    Print #1, "TOTAL......:" + str(Xtotal)
    Print #1, "================================================="
    Xdemmas30 = Xtotal
    Xtotal = 0
    data_inf.Recordset.MoveFirst
    
    Xcadllama = "DEMORAS EN LLAMADOS MAS DE 1 HORA HASTA 2 HORAS"
    Print #1, ""
    Print #1, "================================================="
    Print #1, Xcadllama
    Print #1, "================================================="
    Xcadllama = ""
    Do While Not data_inf.Recordset.EOF
       If Len(data_inf.Recordset("movilpas")) = 1 Then
          Xnromov = "  " + str(data_inf.Recordset("movilpas"))
       End If
       If Len(data_inf.Recordset("movilpas")) = 2 Then
          Xnromov = " " + str(data_inf.Recordset("movilpas"))
       End If
       If Len(data_inf.Recordset("movilpas")) = 3 Then
          Xnromov = str(data_inf.Recordset("movilpas"))
       End If
       If Len(data_inf.Recordset("edad")) = 1 Then
          Xedad = "  " + str(data_inf.Recordset("edad"))
       End If
       If Len(data_inf.Recordset("edad")) = 2 Then
          Xedad = " " + str(data_inf.Recordset("edad"))
       End If
       If Len(data_inf.Recordset("edad")) = 3 Then
          Xedad = str(data_inf.Recordset("edad"))
       End If
       Xmotcon = data_inf.Recordset("motcon")
       XCat = data_inf.Recordset("categ")
       Xdescol = data_inf.Recordset("descol")
       Xnom = data_inf.Recordset("nombre")
       If IsNull(data_inf.Recordset("nommed")) = False Then
          Xnommed = data_inf.Recordset("nommed")
       Else
          Xnommed = ""
       End If
       If Xnom <> "" Then
          Xnom = Mid(Xnom, 1, 50)
          xcuenta = Len(Xnom)
          xcuenta = xcuenta + 1
          For xcuenta = xcuenta To 50
              Xnom = Xnom + " "
          Next
       Else
          Xnom = "                                                  "
       End If
       If Xnommed <> "" Then
          Xnommed = Mid(Xnommed, 1, 25)
          xcuenta = Len(Xnommed)
          xcuenta = xcuenta + 1
          For xcuenta = xcuenta To 25
              Xnommed = Xnommed + " "
          Next
       Else
          Xnommed = "                         "
       End If
       If Xmotcon <> "" Then
          Xmotcon = Mid(Xmotcon, 1, 40)
          xcuenta = Len(Xmotcon)
          xcuenta = xcuenta + 1
          For xcuenta = xcuenta To 40
              Xmotcon = Xmotcon + " "
          Next
       Else
          Xmotcon = "                                        "
       End If
       If XCat <> "" Then
          XCat = Mid(XCat, 1, 6)
          xcuenta = Len(XCat)
          xcuenta = xcuenta + 1
          For xcuenta = xcuenta To 6
              XCat = XCat + " "
          Next
       Else
          XCat = "      "
       End If
       If Xdescol <> "" Then
          Xdescol = Mid(Xdescol, 1, 9)
          xcuenta = Len(Xdescol)
          xcuenta = xcuenta + 1
          For xcuenta = xcuenta To 9
              Xdescol = Xdescol + " "
          Next
       Else
          Xdescol = "         "
       End If
       If data_inf.Recordset("codzon") = 1 Then
          Xzona = "Z.COSTA"
       Else
          Xzona = "Z.NORTE"
       End If
       If data_inf.Recordset("totdem") > "01:00" And data_inf.Recordset("totdem") <= "02:00" Then
          Xtotal = Xtotal + 1
'           Print #1, CStr(data_inf.Recordset("fecha")) + " " + data_inf.Recordset("hora") + " " + Xnom + " " + Xedad _
'           ; " " + data_inf.Recordset("hor_llega") + " " + data_inf.Recordset("totdem") + " " + Xnommed _
'           + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama
          Xtotgral = Xtotgral + 1
       End If
       data_inf.Recordset.MoveNext
    Loop
    Print #1, "================================================="
    Print #1, "TOTAL......:" + str(Xtotal)
    Print #1, "================================================="
    Xdemmas1 = Xtotal
    Xtotal = 0
    data_inf.Recordset.MoveFirst
    Xcadllama = "DEMORAS EN LLAMADOS MAS DE 2 HORAS"
    Print #1, ""
    Print #1, "================================================="
    Print #1, Xcadllama
    Print #1, "================================================="
    Xcadllama = ""
    Do While Not data_inf.Recordset.EOF
       If Len(data_inf.Recordset("movilpas")) = 1 Then
          Xnromov = "  " + str(data_inf.Recordset("movilpas"))
       End If
       If Len(data_inf.Recordset("movilpas")) = 2 Then
          Xnromov = " " + str(data_inf.Recordset("movilpas"))
       End If
       If Len(data_inf.Recordset("movilpas")) = 3 Then
          Xnromov = str(data_inf.Recordset("movilpas"))
       End If
       If Len(data_inf.Recordset("edad")) = 1 Then
          Xedad = "  " + str(data_inf.Recordset("edad"))
       End If
       If Len(data_inf.Recordset("edad")) = 2 Then
          Xedad = " " + str(data_inf.Recordset("edad"))
       End If
       If Len(data_inf.Recordset("edad")) = 3 Then
          Xedad = str(data_inf.Recordset("edad"))
       End If
       Xmotcon = data_inf.Recordset("motcon")
       XCat = data_inf.Recordset("categ")
       Xdescol = data_inf.Recordset("descol")
       Xnom = data_inf.Recordset("nombre")
       If IsNull(data_inf.Recordset("nommed")) = False Then
          Xnommed = data_inf.Recordset("nommed")
       Else
          Xnommed = ""
       End If
       If Xnom <> "" Then
          Xnom = Mid(Xnom, 1, 50)
          xcuenta = Len(Xnom)
          xcuenta = xcuenta + 1
          For xcuenta = xcuenta To 50
              Xnom = Xnom + " "
          Next
       Else
          Xnom = "                                                  "
       End If
       If Xnommed <> "" Then
          Xnommed = Mid(Xnommed, 1, 25)
          xcuenta = Len(Xnommed)
          xcuenta = xcuenta + 1
          For xcuenta = xcuenta To 25
              Xnommed = Xnommed + " "
          Next
       Else
          Xnommed = "                         "
       End If
       If Xmotcon <> "" Then
          Xmotcon = Mid(Xmotcon, 1, 40)
          xcuenta = Len(Xmotcon)
          xcuenta = xcuenta + 1
          For xcuenta = xcuenta To 40
              Xmotcon = Xmotcon + " "
          Next
       Else
          Xmotcon = "                                        "
       End If
       If XCat <> "" Then
          XCat = Mid(XCat, 1, 6)
          xcuenta = Len(XCat)
          xcuenta = xcuenta + 1
          For xcuenta = xcuenta To 6
              XCat = XCat + " "
          Next
       Else
          XCat = "      "
       End If
       If Xdescol <> "" Then
          Xdescol = Mid(Xdescol, 1, 9)
          xcuenta = Len(Xdescol)
          xcuenta = xcuenta + 1
          For xcuenta = xcuenta To 9
              Xdescol = Xdescol + " "
          Next
       Else
          Xdescol = "         "
       End If
       If data_inf.Recordset("codzon") = 1 Then
          Xzona = "Z.COSTA"
       Else
          Xzona = "Z.NORTE"
       End If
       If data_inf.Recordset("totdem") > "02:00" Then
          Xtotal = Xtotal + 1
'           Print #1, CStr(data_inf.Recordset("fecha")) + " " + data_inf.Recordset("hora") + " " + Xnom + " " + Xedad _
'           ; " " + data_inf.Recordset("hor_llega") + " " + data_inf.Recordset("totdem") + " " + Xnommed _
'           + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama
          Xtotgral = Xtotgral + 1
       End If
       data_inf.Recordset.MoveNext
    Loop
    Print #1, "================================================="
    Print #1, "TOTAL......:" + str(Xtotal)
    Print #1, "================================================="
    Xdemmas2 = Xtotal
    Xtotal = 0
    Print #1, ""
    Xcadllama = "TIEMPO EN DOMICILIO DEL MEDICO MAS DE 30 MINUTOS"
    Print #1, "==================================================="
    Print #1, Xcadllama
    Print #1, "==================================================="
    Xcadllama = ""
    data_inf.Recordset.MoveFirst
    Do While Not data_inf.Recordset.EOF
       If IsNull(data_inf.Recordset("hor_llega")) = True Then
          data_inf.Recordset.MoveNext
       Else
           If IsNull(data_inf.Recordset("hor_llega")) = False Then
              xhh = Val(Mid(data_inf.Recordset("hor_llega"), 1, 2))
'                   If xhh = 0 Then
'                      xhh = 24
'                   End If
              xmm = Val(Mid(data_inf.Recordset("hor_llega"), 4, 2))
           End If
           If IsNull(data_inf.Recordset("hor_rea")) = False Then
              Xhhh = Val(Mid(data_inf.Recordset("hor_rea"), 1, 2))
'                   If Xhhh = 0 Then
'                      Xhhh = 24
'                   End If
              Xmmh = Val(Mid(data_inf.Recordset("hor_rea"), 4, 2))
           End If
           xdemh = Xhhh - xhh
           xdemm = Xmmh - xmm
           If data_inf.Recordset("fecha") < data_inf.Recordset("fec_llega") Then
              If xdemh < 0 Then
                 xdemh = Xhhh - xhh
                 xdemh = xdemh + 24
              End If
           Else
              If IsNull(data_inf.Recordset("fec_llega")) = True Then
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
              If xdemm < 0 Then
                 xdemm = xdemm + 60
              End If
           End If
           data_inf.Recordset.Edit
           If xdemh > 9 Then
              If xdemm > 9 Then
                 data_inf.Recordset("totend") = Trim(str(xdemh)) + ":" + Trim(str(xdemm))
              Else
                 data_inf.Recordset("totend") = Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
              End If
           Else
              If xdemm > 9 Then
                 data_inf.Recordset("totend") = "0" + Trim(str(xdemh)) + ":" + Trim(str(xdemm))
              Else
                 data_inf.Recordset("totend") = "0" + Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
              End If
           End If
           data_inf.Recordset.Update
           If Len(data_inf.Recordset("movilpas")) = 1 Then
              Xnromov = "  " + str(data_inf.Recordset("movilpas"))
           End If
           If Len(data_inf.Recordset("movilpas")) = 2 Then
              Xnromov = " " + str(data_inf.Recordset("movilpas"))
           End If
           If Len(data_inf.Recordset("movilpas")) = 3 Then
              Xnromov = str(data_inf.Recordset("movilpas"))
           End If
           If Len(data_inf.Recordset("edad")) = 1 Then
              Xedad = "  " + str(data_inf.Recordset("edad"))
           End If
           If Len(data_inf.Recordset("edad")) = 2 Then
              Xedad = " " + str(data_inf.Recordset("edad"))
           End If
           If Len(data_inf.Recordset("edad")) = 3 Then
              Xedad = str(data_inf.Recordset("edad"))
           End If
           Xmotcon = data_inf.Recordset("motcon")
           XCat = data_inf.Recordset("categ")
           Xdescol = data_inf.Recordset("descol")
           Xnom = data_inf.Recordset("nombre")
           If IsNull(data_inf.Recordset("nommed")) = False Then
              Xnommed = data_inf.Recordset("nommed")
           Else
              Xnommed = ""
           End If
           If Xnom <> "" Then
              Xnom = Mid(Xnom, 1, 50)
              xcuenta = Len(Xnom)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 50
                  Xnom = Xnom + " "
              Next
           Else
              Xnom = "                                                  "
           End If
           If Xnommed <> "" Then
              Xnommed = Mid(Xnommed, 1, 25)
              xcuenta = Len(Xnommed)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 25
                  Xnommed = Xnommed + " "
              Next
           Else
              Xnommed = "                         "
           End If
           If Xmotcon <> "" Then
              Xmotcon = Mid(Xmotcon, 1, 40)
              xcuenta = Len(Xmotcon)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 40
                  Xmotcon = Xmotcon + " "
              Next
           Else
              Xmotcon = "                                        "
           End If
           If XCat <> "" Then
              XCat = Mid(XCat, 1, 6)
              xcuenta = Len(XCat)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 6
                  XCat = XCat + " "
              Next
           Else
              XCat = "      "
           End If
           If Xdescol <> "" Then
              Xdescol = Mid(Xdescol, 1, 9)
              xcuenta = Len(Xdescol)
              xcuenta = xcuenta + 1
              For xcuenta = xcuenta To 9
                  Xdescol = Xdescol + " "
              Next
           Else
              Xdescol = "         "
           End If
           If data_inf.Recordset("codzon") = 1 Then
              Xzona = "Z.COSTA"
           Else
              Xzona = "Z.NORTE"
           End If
           If data_inf.Recordset("totend") > "00:30" Then
              Xtotal = Xtotal + 1
'               Print #1, CStr(data_inf.Recordset("fecha")) + " " + data_inf.Recordset("hor_llega") + " " + Xnom + " " + Xedad _
'               ; " " + data_inf.Recordset("hor_rea") + " " + data_inf.Recordset("totend") + " " + Xnommed _
'               + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama
           End If
           data_inf.Recordset.MoveNext
       End If
    Loop
    Print #1, "================================================="
    Print #1, "TOTAL......:" + str(Xtotal)
    Print #1, "================================================="
    Xtotal = 0
    Print #1, ""
    Print #1, "======================================================="
    Print #1, "TOTAL GENERAL DE LLAMADOS REALIZADOS EN LA FECHA......:" + str(Xtotgral)
    Print #1, "======================================================="
    Print #1, ""
''             Print #1, "======================================================="
''             Print #1, "TOTAL LLAMADOS CON DEMORAS HASTA 30 min...............:" + Str(Xhasta30)
''             Print #1, "======================================================="
''             Print #1, ""
    Print #1, "======================================================="
    Print #1, "TOTAL LLAMADOS CON DEMORAS MAS 30 MIN.................:" + str(Xdemmas30)
    Print #1, "======================================================="
    Print #1, ""
    Print #1, "======================================================="
    Print #1, "TOTAL LLAMADOS CON DEMORAS MAS 1 hora.................:" + str(Xdemmas1)
    Print #1, "======================================================="
    Print #1, ""
    Print #1, "======================================================="
    Print #1, "TOTAL LLAMADOS CON DEMORAS MAS 2 horas................:" + str(Xdemmas2)
    Print #1, "======================================================="
    
    Close #1

frm_infdemoras.MousePointer = 0
MsgBox "Proceso Terminado..."
OLE1.Action = 1
OLE1.DoVerb (-1)


End Sub

Private Sub Command4_Click()
Dim Xobjexel2 As Excel.Application
Dim Xlibexel2 As Excel.Workbook
Dim Xarchexel2 As New Excel.Worksheet

Dim Xarchtex2 As String
Dim Xlin, XCol, Xtotglla, Xtotgllagt, Xpromed, Xtotgraldos As Double
Dim Xdiass, Xhastaqdia As Long
Dim Xfecontrol As Date
Xdiass = 1
If Month(md.Text) = 2 Then
   If Year(md.Text) = 2012 Then
      Xhastaqdia = 28
   Else
      Xhastaqdia = 27
   End If
Else
   Xhastaqdia = 29
End If

If Check2.Value = 1 Then
   Set Xobjexel2 = New Excel.Application
   Set Xlibexel2 = Xobjexel2.Workbooks.Add
   Set Xarchexel2 = Xlibexel2.Worksheets.Add
'   Xlibexel2.SaveAs ("C:\planillas\analisis" & Trim(Str(Month(mfd.Text))) & Trim(Str(Year(mfd.Text))) & ".xls")
   Xlibexel2.SaveAs ("C:\planillas\analisis.xls")
   
   Xarchtex2 = "C:\planillas\analisis.xls"
End If

Xlin = 1
XCol = 1
Xarchexel2.Name = "Analisis"
Xarchexel2.Cells(Xlin, XCol) = "SAPP S.A."
Xlin = Xlin + 2
XCol = 1
Xarchexel2.Range("A1", "C3").Font.Size = 16
Xarchexel2.Cells(Xlin, XCol) = "MES: " & Month(md.Text) & "/" & Year(md.Text)
'Xarchexel.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 120)
Xarchexel2.Range("A3").Interior.color = RGB(115, 120, 0)

XCol = 1
Xlin = Xlin + 1
'Xnrocan = Xnrocan + Xlin
Xarchexel2.Range("B4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("B4", "AN" & Trim(str(15))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(15))).Borders(xlInsideVertical).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(15))).Borders(xlEdgeBottom).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(15))).Borders(xlEdgeTop).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(15))).Borders(xlEdgeLeft).LineStyle = xlContinuous
Xarchexel2.Range("B4", "AN" & Trim(str(15))).Borders(xlEdgeRight).LineStyle = xlContinuous
Xarchexel2.Range("B4" & Trim(str(Xlin)), "AN" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
Xarchexel2.Range("A4" & Trim(str(Xlin))).ColumnWidth = 45
Xarchexel2.Range("B4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("C4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("D4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("E4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("F4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("G4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("H4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("I4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("J4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("K4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("L4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("M4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("N4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("O4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("P4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("Q4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("R4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("S4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("T4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("U4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("V4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("W4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("X4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("Y4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("Z4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AA4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AB4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AC4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AD4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AE4" & Trim(str(Xlin))).ColumnWidth = 4
Xarchexel2.Range("AF4" & Trim(str(Xlin))).ColumnWidth = 4

XCol = 2
Do While Xdiass <= 31
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xdiass))
   Xdiass = Xdiass + 1
   XCol = XCol + 1
Loop
'XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = "TOTAL"
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = "Prom"
XCol = XCol + 1
Xarchexel2.Cells(Xlin, XCol) = "%"
Xlin = Xlin + 1
XCol = 1
If t_mov.Text <> "" Then
   data_llama.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codzon in(1,2,3,5) and movilpas =" & t_mov.Text & " and codmed <>" & 959 & " and cancela is null order by fecha"
   data_llama.Refresh
Else
   data_llama.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codzon in(1,2,3,5) and codmed <>" & 959 & " and cancela is null order by fecha"
   data_llama.Refresh
End If
If data_llama.Recordset.RecordCount > 0 Then
   data_llama.Recordset.MoveLast
End If
Xtotgraldos = data_llama.Recordset.RecordCount
data_llama.Recordset.MoveFirst
Xlin = 9
Xarchexel2.Cells(Xlin, XCol) = "TOTAL DE LLAMADOS T.DOM >30 MIN."
'If data_inf.Recordset("totend") > "00:30" Then
data_inf.RecordSource = "select * from inflla where totend >'" & "00:30" & "' order by fecha"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   XCol = 2
   Xtotglla = 0
   Xtotgllagt = 0
   Xfecontrol = data_inf.Recordset("fecha")
   Do While Not data_inf.Recordset.EOF
      If data_inf.Recordset("fecha") = Xfecontrol Then
         Xtotglla = Xtotglla + 1
         Xtotgllagt = Xtotgllagt + 1
      Else
         Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
         Xtotglla = 1
         Xtotgllagt = Xtotgllagt + 1
         XCol = XCol + 1
      End If
      Xfecontrol = data_inf.Recordset("fecha")
      data_inf.Recordset.MoveNext
   Loop
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
   Xtotglla = 0
   XCol = 33
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotgllagt))
   XCol = 34
   Xpromed = Xtotgllagt / Xdiass
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
   XCol = 35
   Xpromed = Xtotgllagt * 100 / Xtotgraldos
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
End If

Xlin = Xlin + 1
XCol = 1
Xarchexel2.Cells(Xlin, XCol) = "TOTAL DE LLAMADOS CON DEMORAS >30MIN."
'If data_inf.Recordset("totend") > "00:30" Then
data_inf.RecordSource = "select * from inflla where totdem >'" & "00:30" & "' and totdem <='" & "01:00" & "' order by fecha"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   XCol = 2
   Xtotglla = 0
   Xtotgllagt = 0
   Xfecontrol = data_inf.Recordset("fecha")
   Do While Not data_inf.Recordset.EOF
      If data_inf.Recordset("fecha") = Xfecontrol Then
         Xtotglla = Xtotglla + 1
         Xtotgllagt = Xtotgllagt + 1
      Else
         Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
         Xtotglla = 1
         Xtotgllagt = Xtotgllagt + 1
         XCol = XCol + 1
      End If
      Xfecontrol = data_inf.Recordset("fecha")
      data_inf.Recordset.MoveNext
   Loop
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
   Xtotglla = 0
   XCol = 33
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotgllagt))
   XCol = 34
   Xpromed = Xtotgllagt / Xdiass
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
   XCol = 35
   Xpromed = Xtotgllagt * 100 / Xtotgraldos
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
End If

Xlin = Xlin + 1
XCol = 1
Xarchexel2.Cells(Xlin, XCol) = "TOTAL DE LLAMADOS CON DEMORAS >1 HORA"
data_inf.RecordSource = "select * from inflla where totdem >'" & "01:00" & "' and totdem <='" & "02:00" & "' order by fecha"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   XCol = 2
   Xtotglla = 0
   Xtotgllagt = 0
   Xfecontrol = data_inf.Recordset("fecha")
   Do While Not data_inf.Recordset.EOF
      If data_inf.Recordset("fecha") = Xfecontrol Then
         Xtotglla = Xtotglla + 1
         Xtotgllagt = Xtotgllagt + 1
      Else
         Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
         Xtotglla = 1
         Xtotgllagt = Xtotgllagt + 1
         XCol = XCol + 1
      End If
      Xfecontrol = data_inf.Recordset("fecha")
      data_inf.Recordset.MoveNext
   Loop
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
   Xtotglla = 0
   XCol = 33
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotgllagt))
   XCol = 34
   Xpromed = Xtotgllagt / Xdiass
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
   XCol = 35
   Xpromed = Xtotgllagt * 100 / Xtotgraldos
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
End If

Xlin = Xlin + 1
XCol = 1
Xarchexel2.Cells(Xlin, XCol) = "TOTAL DE LLAMADOS CON DEMORAS >2 HORAS"
data_inf.RecordSource = "select * from inflla where totdem >'" & "02:00" & "' order by fecha"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   XCol = 2
   Xtotglla = 0
   Xtotgllagt = 0
   Xfecontrol = data_inf.Recordset("fecha")
   Do While Not data_inf.Recordset.EOF
      If data_inf.Recordset("fecha") = Xfecontrol Then
         Xtotglla = Xtotglla + 1
         Xtotgllagt = Xtotgllagt + 1
      Else
         Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
         Xtotglla = 1
         Xtotgllagt = Xtotgllagt + 1
         XCol = XCol + 1
      End If
      Xfecontrol = data_inf.Recordset("fecha")
      data_inf.Recordset.MoveNext
   Loop
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
   Xtotglla = 0
   XCol = 33
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotgllagt))
   XCol = 34
   Xpromed = Xtotgllagt / Xdiass
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
   XCol = 35
   Xpromed = Xtotgllagt * 100 / Xtotgraldos
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
End If
Dim Xarmofec As String
Dim Xeldianum As Long
Dim Xfecreal As Date
Xeldianum = 1
Xarmofec = Trim(str(Xeldianum)) & "/" & Month(md.Text) & "/" & Year(md.Text)
Xfecreal = CDate(Xarmofec)

Xlin = Xlin + 1
XCol = 1
Xarchexel2.Cells(Xlin, XCol) = "TOTAL DEMORAS >2 HORAS SIN TALA Y CERTIF."
data_inf.RecordSource = "select * from inflla where totdem >'" & "02:00" & "' and codzon in (1,2,3,5) order by fecha"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   XCol = 2
   Xtotglla = 0
   Xtotgllagt = 0
'   Xfecontrol = data_inf.Recordset("fecha")
   Do While Xeldianum <= Xhastaqdia
      data_inf.RecordSource = "select * from inflla where fecha =#" & Format(Xfecreal, "yyyy/mm/dd") & "# and totdem >'" & "02:00" & "' and codzon in (1,2) order by fecha"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveLast
         data_inf.Recordset.MoveFirst
         Xtotglla = data_inf.Recordset.RecordCount
         Xtotgllagt = Xtotgllagt + Xtotglla
         Do While Not data_inf.Recordset.EOF
            If data_inf.Recordset("categ") = "UDEMM" Or _
               data_inf.Recordset("categ") = "CERSEM" Then
               Xtotglla = Xtotglla - 1
               Xtotgllagt = Xtotgllagt - 1
            End If
            data_inf.Recordset.MoveNext
         Loop
         Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
         Xtotglla = 0
         XCol = XCol + 1
      Else
         Xtotglla = 0
         Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
         Xtotglla = 0
         XCol = XCol + 1
      End If
      Xeldianum = Xeldianum + 1
      Xarmofec = Trim(str(Xeldianum)) & "/" & Month(md.Text) & "/" & Year(md.Text)
      Xfecreal = CDate(Xarmofec)
   Loop
   If Xeldianum = 31 Then
      If Month(md.Text) = 1 Or _
         Month(md.Text) = 3 Or _
         Month(md.Text) = 5 Or _
         Month(md.Text) = 7 Or _
         Month(md.Text) = 8 Or _
         Month(md.Text) = 10 Or _
         Month(md.Text) = 12 Then
         Xarmofec = Trim(str(Xeldianum)) & "/" & Month(md.Text) & "/" & Year(md.Text)
         Xfecreal = CDate(Xarmofec)
         data_inf.RecordSource = "select * from inflla where fecha =#" & Format(Xfecreal, "yyyy/mm/dd") & "# and totdem >'" & "02:00" & "' and codzon in (1,2) order by fecha"
         data_inf.Refresh
         If data_inf.Recordset.RecordCount > 0 Then
            data_inf.Recordset.MoveLast
            data_inf.Recordset.MoveFirst
            Xtotglla = data_inf.Recordset.RecordCount
            Xtotgllagt = Xtotgllagt + Xtotglla
            Do While Not data_inf.Recordset.EOF
               If data_inf.Recordset("categ") = "UDEMM" Or _
                  data_inf.Recordset("categ") = "CERSEM" Then
                  Xtotglla = Xtotglla - 1
                  Xtotgllagt = Xtotgllagt - 1
               End If
               data_inf.Recordset.MoveNext
            Loop
            Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
            Xtotglla = 0
            XCol = XCol + 1
         Else
            Xtotglla = 0
            Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
            Xtotglla = 0
            XCol = XCol + 1
         End If
      End If
   End If

'   Xarchexel2.Cells(Xlin, XCol) = Trim(Str(Xtotglla))
   Xtotglla = 0
   XCol = 33
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotgllagt))
   XCol = 34
   Xpromed = Xtotgllagt / Xdiass
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
   XCol = 35
   Xpromed = Xtotgllagt * 100 / Xtotgraldos
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
End If

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
MiBaseact.Execute "Delete * from inflla"
Data1.RecordSource = "inflla"
Data1.Refresh

Xdiass = 1
Xlin = 5
XCol = 1
Xarchexel2.Cells(Xlin, XCol) = "TOTAL GENERAL DE LLAMADOS REALIZADOS"
Xdiass = 0
If t_mov.Text <> "" Then
   data_llama.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codzon in(1,2,3,5) and movilpas =" & t_mov.Text & " and codmed <>" & 959 & " and cancela is null order by fecha"
   data_llama.Refresh
Else
   data_llama.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and codzon in(1,2,3,5) and codmed <>" & 959 & " and cancela is null order by fecha"
   data_llama.Refresh
End If
'data_inf.RecordSource = "select * from inflla order by fecha"
'data_inf.Refresh
If data_llama.Recordset.RecordCount > 0 Then
   data_llama.Recordset.MoveFirst
   XCol = 2
   Xtotglla = 0
   Xtotgllagt = 0
   Xfecontrol = data_llama.Recordset("fecha")
   Do While Not data_llama.Recordset.EOF
         If data_llama.Recordset("fecha") = Xfecontrol Then
            Xtotglla = Xtotglla + 1
            Xtotgllagt = Xtotgllagt + 1
         Else
            Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
            Xtotglla = 1
            Xtotgllagt = Xtotgllagt + 1
            XCol = XCol + 1
            Xdiass = Xdiass + 1
         End If
         Xfecontrol = data_llama.Recordset("fecha")
         Data1.Recordset.AddNew
         Data1.Recordset("nro") = data_llama.Recordset("nro")
         Data1.Recordset("fecha") = data_llama.Recordset("fecha")
         Data1.Recordset("hora") = data_llama.Recordset("hora")
         Data1.Recordset("usuario") = data_llama.Recordset("usuario")
         If data_llama.Recordset("matric") >= 999999999 Then
            Data1.Recordset("matric") = 0
         Else
            Data1.Recordset("matric") = data_llama.Recordset("matric")
         End If
         Data1.Recordset("nombre") = data_llama.Recordset("nombre")
         Data1.Recordset("edad") = data_llama.Recordset("edad")
         Data1.Recordset("unied") = data_llama.Recordset("unied")
         Data1.Recordset("categ") = data_llama.Recordset("categ")
         Data1.Recordset("nomcat") = data_llama.Recordset("nomcat")
         If data_llama.Recordset("ci") >= 999999999 Then
            Data1.Recordset("ci") = 0
         Else
            Data1.Recordset("ci") = data_llama.Recordset("ci")
         End If
         Data1.Recordset("direcc") = data_llama.Recordset("direcc")
         Data1.Recordset("telef") = data_llama.Recordset("telef")
         Data1.Recordset("codzon") = data_llama.Recordset("codzon")
         Data1.Recordset("base") = data_llama.Recordset("base")
         Data1.Recordset("referen") = data_llama.Recordset("referen")
         Data1.Recordset("motcon") = data_llama.Recordset("motcon")
         Data1.Recordset("obsmot") = data_llama.Recordset("obsmot")
         Data1.Recordset("codmot") = data_llama.Recordset("codmot")
         Data1.Recordset("descol") = data_llama.Recordset("descol")
         Data1.Recordset("movilpas") = data_llama.Recordset("movilpas")
         Data1.Recordset("pend") = data_llama.Recordset("pend")
         If IsNull(data_llama.Recordset("fec_rea")) = True Then
            Data1.Recordset("fec_rea") = data_llama.Recordset("fecpas")
         Else
            Data1.Recordset("fec_rea") = data_llama.Recordset("fec_rea")
         End If
         If IsNull(data_llama.Recordset("hor_rea")) = True Then
            Data1.Recordset("hor_rea") = data_llama.Recordset("horpas")
         Else
            Data1.Recordset("hor_rea") = data_llama.Recordset("hor_rea")
         End If
         Data1.Recordset("fecpas") = data_llama.Recordset("fecpas")
         Data1.Recordset("horpas") = data_llama.Recordset("horpas")
         Data1.Recordset("fecsali") = data_llama.Recordset("fecsali")
         Data1.Recordset("horsali") = data_llama.Recordset("horsali")
         If IsNull(data_llama.Recordset("fec_llega")) = True Then
            Data1.Recordset("fec_llega") = data_llama.Recordset("fecpas")
         Else
            Data1.Recordset("fec_llega") = data_llama.Recordset("fec_llega")
         End If
         If IsNull(data_llama.Recordset("hor_llega")) = True Then
            Data1.Recordset("hor_llega") = data_llama.Recordset("horpas")
         Else
            Data1.Recordset("hor_llega") = data_llama.Recordset("hor_llega")
         End If
         Data1.Recordset("diag") = data_llama.Recordset("diag")
         Data1.Recordset("colormot") = data_llama.Recordset("colormot")
         Data1.Recordset("codmed") = data_llama.Recordset("codmed")
         Data1.Recordset("obs") = data_llama.Recordset("obs")
         Data1.Recordset("nommed") = data_llama.Recordset("nommed")
         Data1.Recordset("trasla") = data_llama.Recordset("trasla")
         Data1.Recordset("lugar") = data_llama.Recordset("lugar")
         Data1.Recordset("hsald") = data_llama.Recordset("hsald")
         Data1.Recordset("hllega") = data_llama.Recordset("hllega")
         Data1.Recordset("hzona") = data_llama.Recordset("hzona")
         Data1.Recordset("movil_rea") = data_llama.Recordset("movil_rea")
         Data1.Recordset("totdem") = data_llama.Recordset("totdem")
         Data1.Recordset("totend") = data_llama.Recordset("totend")
         Data1.Recordset("cancela") = data_llama.Recordset("cancela")
         Data1.Recordset.Update
      data_llama.Recordset.MoveNext
   Loop
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
   Xtotglla = 0
   Xdiass = Xdiass + 1
   XCol = 33
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotgllagt))
   XCol = 34
   Xpromed = Xtotgllagt / Xdiass
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
   Xtotgraldos = Xtotgllagt
End If

Xlin = Xlin + 1
XCol = 1
Xarchexel2.Cells(Xlin, XCol) = "TOTAL GENERAL DE LLAMADOS ZONA NORTE"
Data1.RecordSource = "select * from inflla where codzon =" & 2 & " and base =" & 0 & " order by fecha"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   XCol = 2
   Xtotglla = 0
   Xtotgllagt = 0
   Xfecontrol = Data1.Recordset("fecha")
   Do While Not Data1.Recordset.EOF
      If Data1.Recordset("fecha") = Xfecontrol Then
         Xtotglla = Xtotglla + 1
         Xtotgllagt = Xtotgllagt + 1
      Else
         Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
         Xtotglla = 1
         Xtotgllagt = Xtotgllagt + 1
         XCol = XCol + 1
      End If
      Xfecontrol = Data1.Recordset("fecha")
      Data1.Recordset.MoveNext
   Loop
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
   Xtotglla = 0
   XCol = 33
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotgllagt))
   XCol = 34
   Xpromed = Xtotgllagt / Xdiass
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
   XCol = 35
   Xpromed = Xtotgllagt * 100 / Xtotgraldos
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")

End If

Xlin = Xlin + 1
XCol = 1
Xarchexel2.Cells(Xlin, XCol) = "TOTAL GENERAL DE LLAMADOS ZONA COSTA"
Data1.RecordSource = "select * from inflla where codzon =" & 1 & " and base =" & 0 & " order by fecha"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   XCol = 2
   Xtotglla = 0
   Xtotgllagt = 0
   Xfecontrol = Data1.Recordset("fecha")
   Do While Not Data1.Recordset.EOF
      If Data1.Recordset("fecha") = Xfecontrol Then
         Xtotglla = Xtotglla + 1
         Xtotgllagt = Xtotgllagt + 1
      Else
         Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
         Xtotglla = 1
         Xtotgllagt = Xtotgllagt + 1
         XCol = XCol + 1
      End If
      Xfecontrol = Data1.Recordset("fecha")
      Data1.Recordset.MoveNext
   Loop
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
   Xtotglla = 0
   XCol = 33
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotgllagt))
   XCol = 34
   Xpromed = Xtotgllagt / Xdiass
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
   XCol = 35
   Xpromed = Xtotgllagt * 100 / Xtotgraldos
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
End If

'tALA
Xlin = Xlin + 1
XCol = 1
Xarchexel2.Cells(Xlin, XCol) = "TOTAL GENERAL DE LLAMADOS ZONA TALA/SJ"
Data1.RecordSource = "select * from inflla where codzon in (3,5) and base =" & 0 & " order by fecha"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   XCol = 2
   Xtotglla = 0
   Xtotgllagt = 0
   Xfecontrol = Data1.Recordset("fecha")
   Do While Not Data1.Recordset.EOF
      If Data1.Recordset("fecha") = Xfecontrol Then
         Xtotglla = Xtotglla + 1
         Xtotgllagt = Xtotgllagt + 1
      Else
         Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
         Xtotglla = 1
         Xtotgllagt = Xtotgllagt + 1
         XCol = XCol + 1
      End If
      Xfecontrol = Data1.Recordset("fecha")
      Data1.Recordset.MoveNext
   Loop
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotglla))
   Xtotglla = 0
   XCol = 33
   Xarchexel2.Cells(Xlin, XCol) = Trim(str(Xtotgllagt))
   XCol = 34
   Xpromed = Xtotgllagt / Xdiass
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
   XCol = 35
   Xpromed = Xtotgllagt * 100 / Xtotgraldos
   Xarchexel2.Cells(Xlin, XCol) = Format(Xpromed, "Standard")
End If
Xlin = 0
XCol = 0
Xtotglla = 0
Xtotgllagt = 0

''If data_inf.Recordset("totdem") >= "00:00" And data_inf.Recordset("totdem") <= "00:30" Then
''  If Xqdia = Day(data_inf.Recordset("fecha")) Then
'''    Set Xarchexel = Xlibexel.Worksheets.Add
''''      Xarchexel.Name = Trim(Str(Xqdia))
       
DoEvents
Xlibexel2.Save
Xlibexel2.Close
Xobjexel2.Quit
'Shell frm_menu.data_usuac.Recordset("destino") & "excel.exe " & Xarchtex, vbMaximizedFocus


End Sub

Private Sub Command5_Click()

                 If data_inf.Recordset.RecordCount > 0 Then
                    Dim MiBaseact As Database
                    Dim Unasesact As Workspace
                    Set Unasesact = Workspaces(0)
                    Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
                    MiBaseact.Execute "Delete * from inflla where cancela =" & 1
                    MiBaseact.Execute "Delete * from inflla where categ in ('55','MSP','56')"
                    data_inf.Refresh
'                    Do While Not data_inf.Recordset.EOF
'                       If IsNull(data_inf.Recordset("cancela")) = False Then
'                          If data_inf.Recordset("cancela") = 1 Then
'                             data_inf.Recordset.Delete
'                          Else
'                             If data_inf.Recordset("categ") = "55" Or data_inf.Recordset("categ") = "56" Or _
'                                data_inf.Recordset("categ") = "MSP" Then
'                                data_inf.Recordset.Delete
'                             End If
'                          End If
 '                      Else
 '                         If data_inf.Recordset("categ") = "55" Or data_inf.Recordset("categ") = "56" Or _
 '                            data_inf.Recordset("categ") = "MSP" Then
'                             data_inf.Recordset.Delete
'                          End If
'                       End If
'                       data_inf.Recordset.MoveNext
'                    Loop
'                    data_inf.Recordset.MoveFirst
                 End If
                 If Combo2.ListIndex = 1 Or Combo2.ListIndex = 2 Then
                 Else
                    data_inf.RecordSource = "Select * from inflla where codmot ='" & "V" & "' order by codmed"
                    data_inf.Refresh
                    data_inf.Recordset.MoveFirst
                 End If
                 Xcadllama = "TIEMPO EN DOMICILIO POR MEDICO MAS DE 30 MINUTOS (VERDES) Y PROMEDIOS >20 MINUTOS"
                 Open App.path & "\DEMORAS.txt" For Output As #1
                 Print #1, "SAPP S.A.                                                         FECHA: " & Date
                 Print #1, "-----------------------------------------------------------------------------------"
                 Print #1, ""
                 
                 Print #1, Xcadllama
                 Print #1, "=================================================================================="
                 Print #1, ""
                 Print #1, "FECHA/HORA LLEGA   N O M B R E S                                   EDAD  H.REA. DEMORA   MEDICO               MOVIL  MOTIVO CONSULTA                         CATEG.  COLOR     ZONA"
                 Print #1, "===================================================================================================================================================================================="
                 Xcadllama = ""
                 Dim Xtotminpro, Xelcoddemed, Xbandepro, Xtotpormed, Xbandenover As Long
                 data_inf.Recordset.MoveFirst
                 Xelcoddemed = data_inf.Recordset("codmed")
                 If Xelcoddemed = 0 Then
                    MsgBox "ATENCION: hay llamados sin registrar médico que lo realizó, comunique a despacho", vbCritical, "Mensaje"
                 End If
                 Do While Not data_inf.Recordset.EOF
                    If IsNull(data_inf.Recordset("hor_llega")) = True Then
                       data_inf.Recordset.MoveNext
                    Else
                       If IsNull(data_inf.Recordset("hor_llega")) = False Then
                          xhh = Val(Mid(data_inf.Recordset("hor_llega"), 1, 2))
                          xmm = Val(Mid(data_inf.Recordset("hor_llega"), 4, 2))
                       End If
                       If IsNull(data_inf.Recordset("hor_rea")) = False Then
                          Xhhh = Val(Mid(data_inf.Recordset("hor_rea"), 1, 2))
                          Xmmh = Val(Mid(data_inf.Recordset("hor_rea"), 4, 2))
                       End If
                       xdemh = Xhhh - xhh
                       xdemm = Xmmh - xmm
                       If data_inf.Recordset("fecha") < data_inf.Recordset("fec_llega") Then
                          If xdemh < 0 Then
                             xdemh = Xhhh - xhh
                             xdemh = xdemh + 24
                          End If
                       Else
                          If IsNull(data_inf.Recordset("fec_llega")) = True Then
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
                       data_inf.Recordset.Edit
                       Xbandepro = 0
                       If xdemh > 9 Then
                          If xdemm > 9 Then
                             data_inf.Recordset("totend") = Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                          Else
                             data_inf.Recordset("totend") = Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                          End If
                       Else
                          If xdemm > 9 Then
                             If xdemh < 0 Then
                                xdemh = 0
                             End If
                             data_inf.Recordset("totend") = "0" + Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                          Else
                             If xdemh < 0 Then
                                xdemh = 0
                             End If
                             data_inf.Recordset("totend") = "0" + Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                          End If
                       End If
                       Xbandepro = xdemh * 60
                       Xbandepro = Xbandepro + xdemm
                       Xtotminpro = Xtotminpro + Xbandepro
                       
                       data_inf.Recordset.Update
                       If Len(data_inf.Recordset("movilpas")) = 1 Then
                          Xnromov = "  " + str(data_inf.Recordset("movilpas"))
                       End If
                       If Len(data_inf.Recordset("movilpas")) = 2 Then
                          Xnromov = " " + str(data_inf.Recordset("movilpas"))
                       End If
                       If Len(data_inf.Recordset("movilpas")) = 3 Then
                          Xnromov = str(data_inf.Recordset("movilpas"))
                       End If
                       If Len(data_inf.Recordset("edad")) = 1 Then
                          Xedad = "  " + str(data_inf.Recordset("edad"))
                       End If
                       If Len(data_inf.Recordset("edad")) = 2 Then
                          Xedad = " " + str(data_inf.Recordset("edad"))
                       End If
                       If Len(data_inf.Recordset("edad")) = 3 Then
                          Xedad = str(data_inf.Recordset("edad"))
                       End If
                       Xmotcon = data_inf.Recordset("motcon")
                       XCat = data_inf.Recordset("categ")
                       Xdescol = data_inf.Recordset("descol")
                       Xnom = data_inf.Recordset("nombre")
                       If IsNull(data_inf.Recordset("nommed")) = False Then
                          Xnommed = data_inf.Recordset("nommed")
                       Else
                          Xnommed = "No Registrado"
                       End If
                       If Xnom <> "" Then
                          Xnom = Mid(Xnom, 1, 50)
                          xcuenta = Len(Xnom)
                          xcuenta = xcuenta + 1
                          For xcuenta = xcuenta To 50
                              Xnom = Xnom + " "
                          Next
                       Else
                          Xnom = "                                                  "
                       End If
                       If Xnommed <> "" Then
                          Xnommed = Mid(Xnommed, 1, 25)
                          xcuenta = Len(Xnommed)
                          xcuenta = xcuenta + 1
                          For xcuenta = xcuenta To 25
                              Xnommed = Xnommed + " "
                          Next
                       Else
                          Xnommed = "                         "
                       End If
                       If Xmotcon <> "" Then
                          Xmotcon = Mid(Xmotcon, 1, 40)
                          xcuenta = Len(Xmotcon)
                          xcuenta = xcuenta + 1
                          For xcuenta = xcuenta To 40
                              Xmotcon = Xmotcon + " "
                          Next
                       Else
                          Xmotcon = "                                        "
                       End If
                       If XCat <> "" Then
                          XCat = Mid(XCat, 1, 6)
                          xcuenta = Len(XCat)
                          xcuenta = xcuenta + 1
                          For xcuenta = xcuenta To 6
                              XCat = XCat + " "
                          Next
                       Else
                          XCat = "      "
                       End If
                       If Xdescol <> "" Then
                          Xdescol = Mid(Xdescol, 1, 9)
                          xcuenta = Len(Xdescol)
                          xcuenta = xcuenta + 1
                          For xcuenta = xcuenta To 9
                              Xdescol = Xdescol + " "
                          Next
                       Else
                          Xdescol = "         "
                       End If
                       If data_inf.Recordset("codzon") = 1 Then
                          Xzona = "Z.COSTA"
                       Else
                          Xzona = "Z.NORTE"
                       End If
                       Xtotpormed = Xtotpormed + 1
                       If data_inf.Recordset("totend") > "00:30" Then
                          Xtotal = Xtotal + 1
'                          Xtotpormed = Xtotpormed + 1
                          Print #1, CStr(data_inf.Recordset("fec_llega")) + " " + data_inf.Recordset("hor_llega") + " " + Xnom + " " + Xedad _
                          ; " " + data_inf.Recordset("hor_rea") + " " + data_inf.Recordset("totend") + " " + Xnommed _
                          + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama
                          Xbandenover = 9
                       Else
                          If Xbandenover = 9 Then
                          Else
                             Xbandenover = 1
                          End If
                       End If
                       Xelcoddemed = data_inf.Recordset("codmed")
                       data_inf.Recordset.MoveNext
                       If data_inf.Recordset.EOF = True Then
                          Xtotpormed = Xtotpormed - 1
                          If Xtotminpro > 0 Then
                             Xtotminpro = Xtotminpro / Xtotpormed
                          Else
                             Xtotminpro = 0
                          End If
                          If Xtotminpro > 20 Then
                             Print #1, "================================================================================="
                             Print #1, "DEMORA PROMEDIO >20 MINUTOS DR: " + Xnommed + " ......:" + Format(Xtotminpro, "Standard") + "......TOTAL LLAMADOS DEL MEDICO (Verdes) : " + Trim(str(Xtotpormed))
                             Print #1, "================================================================================="
                          Else
                             If Xelcoddemed > 0 Then
                                Print #1, "----------------------------------------------------------------------------------------------"
                             End If
                          End If
                          Xtotpormed = 1
                          Xtotminpro = 0
                       Else
                          If Xelcoddemed = data_inf.Recordset("codmed") Then
                          Else
'                             Print #1, ""
                             Xtotpormed = Xtotpormed - 1
                             If Xtotminpro > 0 Then
                                If Xtotpormed > 0 Then
                                   Xtotminpro = Xtotminpro / Xtotpormed
                                Else
                                  Xtotminpro = 0
                                End If
                             Else
                                Xtotminpro = 0
                             End If
                             If Xtotminpro > 20 Then
                                Print #1, ""
                                Print #1, "================================================================================="
                                Print #1, "DEMORA PROMEDIO >20 MINUTOS DR: " + Xnommed + " ......:" + Format(Xtotminpro, "Standard") + "......TOTAL LLAMADOS DEL MEDICO: (Verdes) : " + Trim(str(Xtotpormed))
                                Print #1, "================================================================================="
                             Else
                                If Xelcoddemed > 0 Then
                                   If Xbandenover = 9 Then
                                      Print #1, "----------------------------------------------------------------------------------------------"
                                      Print #1, "================================================================================="
                                      Print #1, "DEMORA PROMEDIO >20 MINUTOS DR: " + Xnommed + " ......:" + Format(Xtotminpro, "Standard") + "......TOTAL LLAMADOS DEL MEDICO: (Verdes) : " + Trim(str(Xtotpormed))
                                      Print #1, "================================================================================="
                                   
                                   Else
'                                      Print #1, "--------------------------------------"
                                   End If
                                End If
                             End If
                             Xtotpormed = 1
                             Xtotminpro = 0
                             Xbandenover = 1
                          End If
                       End If
                    End If
                 Loop
                 Print #1, "================================================="
                 Print #1, "TOTAL......:" + str(Xtotal)
                 Print #1, "================================================="
                 Xtotal = 0
                 
                 Xtotal = 0
                 data_inf.RecordSource = "Select * from inflla order by codmed,codmot"
                 data_inf.Refresh
                 
                 data_inf.Recordset.MoveFirst
                 Print #1, "-----------------------------------------------------------------------------------------------"
                 Print #1, ""
                 Print #1, "================================================================="
                 Xcadllama = "DEMORAS EN SALIDA DE LLAMADOS ORDENADOS POR MEDICO Y POR CLAVE"
                 Print #1, Xcadllama
                 Print #1, "================================================================="
                 Print #1, "FECHA/HORA PAS.   N O M B R E S                                   EDAD  H.SALE DEMORA   MEDICO               MOVIL  MOTIVO CONSULTA                         CATEG.  COLOR     ZONA"
                 Print #1, "===================================================================================================================================================================================="
                 Xcadllama = ""
                 Do While Not data_inf.Recordset.EOF
                    If IsNull(data_inf.Recordset("horsali")) = True Then
                       data_inf.Recordset.MoveNext
                    Else
                        If IsNull(data_inf.Recordset("horpas")) = False Then
                           xhh = Val(Mid(data_inf.Recordset("horpas"), 1, 2))
                           xmm = Val(Mid(data_inf.Recordset("horpas"), 4, 2))
                        End If
                        If IsNull(data_inf.Recordset("horsali")) = False Then
                           Xhhh = Val(Mid(data_inf.Recordset("horsali"), 1, 2))
                           Xmmh = Val(Mid(data_inf.Recordset("horsali"), 4, 2))
                        End If
                        xdemh = Xhhh - xhh
                        xdemm = Xmmh - xmm
                        If data_inf.Recordset("fecpas") < data_inf.Recordset("fecsali") Then
                           xdemh = Xhhh - xhh
                           xdemh = xdemh + 24
                        End If
                        If xdemh > 0 Then
                           If xdemm < 0 Then
                              xdemm = xdemm + 60
                              xdemh = xdemh - 1
                           End If
                        Else
                           If xdemm < 0 Then
                              xdemm = xdemm + 60
                           End If
                        End If
                        data_inf.Recordset.Edit
                        If xdemh > 9 Then
                           If xdemm > 9 Then
                              data_inf.Recordset("totdem") = Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                           Else
                              data_inf.Recordset("totdem") = Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                           End If
                        Else
                           If xdemm > 9 Then
                              If xdemh < 0 Then
                                 xdemh = 0
                              End If
                              data_inf.Recordset("totdem") = "0" + Trim(str(xdemh)) + ":" + Trim(str(xdemm))
                           Else
                              If xdemh < 0 Then
                                 xdemh = 0
                              End If
                              data_inf.Recordset("totdem") = "0" + Trim(str(xdemh)) + ":0" + Trim(str(xdemm))
                           End If
                        End If
                        data_inf.Recordset.Update
                        If Len(data_inf.Recordset("movilpas")) = 1 Then
                           Xnromov = "  " + str(data_inf.Recordset("movilpas"))
                        End If
                        If Len(data_inf.Recordset("movilpas")) = 2 Then
                           Xnromov = " " + str(data_inf.Recordset("movilpas"))
                        End If
                        If Len(data_inf.Recordset("movilpas")) = 3 Then
                           Xnromov = str(data_inf.Recordset("movilpas"))
                        End If
                        If Len(data_inf.Recordset("edad")) = 1 Then
                           Xedad = "  " + str(data_inf.Recordset("edad"))
                        End If
                        If Len(data_inf.Recordset("edad")) = 2 Then
                           Xedad = " " + str(data_inf.Recordset("edad"))
                        End If
                        If Len(data_inf.Recordset("edad")) = 3 Then
                           Xedad = str(data_inf.Recordset("edad"))
                        End If
                        Xmotcon = data_inf.Recordset("motcon")
                        XCat = data_inf.Recordset("categ")
                        Xdescol = data_inf.Recordset("descol")
                        Xnom = data_inf.Recordset("nombre")
                        If IsNull(data_inf.Recordset("nommed")) = False Then
                           Xnommed = data_inf.Recordset("nommed")
                        Else
                           Xnommed = ""
                        End If
                        If Xnom <> "" Then
                           Xnom = Mid(Xnom, 1, 50)
                           xcuenta = Len(Xnom)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 50
                               Xnom = Xnom + " "
                           Next
                        Else
                           Xnom = "                                                  "
                        End If
                        If Xnommed <> "" Then
                           Xnommed = Mid(Xnommed, 1, 25)
                           xcuenta = Len(Xnommed)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 25
                               Xnommed = Xnommed + " "
                           Next
                        Else
                           Xnommed = "                         "
                        End If
                        If Xmotcon <> "" Then
                           Xmotcon = Mid(Xmotcon, 1, 40)
                           xcuenta = Len(Xmotcon)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 40
                               Xmotcon = Xmotcon + " "
                           Next
                        Else
                           Xmotcon = "                                        "
                        End If
                        If XCat <> "" Then
                           XCat = Mid(XCat, 1, 6)
                           xcuenta = Len(XCat)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 6
                               XCat = XCat + " "
                           Next
                        Else
                           XCat = "      "
                        End If
                        If Xdescol <> "" Then
                           Xdescol = Mid(Xdescol, 1, 9)
                           xcuenta = Len(Xdescol)
                           xcuenta = xcuenta + 1
                           For xcuenta = xcuenta To 9
                               Xdescol = Xdescol + " "
                           Next
                        Else
                           Xdescol = "         "
                        End If
                        If data_inf.Recordset("codzon") = 1 Then
                           Xzona = "Z.COSTA"
                        Else
                           Xzona = "Z.NORTE"
                        End If
                        If data_inf.Recordset("codmot") = "R" Then
                           If data_inf.Recordset("totdem") > "00:03" Then
                              Xtotal = Xtotal + 1
                              Print #1, CStr(data_inf.Recordset("fecpas")) + " " + data_inf.Recordset("horpas") + " " + Xnom + " " + Xedad _
                              ; " " + data_inf.Recordset("horsali") + " " + data_inf.Recordset("totdem") + " " + Xnommed _
                              + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama
                              Xtotgral = Xtotgral + 1
                           End If
                        Else
                           If data_inf.Recordset("codmot") = "A" Then
                              If data_inf.Recordset("totdem") > "00:05" Then
                                 Xtotal = Xtotal + 1
                                 Print #1, CStr(data_inf.Recordset("fecpas")) + " " + data_inf.Recordset("horpas") + " " + Xnom + " " + Xedad _
                                 ; " " + data_inf.Recordset("horsali") + " " + data_inf.Recordset("totdem") + " " + Xnommed _
                                 + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama
                                 Xtotgral = Xtotgral + 1
                              End If
                           Else
                              If data_inf.Recordset("codmot") = "V" Then
                                 If data_inf.Recordset("totdem") > "00:10" Then
                                    Xtotal = Xtotal + 1
                                    Print #1, CStr(data_inf.Recordset("fecpas")) + " " + data_inf.Recordset("horpas") + " " + Xnom + " " + Xedad _
                                    ; " " + data_inf.Recordset("horsali") + " " + data_inf.Recordset("totdem") + " " + Xnommed _
                                    + " " + Xnromov + " " + Xmotcon + " " + XCat + " " + Xdescol + " " + Xzona + " " + Xcadllama
                                    Xtotgral = Xtotgral + 1
                                 End If
                              End If
                           End If
                        End If
                        data_inf.Recordset.MoveNext
                    End If
                 Loop
                 Print #1, "================================================="
                 Print #1, "TOTAL......:" + str(Xtotal)
                 Print #1, "================================================="
                 Xdemmas30 = Xtotal
                 Xtotal = 0
                 data_inf.Recordset.MoveFirst
                 
'                 Print #1, ""
'                 Print #1, "======================================================="
'                 Print #1, "TOTAL GENERAL DE LLAMADOS EN LA FECHA.................:" + Str(Xtotgral)
'                 Print #1, "======================================================="
'                 Print #1, ""
                 
                 Close #1
                 MsgBox "Proceso Terminado..."
                 OLE1.Action = 1
                 OLE1.DoVerb (-1)



End Sub

Private Sub Form_Load()
'data_llama.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_llama.ConnectionString = "dsn=" & Xconexrmt
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.RecordSource = "inflla"
'data_inf.Refresh
'Data1.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
Data1.DatabaseName = App.path & "\informes.mdb"
'Data1.RecordSource = "inflla"
'Data1.Refresh
OLE1.SourceDoc = App.path & "\DEMORAS.txt"

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
   mhd.SetFocus
End If

End Sub

Private Sub mhd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhh.SetFocus
End If

End Sub

Private Sub mhh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub
