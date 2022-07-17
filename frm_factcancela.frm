VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_factcancela 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factura a cancelar"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8100
   Icon            =   "frm_factcancela.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   8100
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      Picture         =   "frm_factcancela.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Seleccionar todos los registros"
      Top             =   7440
      Width           =   495
   End
   Begin VB.Data data_lineas 
      Caption         =   "data_lineas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Data data_estudiobus 
      Caption         =   "data_estudiobus"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   5520
      TabIndex        =   21
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data data_estudio 
      Caption         =   "data_estudio"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_codcaja 
      Caption         =   "data_codcaja"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton btn_graba 
      Caption         =   "Command3"
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ACEPTAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7560
      Width           =   2895
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4471
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Total $."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Medicación"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Factura"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Línea"
         Object.Width           =   776
      EndProperty
   End
   Begin VB.Data data_cance992 
      Caption         =   "data_cance992"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_arq 
      Caption         =   "data_arq"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "arqueo"
      Top             =   7680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSDBGrid.DBGrid DBGrid4 
      Bindings        =   "frm_factcancela.frx":0B14
      Height          =   2535
      Left            =   120
      OleObjectBlob   =   "frm_factcancela.frx":0B2B
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CheckBox charq 
      BackColor       =   &H0080FFFF&
      Caption         =   "Mostrar deudas de arqueo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data data_emi 
      Caption         =   "data_emi"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "frm_factcancela.frx":1ECE
      Height          =   3135
      Left            =   120
      OleObjectBlob   =   "frm_factcancela.frx":1EE8
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Data data_deudas 
      Caption         =   "data_deudas"
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
      Top             =   3840
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Desde deudas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5400
      TabIndex        =   13
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      Picture         =   "frm_factcancela.frx":35EB
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSMask.MaskEdBox mf 
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.TextBox t_nrolin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4320
      MaxLength       =   2
      TabIndex        =   9
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox t_nrofact 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   7
      Top             =   7080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Selección manual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Data data_lintemp 
      Caption         =   "data_lintemp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_factcancela.frx":3B75
      Height          =   3495
      Left            =   120
      OleObjectBlob   =   "frm_factcancela.frx":3B8C
      TabIndex        =   0
      Top             =   360
      Width           =   7815
   End
   Begin VB.Data data_cab 
      Caption         =   "data_cab"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_lin 
      Caption         =   "data_lin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label esmedica 
      Height          =   255
      Left            =   720
      TabIndex        =   27
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label labhayerror 
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   7560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label labidpedido 
      Height          =   255
      Left            =   1320
      TabIndex        =   24
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label labmotivoref 
      Height          =   375
      Left            =   480
      TabIndex        =   23
      Top             =   3960
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label labcedula 
      Height          =   375
      Left            =   4320
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nro.de linea:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nro.FACT:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label labbande 
      Caption         =   "labbande"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Seleccione los registros a anular y presione aceptar."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   4935
   End
   Begin VB.Label labcli 
      BackColor       =   &H00FF8080&
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
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "CUENTA:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1080
      Picture         =   "frm_factcancela.frx":4F6F
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   3615
   End
End
Attribute VB_Name = "frm_factcancela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_graba_Click()
Dim Xrub As Long
Dim Xiva, Xvaltimme As Double
'''On Error GoTo Vererror
Dim XValtim As Long
Dim Xlafenaci As String
Dim Xnomedica, Xlin997, Xcuotabase, Xquerr As Integer
Dim Xnroform As String
Dim Xestco, Xestfl As Long
Dim Xmotivoref As String
Dim Xtelcmt As String
Dim MensajeCMT As String
MensajeCMT = vbNo
Dim Xcantveces As Integer
Xmotivoref = ""

Xtelcmt = ""

data_estudio.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_estudio.RecordSource = "estudios"
data_estudio.Refresh

frm_factura.DBCombo1.Enabled = True

frm_factura.t_cant.Text = 1

If frm_factura.Label7.Caption = "NC E-TICKET" Or frm_factura.Label7.Caption = "NC E-FACTURA" Or frm_factura.Label7.Caption = "DEV.RECIBO" Then
   If frm_factura.Label5.Caption <> "" Then
      data_estudio.Recordset.FindFirst "codest =" & frm_factura.Label5.Caption
   End If
End If
Xvaltimme = 0
Xlin997 = 0
Xestco = data_estudio.Recordset("codest")
Xestfl = data_estudio.Recordset("flia")
Xcuotabase = 0
Xnomedica = 0
Xquerr = 0
Xcantveces = 1
    
If frm_factura.dbcboprom.Visible = True Then
   If frm_factura.dbcboprom.Text <> "" Then
      If frm_factura.labcodpro.Caption <> "" Then
         frm_factura.data_func.RecordSource = "select * from vende_func where idfunc =" & Val(frm_factura.labcodpro.Caption) & " and nombre ='" & frm_factura.dbcboprom.Text & "'"
         frm_factura.data_func.Refresh
         If frm_factura.data_func.Recordset.RecordCount > 0 Then
         Else
            Xnomedica = 33
         End If
      Else
         frm_factura.labcodpro.Caption = "0"
         frm_factura.data_func.RecordSource = "select * from vende_func where idfunc =" & Val(frm_factura.labcodpro.Caption) & " and nombre ='" & frm_factura.dbcboprom.Text & "'"
         frm_factura.data_func.Refresh
         If frm_factura.data_func.Recordset.RecordCount > 0 Then
         Else
            Xnomedica = 33
         End If
      End If
   Else
      Xnomedica = 33
   End If
End If
             
Xconvprom = ""
      
      
If Trim(frm_factura.labmedicacion.Caption) <> "" Then
   Xdescmedic = frm_factura.labmedicacion.Caption
Else
   Xdescmedic = "S/D"
End If

If frm_factura.labtimemi.Caption <> "" Then
   If Val(frm_factura.labtimemi.Caption) > 0 Then
      If frm_factura.txt_precio.Text <> "" Then
         If Val(frm_factura.txt_precio.Text) > 0 Then
            frm_factura.txt_precio.Text = Format(frm_factura.txt_precio.Text, "Standard") - Format(frm_factura.labtimemi.Caption, "Standard")
         End If
      End If
   End If
End If
If frm_factura.labdeudaemi.Caption <> "" Then
   If Val(frm_factura.labdeudaemi.Caption) > 0 Then
      If frm_factura.txt_precio.Text <> "" Then
         If Val(frm_factura.txt_precio.Text) > 0 Then
            frm_factura.txt_precio.Text = Format(frm_factura.txt_precio.Text, "Standard") - Format(frm_factura.labdeudaemi.Caption, "Standard")
         End If
      End If
   End If
End If

If frm_factura.Label5.Caption = "" Then
   MsgBox "No ingresó servicio a facturar!", vbInformation
End If

XValtim = 0
'If frm_factura.data_lineas.Recordset.RecordCount >= 1 And frm_factura.data_estudio.Recordset("codest") = 997 Then
'   Xlin997 = 8
'Else
'   Xlin997 = 0
'End If
If frm_factura.Label7.Caption = "NC E-TICKET" Or frm_factura.Label7.Caption = "NC E-FACTURA" Or frm_factura.Label7.Caption = "ND E-FACTURA" Or frm_factura.Label7.Caption = "ND E-TICKET" Then
   Xnomedica = 46
End If
If data_estudio.Recordset("codest") = 992 Then
   If frm_factura.dbcboprom.Text = "" Then
      Xnomedica = 59
   End If
End If

If frm_factura.data_lineas.Recordset.RecordCount >= 150 Or Xnomedica = 7 Or Xnomedica = 8 Or Xcoddeu = 9 Or Xlin997 = 8 Or Xnomedica = 31 Or Xnomedica = 59 Or Xnomedica = 33 Then
   If Xnomedica = 7 Or Xnomedica = 33 Then
      If Xnomedica = 7 Then
         MsgBox "No ingresó dato solicitado, verifique!", vbCritical, "Facturación"
      Else
         MsgBox "No ingresó Promotor de la Afiliación correctamente.", vbCritical, "Facturación"
      End If
   Else
      If Xnomedica = 8 Then
         MsgBox "No ingresó MEDICO que realiza", vbCritical, "Facturación"
      Else
         If Xnomedica = 31 Then
            MsgBox "ATENCION!! Usuario con servicios restringidos. Verifique con administración al 097215419", vbInformation
         Else
            If Xnomedica = 59 Then
               MsgBox "ATENCION!! No ingresó PROMOTOR de la afiliación!", vbInformation
            Else
                MsgBox "ATENCION!! alcanzó el límite de líneas por factura", vbCritical
            End If
         End If
      End If
   End If
Else
   If data_estudio.Recordset("codest") = 60106 Or _
      data_estudio.Recordset("codest") = 60108 Or _
      data_estudio.Recordset("codest") = 993 Or _
      data_estudio.Recordset("codest") = 994 Or _
      data_estudio.Recordset("codest") = 997 Or _
      data_estudio.Recordset("codest") = 996 Or _
      data_estudio.Recordset("codest") = 60105 Or _
      data_estudio.Recordset("codest") = 999 Or _
      data_estudio.Recordset("codest") = 60109 Or _
      data_estudio.Recordset("codest") = 80011 Or _
      data_estudio.Recordset("codest") = 80012 Or _
      data_estudio.Recordset("codest") = 80013 Or _
      data_estudio.Recordset("codest") = 80014 Or _
      data_estudio.Recordset("codest") = 80016 Or _
      data_estudio.Recordset("codest") = 80015 Then
      If XQuefac = 4 Or XQuefac = 21 Then
         Xquerr = 0
      Else
         Xquerr = 1
      End If
   Else
      If XQuefac <> 4 Then
         Xquerr = 0
      Else
         Xquerr = 1
      End If
   End If
End If

If Xquerr = 0 Then
   If data_estudio.Recordset("flia") = 1 Or _
      data_estudio.Recordset("flia") = 14 Or _
      data_estudio.Recordset("flia") = 10 Or _
      data_estudio.Recordset("flia") = 9 Then
      If frm_factura.dbcbomed.Text = "" And Xnomedica <> 46 Then
         MsgBox "Debe Ingresar Médico para éste servicio", vbInformation, "Facturación"
      Else
         Xcandelin = Xcandelin + 1
         frm_factura.data_lineas.Recordset.AddNew
         If frm_factura.Label7.Caption = "NC E-TICKET" Or frm_factura.Label7.Caption = "NC E-FACTURA" Or frm_factura.Label7.Caption = "ND E-FACTURA" Or frm_factura.Label7.Caption = "ND E-TICKET" Then
            Xmotivoref = labmotivoref.Caption
            If Xmotivoref <> "" Then
               frm_factura.labmotivo.Caption = Xmotivoref
               If frm_factura.labseriecance.Caption = "XX" Then
                  frm_factura.data_lincance.RecordSource = "Select * from linmmdd where factura =" & Val(frm_factura.labfaccance.Caption) & " and fecha =#" & Format(frm_factura.labfeccance.Caption, "yyyy-mm-dd") & "#"
                  frm_factura.data_lincance.Refresh
               Else
                  If frm_factura.labidemi.Caption = 5 Then
                     frm_factura.data_lincance.RecordSource = "Select * from deudas where documento =" & Val(frm_factura.labfaccance.Caption)
                     frm_factura.data_lincance.Refresh
                  Else
                     frm_factura.data_lincance.RecordSource = "Select * from clirespl where cl_numero =" & Val(frm_factura.labfaccance.Caption) & " and cl_socmnro ='" & Trim(frm_factura.labseriecance.Caption) & "' and cl_fnac =#" & Format(frm_factura.labfeccance.Caption, "yyyy-mm-dd") & "#"
                     frm_factura.data_lincance.Refresh
                  End If
               End If
               If frm_factura.data_lincance.Recordset.RecordCount > 0 And Val(frm_factura.labidemi.Caption) <> 5 Then
                  If frm_factura.labseriecance.Caption <> "XX" Then
                     If Val(frm_factura.labtot.Caption) >= frm_factura.data_lincance.Recordset("saldo_doc") Then
                        MsgBox "Ha excedido el importe de la factura", vbCritical
                     End If
                  End If
                  If frm_factura.labseriecance.Caption = "XX" Then
                     frm_factura.data_lineas.Recordset("tipodocref") = 101
                  Else
                     frm_factura.data_lineas.Recordset("tipodocref") = frm_factura.data_lincance.Recordset("cl_tipocli")
                  End If
                  frm_factura.data_lineas.Recordset("serieref") = frm_factura.labseriecance.Caption
                  If Len(frm_factura.labfaccance.Caption) > 7 Then
                     frm_factura.data_lineas.Recordset("nrofactref") = Val(Mid(frm_factura.labfaccance.Caption, 1, 7))
                  Else
                     frm_factura.data_lineas.Recordset("nrofactref") = Val(frm_factura.labfaccance.Caption)
                  End If
                  If frm_factura.labseriecance.Caption = "XX" Then
                     frm_factura.data_lineas.Recordset("fechafact") = CDate(frm_factura.labfeccance.Caption)
                  Else
                     frm_factura.data_lineas.Recordset("fechafact") = frm_factura.data_lincance.Recordset("cl_fnac")
                  End If
                  frm_factura.data_lineas.Recordset("motivoref") = Mid(Xmotivoref, 1, 90)
                  frm_factura.data_lineas.Recordset("linearef") = Val(frm_factura.lablinea.Caption)
               Else
                  If frm_factura.labidemi.Caption = 5 Then
                     If frm_factura.Label7.Caption = "NC E-FACTURA" Then
                        frm_factura.data_lineas.Recordset("tipodocref") = 111
                     Else
                        frm_factura.data_lineas.Recordset("tipodocref") = 101
                     End If
                     frm_factura.data_lineas.Recordset("serieref") = frm_factura.labseriecance.Caption
                     If Len(frm_factura.labfaccance.Caption) > 7 Then
                        frm_factura.data_lineas.Recordset("nrofactref") = Val(Mid(frm_factura.labfaccance.Caption, 1, 7))
                     Else
                        frm_factura.data_lineas.Recordset("nrofactref") = Val(frm_factura.labfaccance.Caption)
                     End If
                     frm_factura.data_lineas.Recordset("fechafact") = CDate(frm_factura.labfeccance.Caption)
                     frm_factura.data_lineas.Recordset("motivoref") = Mid(Xmotivoref, 1, 90)
                     frm_factura.data_lineas.Recordset("linearef") = 1
                  Else
                     If frm_factura.labseriecance.Caption = "XX" Then
                        If frm_factura.Label7.Caption = "NC E-FACTURA" Then
                           frm_factura.data_lineas.Recordset("tipodocref") = 111
                        Else
                           frm_factura.data_lineas.Recordset("tipodocref") = 101
                        End If
                        frm_factura.data_lineas.Recordset("serieref") = "A"
                        If Len(frm_factura.labfaccance.Caption) > 7 Then
                           frm_factura.data_lineas.Recordset("nrofactref") = Val(Mid(frm_factura.labfaccance.Caption, 1, 7))
                        Else
                           frm_factura.data_lineas.Recordset("nrofactref") = Val(frm_factura.labfaccance.Caption)
                        End If
                        frm_factura.data_lineas.Recordset("fechafact") = CDate(frm_factura.labfeccance.Caption)
                        frm_factura.data_lineas.Recordset("motivoref") = Mid(Xmotivoref, 1, 90)
                        frm_factura.data_lineas.Recordset("linearef") = 1
                     Else
                        MsgBox "No se encuentra número de factura a cancelar"
                     End If
                  End If
               End If
            Else
               MsgBox "No ingresó motivo de cancelación"
            End If
            '''frm_factura.btn_fin.SetFocus
         End If

         frm_factura.data_lineas.Recordset("univta") = 0
         If frm_factura.txt_precio.Text <> "" Then
            If frm_factura.txt_precio.Text > 0 Then
               If Val(frm_factura.Label5.Caption) = 995 Or Val(frm_factura.Label5.Caption) = 990 Then
                  frm_factura.data_lineas.Recordset("tipo_mov") = 1 'tipo de iva (indic de fact)
               Else
                  frm_factura.data_lineas.Recordset("tipo_mov") = 2 'tipo de iva (indic de fact)
               End If
            Else
               frm_factura.data_lineas.Recordset("tipo_mov") = 5
            End If
         Else
            frm_factura.data_lineas.Recordset("tipo_mov") = 5
         End If
         frm_factura.data_lineas.Recordset("factura") = 0
         frm_factura.data_lineas.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
         frm_factura.data_lineas.Recordset("fecha") = Format(frm_factura.mf.Text, "dd/mm/yyyy")
         frm_factura.data_lineas.Recordset("cod_cli") = frm_factura.labmatri.Caption
         frm_factura.data_lineas.Recordset("nom_cli") = Mid(frm_factura.labnomb.Caption, 1, 30)
         frm_factura.data_lineas.Recordset("cod_prod") = data_estudio.Recordset("codest")
         frm_factura.data_lineas.Recordset("nom_prod") = Mid(data_estudio.Recordset("descrip"), 1, 50)
         frm_factura.data_lineas.Recordset("cantidad") = 1
         frm_factura.data_lineas.Recordset("moneda") = "SR" 'Serie
         If frm_factura.txt_rut.Visible = True Then
            If Trim(frm_factura.txt_rut.Text) <> "" Then
               frm_factura.data_lineas.Recordset("ruc") = frm_factura.txt_rut.Text
            End If
         End If
         frm_factura.data_lineas.Recordset("operador") = WElusuario
         frm_factura.data_lineas.Recordset("hora") = Format(Time, "HH:mm")
         frm_factura.data_lineas.Recordset("nro_flia") = data_estudio.Recordset("flia")
         frm_factura.data_lineas.Recordset("nom_flia") = data_estudio.Recordset("nomflia")
         frm_factura.data_lineas.Recordset("convenio") = frmabm.data_clientes.Recordset("cl_codconv")
         frm_factura.data_lineas.Recordset("ced_socio") = frmabm.data_clientes.Recordset("cl_cedula")
         frm_factura.data_lineas.Recordset("fact") = frmabm.data_clientes.Recordset("cl_codced")
         If data_estudio.Recordset("codest") = 997 Then
            frm_factura.data_lineas.Recordset("rub_cont") = 113003
            frm_factura.data_lineas.Recordset("tipo_mov") = 1
            Xrub = 113003
         Else
            If data_estudio.Recordset("codest") = 993 Or _
               data_estudio.Recordset("codest") = 994 Then
               frm_factura.data_lineas.Recordset("rub_cont") = 112022
               frm_factura.data_lineas.Recordset("tipo_mov") = 1
               Xrub = 112022
            Else
               If data_estudio.Recordset("codest") = 999 Then
                  frm_factura.data_lineas.Recordset("tipo_mov") = 1
                  If IsNull(frmabm.data_clientes.Recordset("cl_nrocobr")) = False Then
                     If frmabm.data_clientes.Recordset("cl_nrocobr") = 615 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 616 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 635 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 602 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 113 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 653 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 672 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 1 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 10 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 201 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 512 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 636 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 685 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 208 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 209 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 8 Or _
                        frmabm.data_clientes.Recordset("cl_nrocobr") = 0 Then
                        frm_factura.data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("cbrcuo")
                        Xrub = frmabm.data_parsec.Recordset("cbrcuo")
                        frm_factura.data_lineas.Recordset("grupo") = frmabm.data_clientes.Recordset("cl_nrocobr")
                     Else
                        frm_factura.data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("cbrcuo")
                        Xrub = frmabm.data_parsec.Recordset("cbrcuo")
                        frm_factura.data_lineas.Recordset("grupo") = frmabm.data_clientes.Recordset("cl_nrocobr")
                        Xcuotabase = 9
                     End If
                  Else
                     frm_factura.data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("cbrcuo")
                     Xrub = frmabm.data_parsec.Recordset("cbrcuo")
                     frm_factura.data_lineas.Recordset("grupo") = frmabm.data_clientes.Recordset("cl_nrocobr")
                  End If
               Else
                  If data_estudio.Recordset("codest") = 992 Then
                     frm_factura.data_lineas.Recordset("rub_cont") = 513007
                     Xrub = 513007
                  Else
                     If Xfpago = 2 Then
                        frm_factura.data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("srvcrd")
                        Xrub = frmabm.data_parsec.Recordset("srvcrd")
                     Else
                        frm_factura.data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("srvcnt")
                        Xrub = frmabm.data_parsec.Recordset("srvcnt")
                     End If
                  End If
               End If
            End If
         End If
         If data_estudio.Recordset("codest") = 996 Then
            frm_factura.data_lineas.Recordset("rub_cont") = 211473
            frm_factura.data_lineas.Recordset("tipo_mov") = 1
            Xrub = 211473
         End If
         data_codcaja.Recordset.FindFirst "numero =" & Xrub
         If Not data_codcaja.Recordset.NoMatch Then
            frm_factura.data_lineas.Recordset("rub_nomb") = data_codcaja.Recordset("nombre")
         Else
            frm_factura.data_lineas.Recordset("rub_nomb") = "NO REG."
         End If
         If data_estudio.Recordset("codest") = 60106 Then
            frm_factura.data_lineas.Recordset("rub_cont") = 211397
            frm_factura.data_lineas.Recordset("rub_nomb") = "M. SMI"
         End If
         If data_estudio.Recordset("codest") = 60105 Then
            frm_factura.data_lineas.Recordset("rub_cont") = 211397
            frm_factura.data_lineas.Recordset("rub_nomb") = "M. SMI"
         End If
         If data_estudio.Recordset("codest") = 60108 Then
            frm_factura.data_lineas.Recordset("rub_cont") = 211302
            frm_factura.data_lineas.Recordset("rub_nomb") = "M.UNIVERSAL"
         End If
         frm_factura.data_lineas.Recordset("arancel") = Format(frm_factura.txt_precio.Text, "Standard")
         frm_factura.data_lineas.Recordset("tot_lin") = Format(frm_factura.txt_precio.Text, "Standard")
         Xiva = frm_factura.data_lineas.Recordset("tot_lin") / 1.1
         Xiva = Xiva * 0.1
         If frm_factura.Label5.Caption <> "" Then
            If Val(frm_factura.Label5.Caption) = 995 Or Val(frm_factura.Label5.Caption) = 990 Then
               Xiva = 0
            End If
         End If
         frm_factura.data_lineas.Recordset("imp_iva") = Format(Xiva, "Standard")
         If frm_factura.Label8.Caption = "" Then
            frm_factura.Label8.Caption = Format(Xiva, "Standard")
         Else
            frm_factura.Label8.Caption = CDbl(frm_factura.Label8.Caption) + Xiva
            frm_factura.Label8.Caption = Format(frm_factura.Label8.Caption, "Standard")
         End If
         If frm_factura.data_lineas.Recordset("cod_prod") = 60106 Or _
            frm_factura.data_lineas.Recordset("cod_prod") = 60108 Or _
            frm_factura.data_lineas.Recordset("cod_prod") = 994 Or _
            frm_factura.data_lineas.Recordset("cod_prod") = 993 Or _
            frm_factura.data_lineas.Recordset("cod_prod") = 996 Or _
            frm_factura.data_lineas.Recordset("cod_prod") = 60105 Or _
            frm_factura.data_lineas.Recordset("cod_prod") = 999 Or _
            frm_factura.data_lineas.Recordset("cod_prod") = 60109 Then
            frm_factura.data_lineas.Recordset("imp_iva") = 0
            frm_factura.data_lineas.Recordset("porce_est") = Xelnrodeuda
            Xiva = 0
            frm_factura.Label8.Caption = 0
            frm_factura.Label8.Caption = Format(frm_factura.Label8.Caption, "Standard")
         Else
            frm_factura.data_lineas.Recordset("porce_est") = 0
         End If
         If Xcuotabase = 9 Then
            frm_factura.data_lineas.Recordset("imp_iva") = 0
            Xiva = 0
            frm_factura.Label8.Caption = 0
            frm_factura.Label8.Caption = Format(frm_factura.Label8.Caption, "Standard")
         End If
         If frm_factura.dbcbomed.Text <> "" Then
            frm_factura.data_lineas.Recordset("nro_med_a") = frm_factura.labmed.Caption
            frm_factura.data_lineas.Recordset("nom_med_a") = frm_factura.dbcbomed.Text
         End If
         If frm_factura.dbcbomedo.Text <> "" Then
            frm_factura.data_lineas.Recordset("nro_med_s") = frm_factura.labmedo.Caption
            frm_factura.data_lineas.Recordset("nom_med_s") = frm_factura.dbcbomedo.Text
         End If
         frm_factura.data_lineas.Recordset("precio_est") = Format(frm_factura.txt_precio.Text, "Standard")
         If data_estudio.Recordset("codest") = 993 Or data_estudio.Recordset("codest") = 999 Or data_estudio.Recordset("codest") = 994 Then
            frm_factura.data_lineas.Recordset("mes_paga") = Val(frm_factura.txt_mes.Text)
            frm_factura.data_lineas.Recordset("ano_paga") = Val(frm_factura.txt_ano.Text)
         End If
         frm_factura.data_lineas.Recordset("base") = frmabm.data_parsec.Recordset("base")
         If Xfpago = 2 Then
            If frmabm.data_parsec.Recordset("base") = 20 And Val(frm_factura.txt_precio.Text) <= 0 And Val(frm_factura.Label5.Caption) = 10050 Then
               Xfpago = 1
               frm_factura.data_lineas.Recordset("tipo") = "CONTADO"
            Else
               frm_factura.data_lineas.Recordset("tipo") = "CREDITO"
            End If
         Else
            frm_factura.data_lineas.Recordset("tipo") = "CONTADO"
         End If
         If frm_factura.labtot.Caption = "" Then
         Else
            If Xiva = 0 Then
               frm_factura.data_lineas.Recordset("costo") = frm_factura.txt_precio.Text
            Else
               frm_factura.data_lineas.Recordset("costo") = frm_factura.txt_precio.Text - Xiva
            End If
         End If
         frm_factura.data_lineas.Recordset("linea") = Xcandelin
         frm_factura.data_lineas.Recordset("libro_rub") = frm_factura.Label7.Caption ' tipo de documento (Ej.e-ticket)
         frm_factura.data_lineas.Recordset("in_unid") = "INT1"
         If frm_factura.labtot.Caption <> "" Then
            frm_factura.labtot.Caption = Val(frm_factura.labtot.Caption) + Format(frm_factura.txt_precio.Text, "Standard")
            frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
         Else
            frm_factura.labtot.Caption = Format(frm_factura.txt_precio.Text, "Standard")
            frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
         End If
         If Trim(labidpedido.Caption) <> "" Then
            frm_factura.data_lineas.Recordset("nro_pedido") = Val(labidpedido.Caption)
         End If
         frm_factura.data_lineas.Recordset.Update
         frm_factura.data_lineas.Refresh
         If frm_factura.labtimemi.Caption <> "" Then
            If Val(frm_factura.labtimemi.Caption) > 0 Then
               Command4_Click
            End If
         End If
         If frm_factura.labdeudaemi.Caption <> "" Then
            If Val(frm_factura.labdeudaemi.Caption) > 0 Then
               Command5_Click
            End If
         End If
         If frm_factura.cbotim.Text = "SI" And frm_factura.Label7.Caption <> "NC E-TICKET" Then
            data_estudiobus.RecordSource = "select * from estudios where codest =" & 995
            data_estudiobus.Refresh
            If data_estudiobus.Recordset.RecordCount > 0 Then
               XValtim = data_estudiobus.Recordset("cons")
            Else
               XValtim = 85
            End If
            data_lineas.Refresh
            data_lineas.Recordset.FindFirst "cod_prod =" & 995
            If Not data_lineas.Recordset.NoMatch Then
               data_lineas.Recordset.Edit
               data_lineas.Recordset("tot_lin") = data_lineas.Recordset("tot_lin") + XValtim
               data_lineas.Recordset("cantidad") = data_lineas.Recordset("cantidad") + 1
               data_lineas.Recordset.Update
               data_lineas.Refresh
               frm_factura.labtot.Caption = Val(frm_factura.labtot.Caption) + XValtim
               frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
               frm_factura.labtim.Caption = Val(frm_factura.labtim.Caption) + XValtim
            Else
               Xcandelin = Xcandelin + 1
               frm_factura.data_lineas.Recordset.AddNew
               frm_factura.data_lineas.Recordset("reg_cab") = 0
               frm_factura.data_lineas.Recordset("factura") = 0
               frm_factura.data_lineas.Recordset("tipo_mov") = 1
               frm_factura.data_lineas.Recordset("realizada") = Format(frm_factura.mf.Text, "dd/mm/yyyy")
               frm_factura.data_lineas.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
               frm_factura.data_lineas.Recordset("cod_cli") = Val(frm_factura.labmatri.Caption)
               frm_factura.data_lineas.Recordset("nom_cli") = frm_factura.labnomb.Caption
               frm_factura.data_lineas.Recordset("cod_prod") = data_estudiobus.Recordset("codest")
               frm_factura.data_lineas.Recordset("nom_prod") = data_estudiobus.Recordset("descrip")
               If frm_factura.txt_rut.Visible = True Then
                  If Trim(frm_factura.txt_rut.Text) <> "" Then
                     frm_factura.data_lineas.Recordset("ruc") = frm_factura.txt_rut.Text
                  End If
               End If
               frm_factura.data_lineas.Recordset("cantidad") = 1
               frm_factura.data_lineas.Recordset("moneda") = "SR"
               frm_factura.data_lineas.Recordset("operador") = WElusuario
               frm_factura.data_lineas.Recordset("hora") = Format(Time, "HH:mm")
               frm_factura.data_lineas.Recordset("nro_flia") = data_estudiobus.Recordset("flia")
               frm_factura.data_lineas.Recordset("nom_flia") = data_estudiobus.Recordset("nomflia")
               frm_factura.data_lineas.Recordset("nro_superv") = frmabm.data_clientes.Recordset("cl_nro_sup")
               frm_factura.data_lineas.Recordset("nom_superv") = frmabm.data_clientes.Recordset("cl_nom_sup")
               frm_factura.data_lineas.Recordset("convenio") = frmabm.data_clientes.Recordset("cl_codconv")
               frm_factura.data_lineas.Recordset("usa_timbre") = "S"
               If Xestco = 13010 Or Xestco = 13014 Or Xestco = 13017 Or _
                  Xestco = 13017 Or Xestco = 13022 Or Xestco = 13034 Or Xestco = 13026 Or Xestfl = 3 Then
                  frm_factura.data_lineas.Recordset("rub_cont") = 211332
                  frm_factura.data_lineas.Recordset("rub_nomb") = "FERTILAB"
               Else
                  If Xestco = 13019 Or Xestco = 13021 Or Xestco = 13024 Or _
                     Xestco = 13029 Or Xestco = 13037 Or Xestfl = 5 Then
                     frm_factura.data_lineas.Recordset("rub_cont") = 211587
                     frm_factura.data_lineas.Recordset("rub_nomb") = "ECOGRAFISTA"
                  Else
                     If Xestco = 13009 Or Xestco = 13013 Or Xestco = 13023 Or _
                        Xestco = 13027 Or Xestco = 13035 Or Xestfl = 7 Then
                        frm_factura.data_lineas.Recordset("rub_cont") = 211313
                        frm_factura.data_lineas.Recordset("rub_nomb") = "RADIOLOGOS"
                     Else
                        If Xestco = 13012 Or Xestco = 13016 Or Xestco = 13038 Or _
                           Xestfl = 11 Then
                           frm_factura.data_lineas.Recordset("rub_cont") = 211586
                           frm_factura.data_lineas.Recordset("rub_nomb") = "SERV.CARDIOLOGICOS"
                        Else
                           frm_factura.data_lineas.Recordset("rub_cont") = 213076
                           frm_factura.data_lineas.Recordset("rub_nomb") = "OBL.TIMBRES PROF."
                        End If
                     End If
                  End If
               End If
               frm_factura.data_lineas.Recordset("arancel") = XValtim
               frm_factura.data_lineas.Recordset("tot_lin") = XValtim
               If frm_factura.labtot.Caption <> "" Then
                  frm_factura.labtot.Caption = Val(frm_factura.labtot.Caption) + XValtim
                  frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
               Else
                  frm_factura.labtot.Caption = XValtim
                  frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
               End If
               frm_factura.data_lineas.Recordset("precio_est") = XValtim
               frm_factura.data_lineas.Recordset("porce_est") = 0
               frm_factura.data_lineas.Recordset("base") = frmabm.data_parsec.Recordset("base")
               If Xfpago = 2 Then
                  frm_factura.data_lineas.Recordset("tipo") = "CREDITO"
               Else
                  frm_factura.data_lineas.Recordset("tipo") = "CONTADO"
               End If
               frm_factura.data_lineas.Recordset("linea") = Xcandelin
               frm_factura.data_lineas.Recordset("libro_rub") = frm_factura.Label7.Caption ' tipo de documento (Ej.e-ticket)
               frm_factura.data_lineas.Recordset("in_unid") = "INT1"
               frm_factura.data_lineas.Recordset.Update
               frm_factura.labtim.Caption = Val(frm_factura.labtim.Caption) + XValtim
               frm_factura.data_lineas.Refresh
               data_lineas.Refresh
               If data_estudio.Recordset("codest") = 997 Then
                  Xcoddeu = 9
               End If
            End If
         End If
      End If
   Else ' aca es medicacion y más
      Xcandelin = Xcandelin + 1
      frm_factura.data_lineas.Recordset.AddNew
      If frm_factura.Label7.Caption = "NC E-TICKET" Or frm_factura.Label7.Caption = "NC E-FACTURA" Or frm_factura.Label7.Caption = "ND E-FACTURA" Or frm_factura.Label7.Caption = "ND E-TICKET" Then
         ''Xmotivoref = InputBox("Ingrese motivo de la modificación")
         Xmotivoref = labmotivoref.Caption
         If Xmotivoref <> "" Then
            frm_factura.labmotivo.Caption = Xmotivoref
            If frm_factura.labseriecance.Caption = "XX" Then
               frm_factura.data_lincance.RecordSource = "Select * from linmmdd where factura =" & Val(frm_factura.labfaccance.Caption) & " and fecha =#" & Format(frm_factura.labfeccance.Caption, "yyyy-mm-dd") & "#"
               frm_factura.data_lincance.Refresh
            Else
               If frm_factura.labidemi.Caption = 5 Then
                  frm_factura.data_lincance.RecordSource = "Select * from deudas where documento =" & Val(frm_factura.labfaccance.Caption)
                  frm_factura.data_lincance.Refresh
               Else
                  frm_factura.data_lincance.RecordSource = "Select * from clirespl where cl_numero =" & Val(frm_factura.labfaccance.Caption) & " and cl_socmnro ='" & Trim(frm_factura.labseriecance.Caption) & "' and cl_fnac =#" & Format(frm_factura.labfeccance.Caption, "yyyy-mm-dd") & "#"
                  frm_factura.data_lincance.Refresh
               End If
            End If
            If frm_factura.data_lincance.Recordset.RecordCount > 0 And Val(frm_factura.labidemi.Caption) <> 5 Then
               If frm_factura.labseriecance.Caption <> "XX" Then
                  If Val(frm_factura.labtot.Caption) >= frm_factura.data_lincance.Recordset("saldo_doc") Then
                     MsgBox "Ha excedido el importe de la factura", vbCritical
''                     b_cance_Click
                  End If
               End If
               If frm_factura.labseriecance.Caption = "XX" Then
                  frm_factura.data_lineas.Recordset("tipodocref") = 101
               Else
                  frm_factura.data_lineas.Recordset("tipodocref") = frm_factura.data_lincance.Recordset("cl_tipocli")
               End If
               frm_factura.data_lineas.Recordset("serieref") = frm_factura.labseriecance.Caption
               If Len(frm_factura.labfaccance.Caption) > 7 Then
                  frm_factura.data_lineas.Recordset("nrofactref") = Val(Mid(frm_factura.labfaccance.Caption, 1, 7))
               Else
                  frm_factura.data_lineas.Recordset("nrofactref") = Val(frm_factura.labfaccance.Caption)
               End If
               If frm_factura.labseriecance.Caption = "XX" Then
                  frm_factura.data_lineas.Recordset("fechafact") = CDate(frm_factura.labfeccance.Caption)
               Else
                  frm_factura.data_lineas.Recordset("fechafact") = frm_factura.data_lincance.Recordset("cl_fnac")
               End If
               frm_factura.data_lineas.Recordset("motivoref") = Mid(Xmotivoref, 1, 90)
               frm_factura.data_lineas.Recordset("linearef") = Val(frm_factura.lablinea.Caption)
            Else
               If frm_factura.labidemi.Caption = 5 Then
                  If frm_factura.Label7.Caption = "NC E-FACTURA" Then
                     frm_factura.data_lineas.Recordset("tipodocref") = 111
                  Else
                     frm_factura.data_lineas.Recordset("tipodocref") = 101
                  End If
                  frm_factura.data_lineas.Recordset("serieref") = frm_factura.labseriecance.Caption
                  If Len(frm_factura.labfaccance.Caption) > 7 Then
                     frm_factura.data_lineas.Recordset("nrofactref") = Val(Mid(frm_factura.labfaccance.Caption, 1, 7))
                  Else
                     frm_factura.data_lineas.Recordset("nrofactref") = Val(frm_factura.labfaccance.Caption)
                  End If
                  frm_factura.data_lineas.Recordset("fechafact") = CDate(frm_factura.labfeccance.Caption)
                  frm_factura.data_lineas.Recordset("motivoref") = Mid(Xmotivoref, 1, 90)
                  frm_factura.data_lineas.Recordset("linearef") = 1
               Else
                  MsgBox "No se encuentra número de factura a cancelar"
''                  b_cance_Click
               End If
            End If
         Else
            MsgBox "No ingresó motivo de cancelación"
            ''''b_cance_Click
         End If
''''       btn_fin.SetFocus
      Else
         If frm_factura.Label7.Caption = "DEV.RECIBO" Then
            If frm_factura.labfaccance.Caption <> "" Then
               frm_factura.data_lineas.Recordset("nrofactref") = Val(frm_factura.labfaccance.Caption)
            End If
         End If
      End If
      If frm_factura.labcodpro.Caption <> "" Then
         frm_factura.data_lineas.Recordset("numero") = Val(frm_factura.labcodpro.Caption)
      End If
      frm_factura.data_lineas.Recordset("reg_cab") = 99
      frm_factura.data_lineas.Recordset("factura") = 0
      If frm_factura.txt_precio.Text <> "" Then
         If frm_factura.txt_precio.Text > 0 Then
            If Val(frm_factura.Label5.Caption) = 995 Or Val(frm_factura.Label5.Caption) = 990 Then
               frm_factura.data_lineas.Recordset("tipo_mov") = 1 'tipo de iva (indic de fact)
            Else
               frm_factura.data_lineas.Recordset("tipo_mov") = 2 'tipo de iva (indic de fact)
            End If
         Else
            frm_factura.data_lineas.Recordset("tipo_mov") = 5
         End If
      Else
         frm_factura.data_lineas.Recordset("tipo_mov") = 5
      End If
      frm_factura.data_lineas.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
      frm_factura.data_lineas.Recordset("fecha") = Format(frm_factura.mf.Text, "dd/mm/yyyy")
      frm_factura.data_lineas.Recordset("cod_cli") = frm_factura.labmatri.Caption
      frm_factura.data_lineas.Recordset("nom_cli") = Mid(frm_factura.labnomb.Caption, 1, 30)
      frm_factura.data_lineas.Recordset("cod_prod") = data_estudio.Recordset("codest")
      frm_factura.data_lineas.Recordset("nom_prod") = Mid(data_estudio.Recordset("descrip"), 1, 50)
      frm_factura.data_lineas.Recordset("cantidad") = 1
      frm_factura.data_lineas.Recordset("moneda") = "SR" 'Serie
      If frm_factura.txt_rut.Visible = True Then
         If Trim(frm_factura.txt_rut.Text) <> "" Then
''''            Command2_Click
            frm_factura.data_lineas.Recordset("ruc") = frm_factura.txt_rut.Text
         End If
      End If
      If data_estudio.Recordset("flia") = 3 Then
         frm_factura.data_lineas.Recordset("tcambio") = 8
      End If
      If data_estudio.Recordset("codest") = 60103 Or _
         data_estudio.Recordset("codest") = 60105 Or _
         data_estudio.Recordset("codest") = 60106 Or _
         data_estudio.Recordset("codest") = 60107 Or _
         data_estudio.Recordset("codest") = 60108 Or _
         data_estudio.Recordset("codest") = 60109 Then
         If data_estudio.Recordset("codest") <> 60107 Then
            If data_estudio.Recordset("codest") <> 60103 Then
               frm_factura.data_lineas.Recordset("tipo_mov") = 1
            End If
         End If
         frm_factura.data_lineas.Recordset("codelmedic") = XcodelMedicamento
         frm_factura.data_lineas.Recordset("idtablapres") = IdTablaPres
         If Xdescmedic <> "" Then
            frm_factura.data_lineas.Recordset("nom_medic") = Mid(UCase(Xdescmedic), 1, 50)
         End If
         If frm_factura.txt_precio.Text <> "" Then
            If frm_factura.txt_precio.Text < 0 Then
               frm_factura.data_lineas.Recordset("dias") = 1
            Else
               frm_factura.data_lineas.Recordset("dias") = 0
            End If
         Else
            frm_factura.data_lineas.Recordset("dias") = 0
         End If
      End If
      frm_factura.data_lineas.Recordset("operador") = WElusuario
      frm_factura.data_lineas.Recordset("hora") = Format(Time, "HH:mm")
      frm_factura.data_lineas.Recordset("nro_flia") = data_estudio.Recordset("flia")
      frm_factura.data_lineas.Recordset("nom_flia") = data_estudio.Recordset("nomflia")
      frm_factura.data_lineas.Recordset("nro_superv") = frmabm.data_clientes.Recordset("cl_nro_sup")
      frm_factura.data_lineas.Recordset("nom_superv") = frmabm.data_clientes.Recordset("cl_nom_sup")
      frm_factura.data_lineas.Recordset("convenio") = frmabm.data_clientes.Recordset("cl_codconv")
      frm_factura.data_lineas.Recordset("ced_socio") = frmabm.data_clientes.Recordset("cl_cedula")
      frm_factura.data_lineas.Recordset("fact") = frmabm.data_clientes.Recordset("cl_codced")
      If data_estudio.Recordset("codest") = 997 Then
         frm_factura.data_lineas.Recordset("rub_cont") = 113003
         frm_factura.data_lineas.Recordset("tipo_mov") = 1
         Xrub = 113003
      Else
         If data_estudio.Recordset("codest") = 993 Or _
            data_estudio.Recordset("codest") = 994 Then
            frm_factura.data_lineas.Recordset("rub_cont") = 112022
            frm_factura.data_lineas.Recordset("tipo_mov") = 1
            Xrub = 112022
         Else
            If data_estudio.Recordset("codest") = 999 Then
               frm_factura.data_lineas.Recordset("tipo_mov") = 1
               If IsNull(frmabm.data_clientes.Recordset("cl_nrocobr")) = False Then
                  If frmabm.data_clientes.Recordset("cl_nrocobr") = 615 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 616 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 635 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 602 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 113 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 653 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 672 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 1 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 10 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 201 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 512 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 636 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 685 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 208 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 209 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 8 Or _
                     frmabm.data_clientes.Recordset("cl_nrocobr") = 0 Then
                     frm_factura.data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("cbrcuo")
                     Xrub = frmabm.data_parsec.Recordset("cbrcuo")
                     frm_factura.data_lineas.Recordset("grupo") = frmabm.data_clientes.Recordset("cl_nrocobr")
                  Else
                     frm_factura.data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("cbrcuo")
                     Xrub = frmabm.data_parsec.Recordset("cbrcuo")
                     frm_factura.data_lineas.Recordset("grupo") = frmabm.data_clientes.Recordset("cl_nrocobr")
                     Xcuotabase = 9
                  End If
               Else
                  frm_factura.data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("cbrcuo")
                  Xrub = frmabm.data_parsec.Recordset("cbrcuo")
                  frm_factura.data_lineas.Recordset("grupo") = frmabm.data_clientes.Recordset("cl_nrocobr")
               End If
            Else
               If data_estudio.Recordset("codest") = 992 Then
                  frm_factura.data_lineas.Recordset("rub_cont") = 513007
                  Xrub = 513007
               Else
                  If Xfpago = 2 Then
                     frm_factura.data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("srvcrd")
                     Xrub = frmabm.data_parsec.Recordset("srvcrd")
                  Else
                     frm_factura.data_lineas.Recordset("rub_cont") = frmabm.data_parsec.Recordset("srvcnt")
                     Xrub = frmabm.data_parsec.Recordset("srvcnt")
                  End If
               End If
            End If
         End If
      End If
      If data_estudio.Recordset("codest") = 996 Then
         frm_factura.data_lineas.Recordset("rub_cont") = 211473
         frm_factura.data_lineas.Recordset("tipo_mov") = 1
         Xrub = 211473
      End If
      data_codcaja.Recordset.FindFirst "numero =" & Xrub
      If Not data_codcaja.Recordset.NoMatch Then
         frm_factura.data_lineas.Recordset("rub_nomb") = data_codcaja.Recordset("nombre")
      End If
      If data_estudio.Recordset("codest") = 60106 Then
         frm_factura.data_lineas.Recordset("rub_cont") = 211397
         frm_factura.data_lineas.Recordset("rub_nomb") = "M. SMI"
      End If
      If data_estudio.Recordset("codest") = 60105 Then
         frm_factura.data_lineas.Recordset("rub_cont") = 211397
         frm_factura.data_lineas.Recordset("rub_nomb") = "M. SMI"
      End If
      If data_estudio.Recordset("codest") = 60108 Then
         frm_factura.data_lineas.Recordset("rub_cont") = 211302
         frm_factura.data_lineas.Recordset("rub_nomb") = "M.UNIVERSAL"
      End If
      If (Xestco = 60107 Or Xestco = 60103) And frm_factura.Label7.Caption = "E-TICKET" Then
         If frm_factura.txt_precio.Text > 0 Then
            data_estudiobus.RecordSource = "Select * from estudios where codest =" & 990
            data_estudiobus.Refresh
            If data_estudiobus.Recordset.RecordCount > 0 Then
               Xvaltimme = data_estudiobus.Recordset("cons")
               If frm_factura.labtimme.Caption = "" Then
                  frm_factura.labtimme.Caption = Xvaltimme
               Else
                  frm_factura.labtimme.Caption = Val(frm_factura.labtimme.Caption) + Xvaltimme
               End If
            Else
               Xvaltimme = 18
               If frm_factura.labtimme.Caption = "" Then
                  frm_factura.labtimme.Caption = Xvaltimme
               Else
                  frm_factura.labtimme.Caption = Val(frm_factura.labtimme.Caption) + Xvaltimme
               End If
            End If
            If CDbl(frm_factura.txt_precio.Text) >= CDbl(Xvaltimme) Then
               frm_factura.data_lineas.Recordset("arancel") = Format(frm_factura.txt_precio.Text - Xvaltimme, "Standard")
               frm_factura.data_lineas.Recordset("tot_lin") = Format(frm_factura.txt_precio.Text - Xvaltimme, "Standard")
            Else
               MsgBox "Importe menor al timbre, VERIFIQUE!!!", vbCritical
'               frmabm.btn_fact.Enabled = True
'               Unload Me
'               Exit Sub
               Xvaltimme = 0
               frm_factura.labtimme.Caption = 0
               frm_factura.data_lineas.Recordset("arancel") = Format(frm_factura.txt_precio.Text, "Standard")
               frm_factura.data_lineas.Recordset("tot_lin") = Format(frm_factura.txt_precio.Text, "Standard")
            End If
         Else
            frm_factura.data_lineas.Recordset("arancel") = Format(frm_factura.txt_precio.Text, "Standard")
            frm_factura.data_lineas.Recordset("tot_lin") = Format(frm_factura.txt_precio.Text, "Standard")
         End If
      Else
         frm_factura.data_lineas.Recordset("arancel") = Format(frm_factura.txt_precio.Text, "Standard")
         frm_factura.data_lineas.Recordset("tot_lin") = Format(frm_factura.txt_precio.Text, "Standard")
      End If
        'hasta aqui
'''      Xiva = data_lineas.Recordset("tot_lin") / 1.1
      Xiva = Val(frm_factura.txt_precio.Text) / 1.1
      Xiva = Xiva * 0.1
      If frm_factura.Label5.Caption <> "" Then
         If Val(frm_factura.Label5.Caption) = 995 Or Val(frm_factura.Label5.Caption) = 990 Then
            Xiva = 0
         End If
      End If
      frm_factura.data_lineas.Recordset("imp_iva") = Format(Xiva, "Standard")
      If frm_factura.Label8.Caption = "" Then
         frm_factura.Label8.Caption = Format(Xiva, "Standard")
      Else
         frm_factura.Label8.Caption = CDbl(frm_factura.Label8.Caption) + Xiva
         frm_factura.Label8.Caption = Format(frm_factura.Label8.Caption, "Standard")
      End If
      If frm_factura.data_lineas.Recordset("cod_prod") = 60106 Or _
         frm_factura.data_lineas.Recordset("cod_prod") = 60108 Or _
         frm_factura.data_lineas.Recordset("cod_prod") = 994 Or _
         frm_factura.data_lineas.Recordset("cod_prod") = 993 Or _
         frm_factura.data_lineas.Recordset("cod_prod") = 996 Or _
         frm_factura.data_lineas.Recordset("cod_prod") = 60105 Or _
         frm_factura.data_lineas.Recordset("cod_prod") = 60109 Or _
         frm_factura.data_lineas.Recordset("cod_prod") = 999 Or _
         frm_factura.data_lineas.Recordset("cod_prod") = 80011 Or _
         frm_factura.data_lineas.Recordset("cod_prod") = 80012 Or _
         frm_factura.data_lineas.Recordset("cod_prod") = 80013 Or _
         frm_factura.data_lineas.Recordset("cod_prod") = 80014 Or _
         frm_factura.data_lineas.Recordset("cod_prod") = 80016 Or _
         frm_factura.data_lineas.Recordset("cod_prod") = 80015 Then
         frm_factura.data_lineas.Recordset("imp_iva") = 0
         Xiva = 0
         frm_factura.Label8.Caption = 0
         frm_factura.Label8.Caption = Format(frm_factura.Label8.Caption, "Standard")
      End If
      If Xcuotabase = 9 Then
         frm_factura.data_lineas.Recordset("imp_iva") = 0
         Xiva = 0
         frm_factura.Label8.Caption = 0
         frm_factura.Label8.Caption = Format(frm_factura.Label8.Caption, "Standard")
      End If
      If frm_factura.dbcbomed.Text <> "" Then
         frm_factura.data_lineas.Recordset("nro_med_a") = frm_factura.labmed.Caption
         frm_factura.data_lineas.Recordset("nom_med_a") = frm_factura.dbcbomed.Text
      End If
      If frm_factura.dbcbomedo.Text <> "" Then
         frm_factura.data_lineas.Recordset("nro_med_s") = frm_factura.labmedo.Caption
         frm_factura.data_lineas.Recordset("nom_med_s") = frm_factura.dbcbomedo.Text
      End If
      If (Xestco = 60107 Or Xestco = 60103) And frm_factura.Label7.Caption = "E-TICKET" Then
         If frm_factura.txt_precio.Text > 0 Then
            If Val(frm_factura.txt_precio.Text) >= Xvaltimme Then
               frm_factura.data_lineas.Recordset("precio_est") = Format(frm_factura.txt_precio.Text - Xvaltimme, "Standard")
            Else
               frm_factura.data_lineas.Recordset("precio_est") = Format(frm_factura.txt_precio.Text, "Standard")
            End If
         Else
            frm_factura.data_lineas.Recordset("precio_est") = Format(frm_factura.txt_precio.Text, "Standard")
         End If
      Else
         frm_factura.data_lineas.Recordset("precio_est") = Format(frm_factura.txt_precio.Text, "Standard")
      End If
      If Xestco = 993 Or Xestco = 999 Or Xestco = 994 Then
         frm_factura.data_lineas.Recordset("mes_paga") = Val(frm_factura.txt_mes.Text)
         frm_factura.data_lineas.Recordset("ano_paga") = Val(frm_factura.txt_ano.Text)
         frm_factura.data_lineas.Recordset("porce_est") = Xelnrodeuda
      Else
         frm_factura.data_lineas.Recordset("porce_est") = 0
      End If
      frm_factura.data_lineas.Recordset("base") = frmabm.data_parsec.Recordset("base")
      If frm_factura.labtot.Caption = "" Then
      Else
         If (Xestco = 60107 Or Xestco = 60103) And frm_factura.Label7.Caption = "E-TICKET" Then
            If frm_factura.txt_precio.Text > 0 Then
               If Val(frm_factura.txt_precio.Text) >= Xvaltimme Then
                  If Xiva = 0 Then
                     frm_factura.data_lineas.Recordset("costo") = frm_factura.txt_precio.Text - Xvaltimme
                  Else
                     frm_factura.data_lineas.Recordset("costo") = frm_factura.txt_precio.Text - Xvaltimme - Xiva
                  End If
               Else
                  If Xiva = 0 Then
                     frm_factura.data_lineas.Recordset("costo") = frm_factura.txt_precio.Text
                  Else
                     frm_factura.data_lineas.Recordset("costo") = frm_factura.txt_precio.Text - Xiva
                  End If
               End If
            Else
               If Xiva = 0 Then
                  frm_factura.data_lineas.Recordset("costo") = frm_factura.txt_precio.Text
               Else
                  frm_factura.data_lineas.Recordset("costo") = frm_factura.txt_precio.Text - Xiva
               End If
            End If
         Else
            If Xiva = 0 Then
               frm_factura.data_lineas.Recordset("costo") = frm_factura.txt_precio.Text
            Else
               frm_factura.data_lineas.Recordset("costo") = frm_factura.txt_precio.Text - Xiva
            End If
         End If
      End If
      If Xfpago = 2 Then
         frm_factura.data_lineas.Recordset("tipo") = "CREDITO"
      Else
         frm_factura.data_lineas.Recordset("tipo") = "CONTADO"
      End If
      If frm_factura.labtot.Caption <> "" Then
         frm_factura.labtot.Caption = Val(frm_factura.labtot.Caption) + Val(frm_factura.txt_precio.Text)
         frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
      Else
         frm_factura.labtot.Caption = Format(frm_factura.txt_precio.Text, "Standard")
         frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
      End If
      frm_factura.data_lineas.Recordset("linea") = Xcandelin
      frm_factura.data_lineas.Recordset("libro_rub") = frm_factura.Label7.Caption ' tipo de documento (Ej.e-ticket)
      frm_factura.data_lineas.Recordset("in_unid") = "INT1"
      If Trim(labidpedido.Caption) <> "" Then
         frm_factura.data_lineas.Recordset("nro_pedido") = Val(labidpedido.Caption)
      End If
      frm_factura.data_lineas.Recordset.Update
      frm_factura.data_lineas.Refresh
      If frm_factura.labtimemi.Caption <> "" Then
         If Format(frm_factura.labtimemi.Caption, "Standard") > 0 Then
            Command4_Click
         End If
      End If
      If frm_factura.labdeudaemi.Caption <> "" Then
         If Format(frm_factura.labdeudaemi.Caption, "Standard") > 0 Then
            Command5_Click
         End If
      End If
      If (Xestco = 60107 Or Xestco = 60103) And frm_factura.txt_precio.Text > 0 And frm_factura.Label7.Caption = "E-TICKET" Then
         data_lineas.Refresh
         data_lineas.Recordset.FindFirst "cod_prod =" & 990
         If Not data_lineas.Recordset.NoMatch Then
            data_lineas.Recordset.Edit
            data_lineas.Recordset("tot_lin") = data_lineas.Recordset("tot_lin") + Xvaltimme
            data_lineas.Recordset("cantidad") = data_lineas.Recordset("cantidad") + 1
            data_lineas.Recordset.Update
            frm_factura.labtim.Caption = Val(frm_factura.labtim.Caption) + Xvaltimme
         Else
            Xcandelin = Xcandelin + 1
            frm_factura.data_lineas.Recordset.AddNew
            frm_factura.data_lineas.Recordset("reg_cab") = 0
            frm_factura.data_lineas.Recordset("factura") = 0
            frm_factura.data_lineas.Recordset("tipo_mov") = 1
            frm_factura.data_lineas.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
            frm_factura.data_lineas.Recordset("fecha") = Format(frm_factura.mf.Text, "dd/mm/yyyy")
            frm_factura.data_lineas.Recordset("cod_cli") = frm_factura.labmatri.Caption
            frm_factura.data_lineas.Recordset("nom_cli") = frm_factura.labnomb.Caption
            frm_factura.data_lineas.Recordset("cod_prod") = 990
            frm_factura.data_lineas.Recordset("nom_prod") = "TIMBRES PROFESIONAL M"
            frm_factura.data_lineas.Recordset("usa_timbre") = "M"
            frm_factura.data_lineas.Recordset("moneda") = "SR" 'Serie
            If frm_factura.txt_rut.Visible = True Then
               If Trim(frm_factura.txt_rut.Text) <> "" Then
                    ''''Command2_Click
                  frm_factura.data_lineas.Recordset("ruc") = frm_factura.txt_rut.Text
               End If
            End If
            frm_factura.data_lineas.Recordset("cantidad") = 1
            frm_factura.data_lineas.Recordset("operador") = WElusuario
            frm_factura.data_lineas.Recordset("hora") = Format(Time, "HH:mm")
            frm_factura.data_lineas.Recordset("nro_flia") = 8
            frm_factura.data_lineas.Recordset("nom_flia") = "OTROS SERVICIOS"
            frm_factura.data_lineas.Recordset("nro_superv") = frmabm.data_clientes.Recordset("cl_nro_sup")
            frm_factura.data_lineas.Recordset("nom_superv") = frmabm.data_clientes.Recordset("cl_nom_sup")
            frm_factura.data_lineas.Recordset("convenio") = frmabm.data_clientes.Recordset("cl_codconv")
            frm_factura.data_lineas.Recordset("rub_cont") = 213076
            frm_factura.data_lineas.Recordset("rub_nomb") = "OBL.TIMBRES PROF."
            frm_factura.data_lineas.Recordset("arancel") = Xvaltimme
            frm_factura.data_lineas.Recordset("tot_lin") = Xvaltimme
            frm_factura.data_lineas.Recordset("precio_est") = Xvaltimme
            frm_factura.data_lineas.Recordset("porce_est") = 0
            frm_factura.data_lineas.Recordset("base") = frmabm.data_parsec.Recordset("base")
            If Xfpago = 2 Then
               frm_factura.data_lineas.Recordset("tipo") = "CREDITO"
            Else
               frm_factura.data_lineas.Recordset("tipo") = "CONTADO"
            End If
            frm_factura.data_lineas.Recordset("linea") = Xcandelin
            frm_factura.data_lineas.Recordset("libro_rub") = frm_factura.Label7.Caption ' tipo de documento (Ej.e-ticket)
            frm_factura.data_lineas.Recordset("in_unid") = "INT1"
            frm_factura.data_lineas.Recordset.Update
            frm_factura.data_lineas.Refresh
            frm_factura.labtim.Caption = Xvaltimme
         End If
      End If
      If data_estudio.Recordset("codest") = 997 Then
         Xcoddeu = 9
      End If
      If frm_factura.cbotim.Text = "SI" And frm_factura.Label7.Caption <> "NC E-TICKET" Then
         data_estudiobus.RecordSource = "select * from estudios where codest =" & 995
         data_estudiobus.Refresh
         If data_estudiobus.Recordset.RecordCount > 0 Then
            XValtim = data_estudiobus.Recordset("cons")
         Else
            XValtim = 59
         End If
         data_lineas.Refresh
         data_lineas.Recordset.FindFirst "cod_prod =" & 995
         If Not data_lineas.Recordset.NoMatch Then
            data_lineas.Recordset.Edit
            data_lineas.Recordset("tot_lin") = data_lineas.Recordset("tot_lin") + XValtim
            data_lineas.Recordset("cantidad") = data_lineas.Recordset("cantidad") + 1
            data_lineas.Recordset.Update
            If frm_factura.labtot.Caption <> "" Then
               frm_factura.labtot.Caption = Val(frm_factura.labtot.Caption) + XValtim
               frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
               frm_factura.labtim.Caption = Val(frm_factura.labtim.Caption) + XValtim
            Else
               frm_factura.labtot.Caption = XValtim
               frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
               frm_factura.labtim.Caption = Val(frm_factura.labtim.Caption) + XValtim
            End If
         Else
            Xcandelin = Xcandelin + 1
            frm_factura.data_lineas.Recordset.AddNew
            frm_factura.data_lineas.Recordset("reg_cab") = 0
            frm_factura.data_lineas.Recordset("factura") = 0
            frm_factura.data_lineas.Recordset("tipo_mov") = 1
            frm_factura.data_lineas.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
            frm_factura.data_lineas.Recordset("fecha") = Format(frm_factura.mf.Text, "dd/mm/yyyy")
            frm_factura.data_lineas.Recordset("cod_cli") = frm_factura.labmatri.Caption
            frm_factura.data_lineas.Recordset("nom_cli") = frm_factura.labnomb.Caption
            frm_factura.data_lineas.Recordset("cod_prod") = data_estudiobus.Recordset("codest")
            frm_factura.data_lineas.Recordset("nom_prod") = data_estudiobus.Recordset("descrip")
            If frm_factura.txt_rut.Visible = True Then
               If Trim(frm_factura.txt_rut.Text) <> "" Then
                  ''Command2_Click
                  frm_factura.data_lineas.Recordset("ruc") = frm_factura.txt_rut.Text
               End If
            End If
            frm_factura.data_lineas.Recordset("cantidad") = 1
            frm_factura.data_lineas.Recordset("moneda") = "SR"
            frm_factura.data_lineas.Recordset("operador") = WElusuario
            frm_factura.data_lineas.Recordset("hora") = Format(Time, "HH:mm")
            frm_factura.data_lineas.Recordset("nro_flia") = data_estudiobus.Recordset("flia")
            frm_factura.data_lineas.Recordset("nom_flia") = data_estudiobus.Recordset("nomflia")
            frm_factura.data_lineas.Recordset("nro_superv") = frmabm.data_clientes.Recordset("cl_nro_sup")
            frm_factura.data_lineas.Recordset("nom_superv") = frmabm.data_clientes.Recordset("cl_nom_sup")
            frm_factura.data_lineas.Recordset("convenio") = frmabm.data_clientes.Recordset("cl_codconv")
            frm_factura.data_lineas.Recordset("usa_timbre") = "S"
            If Xestco = 13010 Or Xestco = 13014 Or Xestco = 13017 Or _
               Xestco = 13017 Or Xestco = 13022 Or Xestco = 13034 Or Xestco = 13026 Or Xestfl = 3 Then
               frm_factura.data_lineas.Recordset("rub_cont") = 211332
               frm_factura.data_lineas.Recordset("rub_nomb") = "FERTILAB"
            Else
               If Xestco = 13019 Or Xestco = 13021 Or Xestco = 13024 Or _
                  Xestco = 13029 Or Xestco = 13037 Or Xestfl = 5 Then
                  frm_factura.data_lineas.Recordset("rub_cont") = 211587
                  frm_factura.data_lineas.Recordset("rub_nomb") = "ECOGRAFISTA"
               Else
                  If Xestco = 13009 Or Xestco = 13013 Or Xestco = 13023 Or _
                     Xestco = 13027 Or Xestco = 13035 Or Xestfl = 7 Then
                     frm_factura.data_lineas.Recordset("rub_cont") = 211313
                     frm_factura.data_lineas.Recordset("rub_nomb") = "RADIOLOGOS"
                  Else
                     If Xestco = 13012 Or Xestco = 13016 Or Xestco = 13038 Or _
                        Xestfl = 11 Then
                        frm_factura.data_lineas.Recordset("rub_cont") = 211586
                        frm_factura.data_lineas.Recordset("rub_nomb") = "SERV.CARDIOLOGICOS"
                     Else
                        frm_factura.data_lineas.Recordset("rub_cont") = 213076
                        frm_factura.data_lineas.Recordset("rub_nomb") = "OBL.TIMBRES PROF."
                     End If
                  End If
               End If
            End If
            frm_factura.data_lineas.Recordset("arancel") = XValtim
            frm_factura.data_lineas.Recordset("tot_lin") = XValtim
            If frm_factura.labtot.Caption <> "" Then
               frm_factura.labtot.Caption = Val(frm_factura.labtot.Caption) + XValtim
               frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
            Else
               frm_factura.labtot.Caption = XValtim
               frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
            End If
            frm_factura.data_lineas.Recordset("precio_est") = XValtim
            frm_factura.data_lineas.Recordset("porce_est") = 0
            frm_factura.data_lineas.Recordset("base") = frmabm.data_parsec.Recordset("base")
            If Xfpago = 2 Then
               frm_factura.data_lineas.Recordset("tipo") = "CREDITO"
            Else
               frm_factura.data_lineas.Recordset("tipo") = "CONTADO"
            End If
            frm_factura.data_lineas.Recordset("linea") = Xcandelin
            frm_factura.data_lineas.Recordset("libro_rub") = frm_factura.Label7.Caption ' tipo de documento (Ej.e-ticket)
            frm_factura.data_lineas.Recordset("in_unid") = "INT1"
            frm_factura.data_lineas.Recordset.Update
            frm_factura.data_lineas.Refresh
            frm_factura.labtim.Caption = Val(frm_factura.labtim.Caption) + XValtim
            If data_estudiobus.Recordset("codest") = 997 Then
               Xcoddeu = 9
            End If
         End If
      End If
   End If
Else
   MsgBox "Hay un error en la selección, verifique!", vbCritical, "FACTURAR"
   Xquerr = 0
   frm_factura.labtimme.Caption = ""
   frm_factura.labmed.Caption = ""
   frm_factura.txt_precio.Text = 0
   frm_factura.cbotim.ListIndex = 0
   frm_factura.dbcbomed.Text = ""
   frm_factura.DBCombo1.Text = ""
   frm_factura.txt_mes.Text = ""
   frm_factura.txt_ano.Text = ""
   frm_factura.Label5.Caption = ""

End If

End Sub

Private Sub charq_Click()
If charq.Value = 1 Then
   DBGrid1.Visible = False
   DBGrid4.Visible = True
   DBGrid3.Visible = False
   Check1.Visible = False
   charq.Visible = True
   Check2.Visible = False
Else
   DBGrid1.Visible = False
   DBGrid4.Visible = False
   DBGrid3.Visible = True
   Check1.Visible = False
   Check2.Visible = True
   Check2.Value = 1
   charq.Visible = False
End If


End Sub

Private Sub Check1_Click()

If Check1.Value = 1 Then
   Verificar_codigo
Else
   DBGrid1.Enabled = True
   t_nrofact.Text = ""
   t_nrofact.Visible = False
   t_nrolin.Text = ""
   t_nrolin.Visible = False
   mf.Text = "__/__/____"
   mf.Visible = False
   Label3.Visible = False
   Label4.Visible = False
   Label5.Visible = False
   Command1.Visible = False
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   DBGrid1.Visible = False
   DBGrid4.Visible = False
   DBGrid3.Visible = True
   Check1.Visible = False
   charq.Visible = True
Else
   DBGrid1.Visible = True
   DBGrid4.Visible = False
   DBGrid3.Visible = False
   Check1.Visible = True
   charq.Visible = False
End If

End Sub

Private Sub Command1_Click()
If Xfaccancerecep = 1 Then
   If mf.Text <> "__/__/____" Then
      frm_factura.labfeccance.Caption = mf.Text
   Else
      frm_factura.labfeccance.Caption = "01/06/2016"
   End If
   If t_nrolin.Text <> "" Then
      frm_factura.lablinea.Caption = t_nrolin.Text
   Else
      frm_factura.lablinea.Caption = 1
   End If
   If t_nrofact.Text <> "" Then
      frm_factura.labfaccance.Caption = t_nrofact.Text
   Else
      frm_factura.labfaccance.Caption = 2525252
   End If
   frm_factura.labseriecance.Caption = "XX"
   frm_factura.DBCombo1.Enabled = True
   frm_factura.txt_precio.Enabled = True
   frm_factura.labidemi.Caption = 0
   frm_factura.labdeudaemi.Caption = 0
   frm_factura.labtimemi.Caption = 0

Else
   If t_nrolin.Text <> "" Then
      frm_factconve22.labnrolin.Caption = t_nrolin.Text
   Else
      frm_factconve22.labnrolin.Caption = 1
   End If
   If t_nrofact.Text <> "" Then
      frm_factconve22.labnrofactnc.Caption = t_nrofact.Text
   Else
      frm_factconve22.labnrofactnc.Caption = 0
   End If
   frm_factconve22.labseriefactnc.Caption = "XX"
   If mf.Text <> "__/__/____" Then
      frm_factconve22.labfeccance.Caption = mf.Text
   Else
      frm_factconve22.labfeccance.Caption = "01/01/2016"
   End If
'   frm_factura.labidemi.Caption = 0
'   frm_factura.labdeudaemi.Caption = 0
'   frm_factura.labtimemi.Caption = 0

End If

Unload Me

End Sub

Private Sub Command2_Click()

Dim Xselecciona As String
Dim NroFactura As Double
Dim NroLineas As Integer
Dim Xind As Integer
labhayerror.Caption = ""

If Verificar_seleccion() = 1 Then
   If Xfaccancerecep = 1 Then
      labmotivoref.Caption = InputBox("INGRESE MOTIVO DE LA DEVOLUCION:", "Facturación")
   End If
   For Xind = 1 To ListView1.ListItems.count
       ListView1.ListItems(Xind).Selected = True
       If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
          NroFactura = Val(ListView1.SelectedItem.ListSubItems(4).Text)
          NroLineas = Val(ListView1.SelectedItem.ListSubItems(5).Text)
          Retorna_cedula
          data_lin.RecordSource = "select * from linmmdd where factura =" & NroFactura & " and linea =" & NroLineas & " and cod_cli =" & Val(labcli.Caption) & " and fecha =#" & Format(data_cab.Recordset("cl_fnac"), "yyyy/mm/dd") & "#"
          data_lin.Refresh
          If data_lin.Recordset.RecordCount > 0 Then
             data_lin.Recordset.MoveFirst

             If IsNull(data_lin.Recordset("descuento")) = True Then
                If Xfaccancerecep = 1 Then
                   data_lintemp.RecordSource = "Select * from lineas where nrofactref =" & data_lin.Recordset("factura") & " and serieref ='" & data_lin.Recordset("moneda") & "' and linearef =" & data_lin.Recordset("linea")
                   data_lintemp.Refresh
                Else
                   data_lintemp.RecordSource = "Select * from lineas2 where nrofactref =" & data_lin.Recordset("factura") & " and serieref ='" & data_lin.Recordset("moneda") & "' and linearef =" & data_lin.Recordset("linea")
                   data_lintemp.Refresh
                End If
                If data_lintemp.Recordset.RecordCount > 0 Then
                   MsgBox "La línea seleccionada ya se encuentra facturada", vbCritical
                Else
''                   Xselecciona = MsgBox("FACTURA:" & data_lin.Recordset("factura") & " IMPORTE:" & data_lin.Recordset("tot_lin"), vbInformation + vbYesNo)
                   Xselecciona = vbYes
                   If Xselecciona = vbYes Then
                      If data_lin.Recordset("cod_prod") = 992 Then
                         frm_factura.dbcboprom.Visible = True
                         data_cance992.RecordSource = "select * from deudas where documento =" & data_lin.Recordset("factura") & " and mes =" & 0 & " and fecha_pago is null and cliente =" & data_lin.Recordset("cod_cli")
                         data_cance992.Refresh
                         If data_cance992.Recordset.RecordCount > 0 Then
                            If IsNull(data_cance992.Recordset("descimp")) = False Then
                               If data_cance992.Recordset("descimp") >= 0 Then
                                  If data_lin.Recordset("cod_prod") = 992 Then
                                     Wxeljefeid = data_cance992.Recordset("total")
                                  End If
                               Else
                                  Wxeljefeid = data_cance992.Recordset("total")
                               End If
                            End If
                         Else
                             Wxeljefeid = data_lin.Recordset("tot_lin")
                         End If
                      End If
                      If Xfaccancerecep = 1 Then
                         If data_lin.Recordset("cod_prod") = 992 Then
                            frm_factura.txt_precio.Text = Format(Wxeljefeid, "Standard")
                         Else
                            frm_factura.txt_precio.Text = Format(data_lin.Recordset("tot_lin"), "Standard")
                         End If
                         If IsNull(data_lin.Recordset("mes_paga")) = False Then
                            frm_factura.txt_mes.Text = data_lin.Recordset("mes_paga")
                            frm_factura.txt_ano.Text = data_lin.Recordset("ano_paga")
                         End If
                         If IsNull(data_lin.Recordset("nom_prod")) = False Then
                            frm_factura.DBCombo1.Text = data_lin.Recordset("nom_prod")
                         End If
                         frm_factura.Label5.Caption = data_lin.Recordset("cod_prod")
                         If IsNull(data_lin.Recordset("cod_prod")) = False Then
                            If data_lin.Recordset("cod_prod") = 997 Then
                               If IsNull(data_lin.Recordset("porce_est")) = False Then
                                  Xelnrodeuda = data_lin.Recordset("porce_est")
                               Else
                                  Xelnrodeuda = 0
                               End If
                            End If
                            frm_factura.Label5.Caption = data_lin.Recordset("cod_prod")
                         End If
                         If IsNull(data_lin.Recordset("ruc")) = False Then
                            frm_factura.Check1.Value = 1
                            frm_factura.txt_rut.Text = data_lin.Recordset("ruc")
                         End If
                         frm_factura.lablinea.Caption = data_lin.Recordset("linea")
                         frm_factura.labfaccance.Caption = data_lin.Recordset("factura")
                         frm_factura.labfeccance.Caption = data_lin.Recordset("fecha")
                         If IsNull(data_lin.Recordset("moneda")) = False Then
                            frm_factura.labseriecance.Caption = data_lin.Recordset("moneda")
                         Else
                            frm_factura.labseriecance.Caption = "A"
                         End If
                         If frm_factura.Label7.Caption = "NC E-TICKET" Or frm_factura.Label7.Caption = "NC E-FACTURA" Then
                            If data_lin.Recordset("cod_prod") = 995 Or data_lin.Recordset("cod_prod") = 990 Then
                               frm_factura.txt_precio.Text = data_lin.Recordset("tot_lin")
                               If Trim(frm_factura.labtim.Caption) = "" Then
                                  frm_factura.labtim.Caption = 0
                               End If
                               frm_factura.labtim.Caption = Val(frm_factura.labtim.Caption) + Val(data_lin.Recordset("tot_lin"))
                               frm_factura.cbotim.ListIndex = 0
                            End If
'                            If data_lin.Recordset("tot_lin") <= 0 Then
'                               MsgBox "ATENCION! Importe de la línea sin valor para realizar NC", vbCritical
'                               End
'                            End If
                            frm_factura.txt_precio.Enabled = False
                            frm_factura.DBCombo1.Enabled = False
                         End If
                         frm_factura.labidemi.Caption = 0
                         frm_factura.labdeudaemi.Caption = ""
                         frm_factura.labtimemi.Caption = ""
                         If data_lin.Recordset("cod_prod") = 992 Then
                            frm_factura.dbcboprom.Visible = True
                         End If
                         If data_lin.Recordset("cod_prod") = 60103 Or _
                            data_lin.Recordset("cod_prod") = 60106 Or _
                            data_lin.Recordset("cod_prod") = 60107 Or _
                            data_lin.Recordset("cod_prod") = 60108 Then
                            If IsNull(data_lin.Recordset("nom_medic")) = False Then
                               frm_factura.labmedicacion.Caption = data_lin.Recordset("nom_medic")
                            Else
                                frm_factura.labmedicacion.Caption = ""
                             End If
                             If IsNull(data_lin.Recordset("cod_medic")) = False Then
                                XcodelMedicamento = data_lin.Recordset("cod_medic")
                             Else
                                XcodelMedicamento = 0
                             End If
                             IdTablaPres = 0
                             Verificar_Pedidos
                          Else
                             frm_factura.labmedicacion.Caption = ""
                          End If
                          If Trim(labhayerror.Caption) = "1" Then
                          Else
                             btn_graba_Click
                          End If
                          '''Unload Me
                      Else
                         frm_factconve22.t_imp.Text = Format(data_lin.Recordset("tot_lin"), "Standard")
                         If IsNull(data_lin.Recordset("mes_paga")) = False Then
                            frm_factconve22.t_mes.Text = data_lin.Recordset("mes_paga")
                            frm_factconve22.t_ano.Text = data_lin.Recordset("ano_paga")
                         End If
                         If IsNull(data_lin.Recordset("cantidad")) = False Then
                            frm_factconve22.t_cant.Text = data_lin.Recordset("cantidad")
                         Else
                            frm_factconve22.t_cant.Text = 1
                         End If
                         If IsNull(data_lin.Recordset("tipo_mov")) = False Then
                            If data_lin.Recordset("tipo_mov") = "2" Then
                               frm_factconve22.cboiva.ListIndex = 0
                            Else
                               If data_lin.Recordset("tipo_mov") = "3" Then
                                  frm_factconve22.cboiva.ListIndex = 1
                               Else
                                  If data_lin.Recordset("tipo_mov") = "1" Then
                                     frm_factconve22.cboiva.ListIndex = 2
                                  Else
                                     frm_factconve22.cboiva.ListIndex = 0
                                  End If
                               End If
                            End If
                         Else
                            frm_factconve22.cboiva.ListIndex = 0
                         End If
                         frm_factconve22.t_desc.Text = data_lin.Recordset("nom_prod")
                         frm_factconve22.labnrolin.Caption = data_lin.Recordset("linea")
                         frm_factconve22.labnrofactnc.Caption = data_lin.Recordset("factura")
                         If IsNull(data_lin.Recordset("moneda")) = False Then
                            frm_factconve22.labseriefactnc.Caption = data_lin.Recordset("moneda")
                         Else
                            frm_factconve22.labseriefactnc.Caption = "A"
                         End If
                         If IsNull(data_lin.Recordset("solicitant")) = False Then
                            If data_lin.Recordset("solicitant") = "TRASLADO" Then
                               frm_factconve22.Combo1.ListIndex = 1
                            Else
                               frm_factconve22.Combo1.ListIndex = 0
                            End If
                         Else
                            frm_factconve22.Combo1.ListIndex = 0
                         End If
                         If IsNull(data_cab.Recordset("usu_baja")) = False Then
                            If data_cab.Recordset("usu_baja") = "USD" Then
                               frm_factconve22.cbomon.ListIndex = 1
                            Else
                               frm_factconve22.cbomon.ListIndex = 0
                            End If
                         End If
'''                         Unload Me
                      End If
                   End If
                End If
             Else
                MsgBox "La línea seleccionada ya tiene realizado NC", vbCritical
             End If
          Else
             MsgBox "No se encuentran datos para cancelar.", vbInformation
          End If
       End If
   Next Xind
   If Xfaccancerecep = 1 Then
      cancela_Seleccionado
      If frm_factura.data_lineas.Recordset.RecordCount > 0 Then
         frm_factura.data_lineas.Refresh
         frm_factura.data_lineas.Recordset.MoveFirst
         frm_factura.labtot.Caption = 0
         frm_factura.Label8.Caption = 0
         Do While Not frm_factura.data_lineas.Recordset.EOF
            frm_factura.labtot.Caption = Val(frm_factura.labtot.Caption) + frm_factura.data_lineas.Recordset("tot_lin")
            frm_factura.Label8.Caption = Val(frm_factura.Label8.Caption) + frm_factura.data_lineas.Recordset("imp_iva")
            frm_factura.data_lineas.Recordset.MoveNext
         Loop
         frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard")
         frm_factura.Label8.Caption = Format(frm_factura.Label8.Caption, "Standard")
         frm_factura.Label5.Caption = ""
         frm_factura.DBCombo1.Text = ""
         frm_factura.txt_precio.Text = ""
         frm_factura.data_lindbgri.RecordSource = "select * from lineas"
         frm_factura.data_lindbgri.Refresh
         frm_factura.data_lindbgri.RecordSource = ""
    
      End If
      If Trim(labhayerror.Caption) = "1" Then
         MsgBox "No se puede facturar! VERIFIQUE DEVOLUCION EN MAGIK!", vbCritical
         frm_factura.labtot.Caption = 0
         frm_factura.Label8.Caption = 0
      Else
         Unload Me
      End If
   Else
      Unload Me
   End If
Else
    MsgBox "No hay datos seleccionados.", vbCritical
End If


End Sub

Private Sub Command3_Click()
Dim Xind As Integer

If Trim(esmedica.Caption) = "0" Then
   For Xind = 1 To ListView1.ListItems.count
       ListView1.ListItems(Xind).Selected = True
       If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
          ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = False
       Else
          ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True
       End If
    Next Xind
End If


End Sub

Private Sub Command4_Click()
'''On Error GoTo Alcomman4

frm_factura.data_lineas.Recordset.AddNew
If frm_factura.Label7.Caption = "NC E-FACTURA" Then
   frm_factura.data_lineas.Recordset("tipodocref") = 111
Else
   frm_factura.data_lineas.Recordset("tipodocref") = 101
End If
frm_factura.data_lineas.Recordset("serieref") = frm_factura.labseriecance.Caption
If Len(frm_factura.labfaccance.Caption) > 7 Then
   frm_factura.data_lineas.Recordset("nrofactref") = Val(Mid(frm_factura.labfaccance.Caption, 1, 7))
Else
   frm_factura.data_lineas.Recordset("nrofactref") = Val(frm_factura.labfaccance.Caption)
End If
frm_factura.data_lineas.Recordset("fechafact") = CDate(frm_factura.labfeccance.Caption)
frm_factura.data_lineas.Recordset("motivoref") = Mid(frm_factura.labmotivo.Caption, 1, 90)
frm_factura.data_lineas.Recordset("linearef") = 2

frm_factura.data_lineas.Recordset("reg_cab") = 0
frm_factura.data_lineas.Recordset("factura") = 0
frm_factura.data_lineas.Recordset("tipo_mov") = 1
frm_factura.data_lineas.Recordset("realizada") = Format(frm_factura.mf.Text, "dd/mm/yyyy")
frm_factura.data_lineas.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
frm_factura.data_lineas.Recordset("cod_cli") = frm_factura.labmatri.Caption
frm_factura.data_lineas.Recordset("nom_cli") = frm_factura.labnomb.Caption
frm_factura.data_lineas.Recordset("cod_prod") = 995
frm_factura.data_lineas.Recordset("nom_prod") = "TIMBRE PROFESIONAL"
If frm_factura.txt_rut.Visible = True Then
   If Trim(frm_factura.txt_rut.Text) <> "" Then
      frm_factura.data_lineas.Recordset("ruc") = frm_factura.txt_rut.Text
   End If
End If
frm_factura.data_lineas.Recordset("cantidad") = 1
frm_factura.data_lineas.Recordset("moneda") = "SR"
frm_factura.data_lineas.Recordset("operador") = WElusuario
frm_factura.data_lineas.Recordset("hora") = Format(Time, "HH:mm")
frm_factura.data_lineas.Recordset("nro_flia") = 8
frm_factura.data_lineas.Recordset("nom_flia") = "OTROS SERVICIOS"
frm_factura.data_lineas.Recordset("nro_superv") = frmabm.data_clientes.Recordset("cl_nro_sup")
frm_factura.data_lineas.Recordset("nom_superv") = frmabm.data_clientes.Recordset("cl_nom_sup")
frm_factura.data_lineas.Recordset("convenio") = frmabm.data_clientes.Recordset("cl_codconv")
frm_factura.data_lineas.Recordset("rub_cont") = 213076
frm_factura.data_lineas.Recordset("rub_nomb") = "OBL.TIMBRES PROF."
frm_factura.data_lineas.Recordset("arancel") = Val(frm_factura.labtimemi.Caption)
frm_factura.data_lineas.Recordset("tot_lin") = Val(frm_factura.labtimemi.Caption)
frm_factura.data_lineas.Recordset("precio_est") = Val(frm_factura.labtimemi.Caption)
frm_factura.data_lineas.Recordset("porce_est") = 0
frm_factura.data_lineas.Recordset("base") = frmabm.data_parsec.Recordset("base")
frm_factura.data_lineas.Recordset("tipo") = frm_factura.labfpago.Caption
frm_factura.data_lineas.Recordset("linea") = 2
frm_factura.data_lineas.Recordset("libro_rub") = frm_factura.Label7.Caption ' tipo de documento (Ej.e-ticket)
frm_factura.data_lineas.Recordset("in_unid") = "INT1"
frm_factura.data_lineas.Recordset.Update
frm_factura.data_lineas.Refresh
If frm_factura.labtot.Caption <> "" Then
   frm_factura.labtot.Caption = Val(frm_factura.labtot.Caption) + Val(frm_factura.labtimemi.Caption)
Else
   frm_factura.labtot.Caption = Val(frm_factura.labtimemi.Caption)
End If

'Exit Sub

'Alcomman4:
'         If Err.Number = 5 Then
'            data_errfact.Recordset.AddNew
'            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
''            data_errfact.Recordset("fecha") = Date
'            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
'            data_errfact.Recordset("nroerr") = Err.Number
'            data_errfact.Recordset("desc") = "Al Comman4"
'            data_errfact.Recordset.Update
''            Unload Me
'         Else
'            data_errfact.Recordset.AddNew
'            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
'            data_errfact.Recordset("fecha") = Date
'            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
''            data_errfact.Recordset("nroerr") = Err.Number
'            data_errfact.Recordset("desc") = "Al Comman4"
'            data_errfact.Recordset.Update
'            Unload Me
'         End If

End Sub

Private Sub Command5_Click()
''On Error GoTo Alcomman5

frm_factura.data_lineas.Recordset.AddNew
If frm_factura.Label7.Caption = "NC E-FACTURA" Then
   frm_factura.data_lineas.Recordset("tipodocref") = 111
Else
   frm_factura.data_lineas.Recordset("tipodocref") = 101
End If
frm_factura.data_lineas.Recordset("serieref") = frm_factura.labseriecance.Caption
If Len(frm_factura.labfaccance.Caption) > 7 Then
   frm_factura.data_lineas.Recordset("nrofactref") = Val(Mid(frm_factura.labfaccance.Caption, 1, 7))
Else
   frm_factura.data_lineas.Recordset("nrofactref") = Val(frm_factura.labfaccance.Caption)
End If
frm_factura.data_lineas.Recordset("fechafact") = CDate(frm_factura.labfeccance.Caption)
frm_factura.data_lineas.Recordset("motivoref") = Mid(frm_factura.labmotivo.Caption, 1, 90)
If frm_factura.labtimemi.Caption <> "" Then
   If Format(frm_factura.labtimemi.Caption, "Standard") > 0 Then
      frm_factura.data_lineas.Recordset("linearef") = 3
   Else
      frm_factura.data_lineas.Recordset("linearef") = 2
   End If
Else
   frm_factura.data_lineas.Recordset("linearef") = 2
End If
frm_factura.data_lineas.Recordset("reg_cab") = 0
frm_factura.data_lineas.Recordset("factura") = 0
frm_factura.data_lineas.Recordset("tipo_mov") = 1
frm_factura.data_lineas.Recordset("realizada") = Format(frm_factura.mf.Text, "dd/mm/yyyy")
frm_factura.data_lineas.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
frm_factura.data_lineas.Recordset("cod_cli") = frm_factura.labmatri.Caption
frm_factura.data_lineas.Recordset("nom_cli") = frm_factura.labnomb.Caption
frm_factura.data_lineas.Recordset("cod_prod") = 882
frm_factura.data_lineas.Recordset("nom_prod") = "DEUDAS POR SERVICIOS"
If frm_factura.txt_rut.Visible = True Then
   If Trim(frm_factura.txt_rut.Text) <> "" Then
      frm_factura.data_lineas.Recordset("ruc") = frm_factura.txt_rut.Text
   End If
End If
frm_factura.data_lineas.Recordset("cantidad") = 1
frm_factura.data_lineas.Recordset("moneda") = "SR"
frm_factura.data_lineas.Recordset("operador") = WElusuario
frm_factura.data_lineas.Recordset("hora") = Format(Time, "HH:mm")
frm_factura.data_lineas.Recordset("nro_flia") = 8
frm_factura.data_lineas.Recordset("nom_flia") = "OTROS SERVICIOS"
frm_factura.data_lineas.Recordset("nro_superv") = frmabm.data_clientes.Recordset("cl_nro_sup")
frm_factura.data_lineas.Recordset("nom_superv") = frmabm.data_clientes.Recordset("cl_nom_sup")
frm_factura.data_lineas.Recordset("convenio") = frmabm.data_clientes.Recordset("cl_codconv")
frm_factura.data_lineas.Recordset("rub_cont") = 213041
frm_factura.data_lineas.Recordset("rub_nomb") = "PROVISORIOS"
frm_factura.data_lineas.Recordset("arancel") = Format(frm_factura.labdeudaemi.Caption, "Standard")
frm_factura.data_lineas.Recordset("tot_lin") = Format(frm_factura.labdeudaemi.Caption, "Standard")
frm_factura.data_lineas.Recordset("precio_est") = Format(frm_factura.labdeudaemi.Caption, "Standard")
frm_factura.data_lineas.Recordset("porce_est") = 0
frm_factura.data_lineas.Recordset("base") = frmabm.data_parsec.Recordset("base")
frm_factura.data_lineas.Recordset("tipo") = frm_factura.labfpago.Caption
frm_factura.data_lineas.Recordset("linea") = 2
frm_factura.data_lineas.Recordset("libro_rub") = frm_factura.Label7.Caption ' tipo de documento (Ej.e-ticket)
frm_factura.data_lineas.Recordset("in_unid") = "INT1"
frm_factura.data_lineas.Recordset.Update
frm_factura.data_lineas.Refresh
Dim Xelivaemi As Double
Xelivaemi = Format(frm_factura.labdeudaemi.Caption, "Standard") / 1.1 * 0.1

If frm_factura.labtot.Caption <> "" Then
   frm_factura.labtot.Caption = Format(frm_factura.labtot.Caption, "Standard") + Format(frm_factura.labdeudaemi.Caption, "Standard")
   frm_factura.Label8.Caption = Format(frm_factura.Label8.Caption, "Standard") + Xelivaemi
   frm_factura.Label8.Caption = Format(frm_factura.Label8.Caption, "Standard")
Else
   frm_factura.labtot.Caption = Format(frm_factura.labdeudaemi.Caption, "Standard")
   frm_factura.Label8.Caption = Format(frm_factura.Label8.Caption, "Standard") + Xelivaemi
   frm_factura.Label8.Caption = Format(frm_factura.Label8.Caption, "Standard")
End If

'Exit Sub

'Alcomman5:
'         If Err.Number = 5 Then
'            data_errfact.Recordset.AddNew
'            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
'            data_errfact.Recordset("fecha") = Date
'            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
'            data_errfact.Recordset("nroerr") = Err.Number
'            data_errfact.Recordset("desc") = "Al comman5"
''            data_errfact.Recordset.Update
'            Unload Me
'         Else
'            data_errfact.Recordset.AddNew
'            data_errfact.Recordset("id") = frm_menu.data_parse.Recordset("base")
'            data_errfact.Recordset("fecha") = Date
'            data_errfact.Recordset("hora") = Format(Time, "HH:mm")
'            data_errfact.Recordset("nroerr") = Err.Number
'            data_errfact.Recordset("desc") = "Al comman5"
'            data_errfact.Recordset.Update
'            Unload Me
'         End If

End Sub

Private Sub DBGrid1_DblClick()
Carga_grid

End Sub


Private Sub DBGrid3_DblClick()
Dim Xseleccionadeu As String
Dim Nomlaemi As String
Nomlaemi = "emi"
If frm_menu.data_parse.Recordset("base") = 38 Then
Else
    If IsNull(data_deudas.Recordset("mes")) = False Then
       If data_deudas.Recordset("mes") > 0 Then
          If Month(data_deudas.Recordset("fecha")) < 10 Then
             'seleccionar desde la fecha de creado
             Nomlaemi = Nomlaemi & "0" & Trim(str(Month(data_deudas.Recordset("fecha")))) & Mid(Trim(str(Year(data_deudas.Recordset("fecha")))), 3, 2)
             'Nomlaemi = Nomlaemi & "0" & Trim(str(data_deudas.Recordset("mes"))) & Mid(Trim(str(data_deudas.Recordset("ano"))), 3, 2)
          Else
'             Nomlaemi = Nomlaemi & Trim(str(data_deudas.Recordset("mes"))) & Mid(Trim(str(data_deudas.Recordset("ano"))), 3, 2)
             Nomlaemi = Nomlaemi & Trim(str(Month(data_deudas.Recordset("fecha")))) & Mid(Trim(str(Month(data_deudas.Recordset("fecha")))), 3, 2) & Mid(Trim(str(Year(data_deudas.Recordset("fecha")))), 3, 2)
          End If
          If data_deudas.Recordset("ano") > 2015 Then
             
            data_emi.RecordSource = "Select * from " & Nomlaemi & " where cliente =" & data_deudas.Recordset("cliente") & " and documento =" & data_deudas.Recordset("documento")
            data_emi.Refresh
            If data_emi.Recordset.RecordCount > 0 Then
               If IsNull(data_emi.Recordset("ruc")) = False Then
                  If frm_factura.Label7.Caption = "NC E-TICKET" Then
                     MsgBox "Debe facturar NC de E-FACTURA", vbInformation
                     End
                  Else
                     frm_factura.txt_rut.Text = data_emi.Recordset("ruc")
                     frm_factura.Check1.Value = 1
                  End If
               Else
                  If frm_factura.Label7.Caption = "NC E-FACTURA" Then
                     MsgBox "Debe facturar NC de E-TICKET", vbInformation
                     End
                  End If
               End If
            End If
          End If
       End If
    End If
End If
Xseleccionadeu = MsgBox("FACTURA:" & data_deudas.Recordset("documento") & " IMPORTE:" & data_deudas.Recordset("total"), vbInformation + vbYesNo)
If Xseleccionadeu = vbYes Then
   frm_factura.txt_precio.Text = Format(data_deudas.Recordset("total"), "Standard")
   If IsNull(data_deudas.Recordset("mes")) = False Then
      frm_factura.txt_mes.Text = data_deudas.Recordset("mes")
      frm_factura.txt_ano.Text = data_deudas.Recordset("ano")
   End If
   If IsNull(data_deudas.Recordset("tiquet")) = False Then
      frm_factura.labtimemi.Caption = Format(data_deudas.Recordset("tiquet"), "Standard")
      Xeltotalemi = Xeltotalemi - data_deudas.Recordset("tiquet")
   Else
      frm_factura.labtimemi.Caption = 0
   End If
   If IsNull(data_deudas.Recordset("deudas")) = False Then
      frm_factura.labdeudaemi.Caption = Format(data_deudas.Recordset("deudas"), "Standard")
   Else
      frm_factura.labdeudaemi.Caption = 0
   End If
   frm_factura.DBCombo1.Text = "CUOTA MENSUAL"
   frm_factura.Label5.Caption = 881
   frm_factura.lablinea.Caption = 1
   frm_factura.labfaccance.Caption = data_deudas.Recordset("documento")
   frm_factura.labfeccance.Caption = data_deudas.Recordset("fecha")
   If IsNull(data_deudas.Recordset("tipocta")) = False Then
      frm_factura.labseriecance.Caption = data_deudas.Recordset("tipocta")
   Else
      frm_factura.labseriecance.Caption = "A"
   End If
   frm_factura.labidemi.Caption = 5
   Unload Me
End If


End Sub

Private Sub DBGrid4_DblClick()
Dim Xseleccionarq As String
Dim Xeltotalemi As Double

frm_factura.labidemi.Caption = 0
frm_factura.labdeudaemi.Caption = 0
frm_factura.labtimemi.Caption = 0

Xeltotalemi = 0
Xseleccionarq = MsgBox("FACTURA:" & data_arq.Recordset("nrorec") & " IMPORTE:" & data_arq.Recordset("total"), vbInformation + vbYesNo)
If Xseleccionarq = vbYes Then
   frm_factura.txt_precio.Text = Format(data_arq.Recordset("total"), "Standard")
   If IsNull(data_arq.Recordset("mes")) = False Then
      frm_factura.txt_mes.Text = data_arq.Recordset("mes")
      frm_factura.txt_ano.Text = data_arq.Recordset("ano")
   End If
   If IsNull(data_arq.Recordset("tiquet")) = False Then
      frm_factura.labtimemi.Caption = Format(data_arq.Recordset("tiquet"), "Standard")
      Xeltotalemi = Xeltotalemi - data_arq.Recordset("tiquet")
   Else
      frm_factura.labtimemi.Caption = 0
   End If
   If IsNull(data_arq.Recordset("deudas")) = False Then
      frm_factura.labdeudaemi.Caption = Format(data_arq.Recordset("deudas"), "Standard")
   Else
      frm_factura.labdeudaemi.Caption = 0
   End If
   frm_factura.DBCombo1.Text = "CUOTA MENSUAL"
   frm_factura.Label5.Caption = 881
   frm_factura.lablinea.Caption = 1
   frm_factura.labfaccance.Caption = data_arq.Recordset("nrorec")
   frm_factura.labfeccance.Caption = data_arq.Recordset("fecha")
   frm_factura.labseriecance.Caption = "A"
   frm_factura.labidemi.Caption = 5
   Unload Me
End If

End Sub

Private Sub Form_Load()

data_estudiobus.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_codcaja.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_codcaja.RecordSource = "cod_caja"
data_codcaja.Refresh

data_lineas.DatabaseName = App.path & "\factura.mdb"
data_lineas.RecordSource = "lineas"
data_lineas.Refresh

Wxeljefeid = 0
data_arq.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cance992.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_lintemp.DatabaseName = App.path & "\factura.mdb"
If frm_menu.data_parse.Recordset("base") = 38 Then
   data_deudas.Connect = "odbc;dsn=sappfact;"
   data_cab.Connect = "odbc;dsn=sappfact;"
   data_lin.Connect = "odbc;dsn=sappfact;"
Else
   data_deudas.Connect = "odbc;dsn=" & Xconexrmt & ";"
   data_cab.Connect = "odbc;dsn=" & Xconexrmt & ";"
   data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If
If Xfaccancerecep = 1 Then
   labcli.Caption = frm_factura.labmatri.Caption
   If XQuefac = 102 Or XQuefac = 103 Then
      data_cab.RecordSource = "Select * from clirespl where cl_codigo =" & labcli.Caption & " and cl_tipocli =" & 101 & " order by cl_fnac DESC"
      data_cab.Refresh
   Else
      If XQuefac = 112 Or XQuefac = 113 Then
         data_cab.RecordSource = "Select * from clirespl where cl_codigo =" & labcli.Caption & " and cl_tipocli =" & 111 & " order by cl_fnac DESC"
         data_cab.Refresh
      Else
         If XQuefac = 21 Then
            data_cab.RecordSource = "Select * from clirespl where cl_codigo =" & labcli.Caption & " and cl_telefon ='" & "RECIBO" & "' order by cl_fnac DESC"
            data_cab.Refresh
         End If
      End If
   End If
Else
   labcli.Caption = frm_factconve22.labnrocli.Caption
   If XAlta = 12 Or XAlta = 16 Then
      data_cab.RecordSource = "Select * from clirespl where cl_codigo =" & labcli.Caption & " and cl_tipocli =" & 101 & " order by cl_fnac DESC"
      data_cab.Refresh
   Else
      data_cab.RecordSource = "Select * from clirespl where cl_codigo =" & labcli.Caption & " and cl_tipocli =" & 111 & " order by cl_fnac DESC"
      data_cab.Refresh
   End If
End If
data_deudas.RecordSource = "Select * from deudas where cliente =" & labcli.Caption & " and fecha_pago is null and mes >" & 0 & " order by ano,mes"
data_deudas.Refresh

data_arq.RecordSource = "Select * from arqueo where matricula =" & labcli.Caption & " order by ano,mes"
data_arq.Refresh

data_emi.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub Form_Resize()
With Image1
     Image1.Left = 0
     Image1.Top = 0
     Image1.Height = Me.Height
     Image1.Width = Me.Width
End With

End Sub

Private Sub t_nrofact_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nrolin.SetFocus
End If

End Sub

Private Sub t_nrolin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mf.SetFocus
End If

End Sub

Public Sub Verificar_Pedidos()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xcanti As Double

Xcanti = 0

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD

ConbdSapp.Open

Xsqlpromo = "Select * from pedidos_facturar where matricula =" & Val(labcedula.Caption) & " and fecha_fact is null and cantidad <" & 0 & " and cod_medicacion =" & XcodelMedicamento & " and seleccionado is null"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
Xcount = 1

If Xrecclii.RecordCount > 0 Then
   Xcanti = Xrecclii("cantidad")
   If Xcanti = -1 Then
      Xrecclii("seleccionado") = 1
      Xrecclii.Update
   End If
   labidpedido.Caption = Xrecclii("id")
Else
   labidpedido.Caption = ""
   MsgBox "ATENCION! No existe registro de devolución en Apraful de:" & data_lin.Recordset("nom_medic") & "!!!! VERIFIQUE!", vbCritical
   labhayerror.Caption = "1"
End If

Xrecclii.Close
ConbdSapp.Close

End Sub
Public Sub Retorna_cedula()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD
ConbdSapp.Open
                          
Xsqlpromo = "Select * from clientes where cl_codigo =" & Val(labcli.Caption)

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   If IsNull(Xrecclii("cl_cedula")) = False Then
      labcedula.Caption = Xrecclii("cl_cedula")
   Else
      labcedula.Caption = "0"
   End If
Else
   labcedula.Caption = "0"
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Carga_grid()
Dim Xcount As Long
Dim a, b, c, d, e, f, g, h, i, j, k As String
a = "a"
b = "b"
c = "c"
d = "d"
e = "e"
f = "f"
g = "g"
h = "h"
i = "i"
j = "j"
k = "k"
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD

ConbdSapp.Open
If XQuefac = 21 Then
   Xsqlpromo = "Select * from linmmdd where factura =" & data_cab.Recordset("cl_numero") & " and cod_cli =" & data_cab.Recordset("cl_codigo")
Else
   Xsqlpromo = "Select * from linmmdd where factura =" & data_cab.Recordset("cl_numero") & " and moneda ='" & data_cab.Recordset("cl_socmnro") & "' and fecha ='" & Format(data_cab.Recordset("cl_fnac"), "yyyy-mm-dd") & "'"
End If
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
Xcount = 1
ListView1.ListItems.Clear
esmedica.Caption = "0"
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      If IsNull(Xrecclii("factura")) = False Then
         ListView1.ListItems.Add Xcount, , Xrecclii("cod_prod")
         If Xrecclii("cod_prod") = 60103 Or Xrecclii("cod_prod") = 60106 Or _
            Xrecclii("cod_prod") = 60107 Or Xrecclii("cod_prod") = 60108 Or _
            Xrecclii("cod_prod") = 990 Then
            ListView1.ListItems(Xcount).Selected = True
            ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True
            esmedica.Caption = "1"
         End If
      Else
         ListView1.ListItems.Add Xcount, , "0"
      End If
      If IsNull(Xrecclii("nom_prod")) = False Then
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xrecclii("nom_prod")
      Else
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/D"
      End If
      If IsNull(Xrecclii("tot_lin")) = False Then
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xrecclii("tot_lin")
      Else
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
      End If
      If IsNull(Xrecclii("nom_medic")) = False Then
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xrecclii("nom_medic")
      Else
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "S/D"
      End If
      If IsNull(Xrecclii("factura")) = True Then
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
      Else
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xrecclii("factura")
      End If
      If IsNull(Xrecclii("linea")) = False Then
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Xrecclii("linea")
      Else
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
      End If
      Xrecclii.MoveNext
      Xcount = Xcount + 1
   Loop
Else
    MsgBox "No existe pedido de medicación.", vbInformation, "Pedidos"
End If
If esmedica.Caption = "1" Then
   ListView1.Enabled = False
Else
   ListView1.Enabled = True
End If
Xrecclii.Close
ConbdSapp.Close


End Sub
Public Function Verificar_seleccion() As Integer
Dim Xind, Xsihay As Integer
Xsihay = 0
For Xind = 1 To ListView1.ListItems.count
    ListView1.ListItems(Xind).Selected = True
    If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
       Xsihay = 1
    End If
Next Xind
Verificar_seleccion = Xsihay


End Function


Public Sub Verificar_codigo()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xcodigo As String

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If
Xcodigo = ""

ConectarBD
ConbdSapp.Open

Xcodigo = InputBox("Ingrese código de autorización:", "Autorización para anular registro.")
If Trim(Xcodigo) <> "" Then
   Xsqlpromo = "Select * from codaut_devol where codigo =" & Val(Xcodigo) & " and usado in (0)"
   With Xrecclii
      .CursorLocation = adUseClient
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .Open Xsqlpromo, ConbdSapp, , , adCmdText
   End With
   If Xrecclii.RecordCount > 0 Then
      DBGrid1.Enabled = False
      t_nrofact.Visible = True
      t_nrolin.Visible = True
      mf.Visible = True
      Label3.Visible = True
      Label4.Visible = True
      Label5.Visible = True
      Command1.Visible = True
      Xrecclii("usado") = 1
      Xrecclii.Update
   Else
      MsgBox "Código incorrecto", vbCritical
      DBGrid1.Enabled = True
      t_nrofact.Text = ""
      t_nrofact.Visible = False
      t_nrolin.Text = ""
      t_nrolin.Visible = False
      mf.Text = "__/__/____"
      mf.Visible = False
      Label3.Visible = False
      Label4.Visible = False
      Label5.Visible = False
      Command1.Visible = False
   End If
   Xrecclii.Close
Else
   MsgBox "No ingresó código.", vbCritical
   DBGrid1.Enabled = True
   t_nrofact.Text = ""
   t_nrofact.Visible = False
   t_nrolin.Text = ""
   t_nrolin.Visible = False
   mf.Text = "__/__/____"
   mf.Visible = False
   Label3.Visible = False
   Label4.Visible = False
   Label5.Visible = False
   Command1.Visible = False
End If

ConbdSapp.Close

End Sub


Public Sub cancela_Seleccionado()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xcanti As Integer
Xcanti = 0

If ConbdSapp.State = 1 Then
   ConbdSapp.Close
End If

ConectarBD

ConbdSapp.Open

Xsqlpromo = "Select * from pedidos_facturar where matricula =" & Val(labcedula.Caption) & " and fecha_fact is null and cantidad <" & 0 & " and seleccionado in (1)"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
Xcount = 1

If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      Xrecclii("seleccionado") = Null
      Xrecclii.Update
      Xrecclii.MoveNext
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub
