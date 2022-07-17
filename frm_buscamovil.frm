VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_buscamovil 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar datos de móviles"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7590
   Icon            =   "frm_buscamovil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7590
   StartUpPosition =   1  'CenterOwner
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_buscamovil.frx":0442
      Height          =   3735
      Left            =   240
      OleObjectBlob   =   "frm_buscamovil.frx":0456
      TabIndex        =   1
      Top             =   480
      Width           =   7095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   2895
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
      Left            =   6600
      Picture         =   "frm_buscamovil.frx":0E31
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MOVILES:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   3480
      Picture         =   "frm_buscamovil.frx":13BB
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1455
   End
End
Attribute VB_Name = "frm_buscamovil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
frm_opsdesp.data_mov.Recordset.FindFirst "movil =" & Data1.Recordset("movil")
If Not frm_opsdesp.data_mov.Recordset.NoMatch Then
   frm_opsdesp.txt_nro.Text = Data1.Recordset("movil")
   If IsNull(Data1.Recordset("codmed")) = False Then
      frm_opsdesp.txt_codmed.Text = Data1.Recordset("codmed")
   Else
      frm_opsdesp.txt_codmed.Text = 0
   End If
   If IsNull(Data1.Recordset("nommed")) = False Then
      frm_opsdesp.dbmedic.Text = Data1.Recordset("nommed")
   Else
      frm_opsdesp.dbmedic.Text = "Sin Datos"
   End If
   If IsNull(Data1.Recordset("fecha_act")) = False Then
      frm_opsdesp.mfec.Text = Format(Data1.Recordset("fecha_act"), "dd/mm/yyyy")
   Else
      frm_opsdesp.mfec.Text = "__/__/____"
   End If
   If IsNull(Data1.Recordset("ano")) = False Then
      frm_opsdesp.t_base.Text = Data1.Recordset("ano")
   Else
      frm_opsdesp.t_base.Text = 0
   End If
   If IsNull(Data1.Recordset("codchof")) = False Then
      frm_opsdesp.labchof.Caption = Data1.Recordset("codchof")
   Else
      frm_opsdesp.labchof.Caption = 0
   End If
   If IsNull(Data1.Recordset("nomchof")) = False Then
      frm_opsdesp.t_chof(0).Text = Data1.Recordset("nomchof")
   Else
      frm_opsdesp.t_chof(0).Text = ""
   End If
   If IsNull(Data1.Recordset("codenf")) = False Then
      frm_opsdesp.labenf.Caption = Data1.Recordset("codenf")
   Else
      frm_opsdesp.labenf.Caption = 0
   End If
   If IsNull(Data1.Recordset("nomenf")) = False Then
      frm_opsdesp.t_enf(0).Text = Data1.Recordset("nomenf")
   Else
      frm_opsdesp.t_enf(0).Text = ""
   End If
   Unload Me
Else
   MsgBox "Móvil no encontrado", vbInformation, "Buscar..."
End If

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "Select * from moviles order by movil"
Data1.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Top = 0
    .Left = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub
