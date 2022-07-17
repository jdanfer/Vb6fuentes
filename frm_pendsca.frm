VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_pendsca 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pendientes para seguimiento de SCA"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11160
   Icon            =   "frm_pendsca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   11160
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Top             =   4320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_verctrol 
      Caption         =   "data_verctrol"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_grabaseg 
      Caption         =   "data_grabaseg"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   7800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_medicos 
      Caption         =   "data_medicos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_param 
      Caption         =   "data_param"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10320
      Picture         =   "frm_pendsca.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Actualizar registros"
      Top             =   3600
      Width           =   735
   End
   Begin VB.Data data_lin 
      Caption         =   "data_lin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   7800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      Picture         =   "frm_pendsca.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cancelar selección"
      Top             =   3600
      Width           =   735
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Control Actual"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   10815
      Begin MSMask.MaskEdBox mdefh 
         Height          =   375
         Left            =   9240
         TabIndex        =   26
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin MSMask.MaskEdBox mdef 
         Height          =   375
         Left            =   7800
         TabIndex        =   25
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.CommandButton b_graba 
         Height          =   375
         Left            =   720
         Picture         =   "frm_pendsca.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Graba el control con los datos registrados"
         Top             =   1080
         Width           =   495
      End
      Begin MSMask.MaskEdBox mfalta 
         Height          =   375
         Left            =   9480
         TabIndex        =   16
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfprox 
         Height          =   375
         Left            =   7800
         TabIndex        =   15
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mhctr 
         Height          =   375
         Left            =   6840
         TabIndex        =   12
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfctr 
         Height          =   375
         Left            =   5640
         TabIndex        =   11
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox t_deta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C00000&
         Caption         =   "ALTA  de SAPP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   5640
         TabIndex        =   24
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alta de Ctrol."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9480
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Próximo Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7800
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha/Hora control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5640
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "En suma:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Historial de Controles"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   6240
      Width           =   10815
      Begin VB.TextBox t_datos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1095
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   8415
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Doble click sobre el número de control para visualizar datos"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Nro.de Control"
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
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   10560
      Top             =   720
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_pendsca.frx":1628
      Height          =   3375
      Left            =   240
      OleObjectBlob   =   "frm_pendsca.frx":163C
      TabIndex        =   1
      Top             =   240
      Width           =   10815
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Top             =   3000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_pendsca.frx":29F3
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Seleccionar el paciente para registrar datos de siguimiento"
      Top             =   3600
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      RecordSource    =   "cabezal_hcdig"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackColor       =   &H00404040&
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
      Height          =   495
      Left            =   4440
      TabIndex        =   23
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00404040&
      Caption         =   "Datos de contacto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   2760
      TabIndex        =   22
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      Top             =   4320
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
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
      Left            =   1920
      TabIndex        =   9
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      Caption         =   "Usuario actual:"
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
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frm_pendsca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NroControl As Integer

Private Sub b_graba_Click()
On Error GoTo Algrabar

If t_deta.Text <> "" Then
   If mfctr.Text <> "__/__/____" And mfprox.Text <> "__/__/____" Then
      If mhctr.Text <> "__:__" Then
         data_grabaseg.Recordset.AddNew
         data_grabaseg.Recordset("fecha") = Date
         data_grabaseg.Recordset("hora") = Format(Time, "HH:mm")
         data_grabaseg.Recordset("mat") = Data4.Recordset("mat")
         data_grabaseg.Recordset("medicocod") = Data4.Recordset("medicocod")
         data_grabaseg.Recordset("obs") = t_deta.Text
         data_grabaseg.Recordset("fecha_prox") = Format(mfprox.Text, "yyyy-mm-dd")
         data_grabaseg.Recordset("fecha_ctrol") = Format(mfctr.Text, "yyyy-mm-dd")
         data_grabaseg.Recordset("hora_ctrol") = Format(Time, "HH:mm")
         If mfalta.Text <> "__/__/____" Then
            data_grabaseg.Recordset("fecha_alta") = Format(mfalta.Text, "yyyy-mm-dd")
         End If
         If mdef.Text <> "__/__/____" Then
            data_grabaseg.Recordset("fecha_altafin") = Format(mdef.Text, "yyyy-mm-dd")
         End If
         If mdefh.Text <> "__:__" Then
            data_grabaseg.Recordset("hora_altafin") = Format(mdefh.Text, "HH:mm")
         End If
         
         data_grabaseg.Recordset("nro_ctrol") = NroControl + 1
         data_grabaseg.Recordset("id_seguimiento") = Data4.Recordset("id")
         data_grabaseg.Recordset.Update
         Data4.Recordset.Edit
         Data4.Recordset("fecha_prox") = Format(mfprox.Text, "yyyy-mm-dd")
         Data4.Recordset("usuario_cierre") = WElusuario
         Data4.Recordset("hora_modif") = Format(Time, "HH:mm:ss")
         
         If mfalta.Text <> "__/__/____" Then
            Data4.Recordset("fecha_cierre") = Format(mfalta.Text, "yyyy-mm-dd")
            Data4.Recordset("hora_cierre") = Format(Time, "HH:mm")
            Data4.Recordset("usuario_cierre") = WElusuario
            If mdef.Text <> "__/__/____" Then
               Data4.Recordset("fecha_altafin") = Format(mdef.Text, "yyyy-mm-dd")
            End If
            If mdefh.Text <> "__:__" Then
               Data4.Recordset("hora_altafin") = Format(mdefh.Text, "HH:mm")
            End If
         End If
         Data4.Recordset.Update
         Label9.Caption = ""
         DBGrid1.Enabled = True
         Command2.Enabled = False
         Command1.Enabled = True
         Command3.Enabled = True
         Frame1.Enabled = False
         Frame2.Enabled = False
        
         List1.Clear
         t_datos.Text = ""
         t_deta.Text = ""
         mfctr.Text = "__/__/____"
         mhctr.Text = "__:__"
         mfprox.Text = "__/__/____"
         mfalta.Text = "__/__/____"
         NroControl = 0
         Command3_Click
         DBGrid1.SetFocus
      
      Else
         MsgBox "No ingresó hora"
      End If
   Else
      MsgBox "No ingresó fecha"
   End If
Else
   MsgBox "No ingresó detalles"
   
End If

Exit Sub

Algrabar:
        If Err.Number = 3155 Then
           MsgBox "Error al grabar, verifique datos " & Err.Description
        Else
           MsgBox "Error al graba, verifique datos " & Err.Description
        End If
        
End Sub

Private Sub Command1_Click()

Label9.Caption = Data4.Recordset("nombre")
DBGrid1.Enabled = False
Command2.Enabled = True
Command1.Enabled = False
Command3.Enabled = False
Frame1.Enabled = True
Frame2.Enabled = True

List1.Clear
t_datos.Text = ""
t_deta.Text = ""
mfctr.Text = "__/__/____"
mhctr.Text = "__:__"
mfprox.Text = "__/__/____"
mfalta.Text = "__/__/____"
NroControl = 0

'VerSegui
data_grabaseg.RecordSource = "select * from seguimiento_sca where mat =" & Data4.Recordset("mat") & " and id_seguimiento =" & Data4.Recordset("id") & " order by nro_ctrol"
data_grabaseg.Refresh
If data_grabaseg.Recordset.RecordCount > 0 Then
   data_grabaseg.Recordset.MoveFirst
   Do While Not data_grabaseg.Recordset.EOF
      NroControl = NroControl + 1
      List1.AddItem data_grabaseg.Recordset("nro_ctrol")
      data_grabaseg.Recordset.MoveNext
   Loop
   Data5.RecordSource = "select * from clientes where cl_codigo =" & Data4.Recordset("mat")
   Data5.Refresh
   If Data5.Recordset.RecordCount > 0 Then
      If IsNull(Data5.Recordset("cl_dpto")) = False Then
         If Data5.Recordset("cl_dpto") = "NO APLICA" Then
            Label10.Caption = "CEL: Sin Dato "
         Else
            Label10.Caption = "CEL: " & Data5.Recordset("cl_dpto")
         End If
      Else
         Label10.Caption = "CEL: Sin Dato "
      End If
      If IsNull(Data5.Recordset("cl_telefon")) = False Then
         If Data5.Recordset("cl_telefon") = "NO APLICA" Then
            Label10.Caption = Label10.Caption & " TEL: Sin Dato "
         Else
            Label10.Caption = Label10.Caption & " TEL: " & Data5.Recordset("cl_telefon")
         End If
      Else
         Label10.Caption = Label10.Caption & " TEL: Sin Dato "
      End If
   Else
      Label10.Caption = "sin datos"
   End If
End If

t_deta.SetFocus


End Sub

Private Sub Command2_Click()
         Label9.Caption = ""
         DBGrid1.Enabled = True
         Command2.Enabled = False
         Command1.Enabled = True
         Command3.Enabled = True
         Frame1.Enabled = False
         Frame2.Enabled = False
        
         List1.Clear
         t_datos.Text = ""
         t_deta.Text = ""
         mfctr.Text = "__/__/____"
         mhctr.Text = "__:__"
         mfprox.Text = "__/__/____"
         mfalta.Text = "__/__/____"
         NroControl = 0
         DBGrid1.SetFocus
         
End Sub

Private Sub Command3_Click()
frm_pendsca.MousePointer = 11

Data4.RecordSource = "select * from pendiente_sca where fecha_cierre is null and (fecha_prox is null or fecha_prox =#" & Format(Date, "yyyy/mm/dd") & "#)"
Data4.Refresh
frm_pendsca.MousePointer = 0

End Sub

Private Sub Form_Load()

'810 SERVICIO CONTROL AMBULATORIO

Data3.Connect = "odbc;dsn=" & Xconexrmt & ";"
frm_pendsca.MousePointer = 11
Data4.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_grabaseg.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_grabaseg.ConnectionString = "dsn=" & Xconexrmt

data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"

Data5.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_verctrol.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_medicos.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_param.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_param.RecordSource = "param_gral"
data_param.Refresh

Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "select cabezal_hcdig.id,cabezal_hcdig.mat,cabezal_hcdig.fecha,cabezal_hcdig.hora,cabezal_hcdig.hc_sca,cabezal_hcdig.hc_scasi," & _
"cabezal_hcdig.hc_base,cabezal_hcdig.hc_codmed,cabezal_hcdig.hc_nommed,clientes.cl_codigo,clientes.cl_apellid,clientes.cl_codconv,clientes.cl_cedula,clientes.cl_codced from " & _
"cabezal_hcdig inner join clientes on cabezal_hcdig.mat=clientes.cl_codigo where cabezal_hcdig.hc_sca in (1) and cabezal_hcdig.hc_scasi is null"
Data1.Refresh

Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"

If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      Data2.RecordSource = "select * from pendiente_sca where mat =" & Data1.Recordset("mat") & " and fecha_cierre is null"
      Data2.Refresh
      If Data2.Recordset.RecordCount > 0 Then
         frm_pendsca.MousePointer = 0
         MsgBox "El paciente con matrícula " & Data1.Recordset("mat") & " ya existe en seguimiento SCA como pendiente.", vbInformation
         frm_pendsca.MousePointer = 11
      Else
         Data2.Recordset.AddNew
         Data2.Recordset("mat") = Data1.Recordset("mat")
         Data2.Recordset("nombre") = Data1.Recordset("cl_apellid")
         Data2.Recordset("fecha") = Data1.Recordset("fecha")
         Data2.Recordset("hora") = Data1.Recordset("hora")
         Data2.Recordset("hc_nro") = Data1.Recordset("id")
         Data2.Recordset("base") = Data1.Recordset("hc_base")
         Data2.Recordset("medicocod") = Data1.Recordset("hc_codmed")
         Data2.Recordset("mediconom") = Data1.Recordset("hc_nommed")
         Data2.Recordset.Update
         Data3.RecordSource = "select * from cabezal_hcdig where mat=" & Data1.Recordset("mat") & " and id =" & Data1.Recordset("id") & " and hc_scasi is null"
         Data3.Refresh
         If Data3.Recordset.RecordCount > 0 Then
            Data3.Recordset.Edit
            Data3.Recordset("hc_scasi") = 1
            Data3.Recordset.Update
         End If
'         data_param.Recordset.Edit
'         data_param.Recordset("p_linmmdd") = data_param.Recordset("p_linmmdd") + 1
'         data_param.Refresh
'         data_lin.RecordSource = "select * from linmmdd where cod_cli=" & Data1.Recordset("mat")
'         data_lin.Refresh
'         data_lin.Recordset.AddNew
'         data_lin.Recordset("factura") = data_param.Recordset("p_linmmdd")
'         data_lin.Recordset("tipo") = "REG."
'         data_lin.Recordset("realizada") = Data1.Recordset("fecha")
'         data_lin.Recordset("fecha") = Data1.Recordset("fecha")
'         data_lin.Recordset("cod_cli") = Data1.Recordset("mat")
'         data_lin.Recordset("nom_cli") = Mid(Data1.Recordset("cl_apellid"), 1, 30)
'         data_lin.Recordset("cod_prod") = 810
'         data_lin.Recordset("nom_prod") = "SERVICIO CONTROL AMBULATORIO"
'         data_lin.Recordset("cantidad") = 1
'         data_lin.Recordset("operador") = "AUT."
'         data_lin.Recordset("hora") = Format(Time, "HH:mm")
'         data_lin.Recordset("nro_flia") = 8
'         data_lin.Recordset("nom_flia") = "OTROS SERVICIOS"
'         data_lin.Recordset("linea") = 1
'         data_lin.Recordset("convenio") = Data1.Recordset("cl_codconv")
'         data_lin.Recordset("rub_cont") = 512118
'         data_lin.Recordset("pendiente") = "X"
'         data_lin.Recordset("arancel") = 0
'         data_lin.Recordset("imp_timbre") = 0
'         data_lin.Recordset("ced_socio") = Data1.Recordset("cl_cedula")
'         data_lin.Recordset("fact") = Data1.Recordset("cl_codced")
'         data_lin.Recordset("tot_lin") = 0
'         data_medicos.RecordSource = "select * from medicos where med_socnro =" & Data1.Recordset("hc_codmed")
'         data_medicos.Refresh
'         If data_medicos.Recordset.RecordCount > 0 Then
'            data_lin.Recordset("nro_med_a") = data_medicos.Recordset("med_cod")
'            data_lin.Recordset("nom_med_a") = data_medicos.Recordset("med_nombre")
'         Else
'            data_lin.Recordset("nro_med_a") = 440
'            data_lin.Recordset("nom_med_a") = "OTROS MEDICOS"
'         End If
'         data_lin.Recordset("precio_est") = 0
'         data_lin.Recordset("base") = 19
'         data_lin.Recordset("imp_iva") = 0
'         data_lin.Recordset.Update
         
      End If
      Data1.Recordset.MoveNext
   Loop
End If

Data4.RecordSource = "select * from pendiente_sca where fecha_cierre is null and (fecha_prox is null or fecha_prox =#" & Format(Date, "yyyy/mm/dd") & "#)"
Data4.Refresh
frm_pendsca.MousePointer = 0

End Sub

Private Sub List1_DblClick()
t_datos.Text = ""

If List1.ListCount > 0 Then
   data_verctrol.RecordSource = "select seguimiento_sca.fecha_altafin,seguimiento_sca.hora_altafin,seguimiento_sca.fecha,seguimiento_sca.hora,seguimiento_sca.mat,seguimiento_sca.medicocod,seguimiento_sca.obs,seguimiento_sca.nro_ctrol,seguimiento_sca.fecha_prox,seguimiento_sca.fecha_alta,seguimiento_sca.id_seguimiento,us.id,us.nombre,us.apellidos " & _
   "from seguimiento_sca inner join us on seguimiento_sca.medicocod=us.id where seguimiento_sca.nro_ctrol =" & Val(List1.List(List1.ListIndex)) & " and seguimiento_sca.mat =" & Data4.Recordset("mat") & " and seguimiento_sca.id_seguimiento =" & Data4.Recordset("id")
   data_verctrol.Refresh
   If data_verctrol.Recordset.RecordCount > 0 Then
      t_datos.Text = "FECHA:" & Format(data_verctrol.Recordset("fecha"), "dd/mm/yyyy") & " HORA:" & data_verctrol.Recordset("hora") & " MÉDICO:" & data_verctrol.Recordset("nombre") & " " & data_verctrol.Recordset("apellidos") & vbCrLf & data_verctrol.Recordset("obs") & " PROX.CONTROL: " & Format(data_verctrol.Recordset("fecha_prox"), "dd/mm/yyyy") & vbCrLf
      If IsNull(data_verctrol.Recordset("fecha_altafin")) = False Then
         If IsNull(data_verctrol.Recordset("hora_altafin")) = False Then
            t_datos.Text = t_datos.Text & vbCrLf & "FECHA y HORA ALTA de SAPP: " & Format(data_verctrol.Recordset("fecha_altafin"), "dd/mm/yyyy") & " " & data_verctrol.Recordset("hora_altafin")
         Else
            t_datos.Text = t_datos.Text & vbCrLf & "FECHA y HORA ALTA de SAPP: " & Format(data_verctrol.Recordset("fecha_altafin"), "dd/mm/yyyy")
         End If
      End If
      If IsNull(data_verctrol.Recordset("fecha_alta")) = False Then
         t_datos.Text = t_datos.Text & "ALTA: " & Format(data_verctrol.Recordset("fecha_alta"), "dd/mm/yyyy")
      End If
   End If
End If

End Sub

Private Sub mfctr_GotFocus()
If mfctr.Text <> "__/__/____" Then
Else
   mfctr.Text = Format(Date, "dd/mm/yyyy")
   mhctr.Text = Format(Time, "HH:mm")
   
End If

End Sub

Private Sub mfctr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfctr.SetFocus
End If

End Sub
Public Sub VerSegui()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
ConectarBD
ConbdSapp.Open
Xsqlpromo = "select * from seguimiento_sca where mat =" & Data4.Recordset("mat") & " and id_seguimiento =" & Data4.Recordset("id") & " order by nro_ctrol"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      List1.AddItem Xrecclii("nro_ctrol")
      Xrecclii.MoveNext
      NroControl = NroControl + 1
   Loop
End If
Xrecclii.Close
ConbdSapp.Close

End Sub

