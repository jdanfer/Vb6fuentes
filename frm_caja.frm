VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_caja 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caja diaria"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   Icon            =   "frm_caja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8445
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_caja 
      Height          =   495
      Left            =   3720
      Top             =   3000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      DataSourceName  =   "sappnew"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data_caja"
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
   Begin MSDBCtls.DBCombo dbcbonomrub 
      Bindings        =   "frm_caja.frx":0442
      Height          =   660
      Left            =   1320
      TabIndex        =   12
      Top             =   1440
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1164
      _Version        =   393216
      Style           =   1
      ListField       =   "NOMBRE"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txt_impiva 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   4680
      TabIndex        =   31
      Top             =   2520
      Width           =   975
   End
   Begin VB.ComboBox cboiva 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frm_caja.frx":045C
      Left            =   3480
      List            =   "frm_caja.frx":0469
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton btn_cierra 
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
      Left            =   3960
      Picture         =   "frm_caja.frx":047B
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Salir"
      Top             =   5040
      Width           =   615
   End
   Begin VB.Data data_parsec 
      Caption         =   "data_parsec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_usuar 
      Caption         =   "data_usuar"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\usapp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "usuarioact"
      Top             =   4320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton btn_ant 
      Caption         =   "<..Anterior"
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
      Left            =   6600
      TabIndex        =   26
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton btn_sig 
      Caption         =   "Siguiente..>"
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
      Left            =   6600
      TabIndex        =   25
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton btn_imprime 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Picture         =   "frm_caja.frx":0A05
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Informes"
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton btn_cancela 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Picture         =   "frm_caja.frx":0F8F
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Cancelar acción"
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton btn_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   1080
      Picture         =   "frm_caja.frx":1519
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Grabar"
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton btn_alta 
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
      Left            =   120
      Picture         =   "frm_caja.frx":1AA3
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Alta de nuevo registro"
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton btn_saldos 
      BackColor       =   &H00C0FFC0&
      Caption         =   "SALDOS..."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Data data_rubros 
      Caption         =   "data_rubros"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from COD_CAJA order by nombre"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txt_obs2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   4080
      Width           =   4215
   End
   Begin VB.TextBox txt_obs1 
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
      Left            =   1680
      MaxLength       =   120
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   3480
      Width           =   4215
   End
   Begin VB.TextBox txt_nrobol 
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
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   16
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txt_imp 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   14346
         SubFormatType   =   1
      EndProperty
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
      Left            =   1680
      TabIndex        =   14
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txt_rubro 
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
      Left            =   1320
      TabIndex        =   11
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txt_base 
      Alignment       =   2  'Center
      Enabled         =   0   'False
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
      Left            =   6600
      TabIndex        =   7
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txt_hora 
      Enabled         =   0   'False
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
      Left            =   4680
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txt_fecha 
      Enabled         =   0   'False
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
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C00000&
      Caption         =   "I.V.A."
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
      TabIndex        =   29
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label labsaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6360
      TabIndex        =   27
      Top             =   3840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   8400
      X2              =   0
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   8400
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C00000&
      Caption         =   "Observ.:"
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
      Height          =   615
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C00000&
      Caption         =   "No.Boleta:"
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
      Left            =   240
      TabIndex        =   15
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C00000&
      Caption         =   "Importe($):"
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
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "Rubro:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   6120
      X2              =   8400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   6120
      X2              =   6120
      Y1              =   1200
      Y2              =   6000
   End
   Begin VB.Label labusurea 
      BackColor       =   &H00FFFFC0&
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
      Left            =   6360
      TabIndex        =   9
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      Caption         =   "Realizado por:"
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
      Left            =   6360
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   8400
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      Caption         =   "Base:"
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
      Left            =   5760
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "Hora:"
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
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Fecha:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   8400
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label labusuario 
      BackColor       =   &H00C00000&
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
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Usuario Actual:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   600
      Picture         =   "frm_caja.frx":202D
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1575
   End
End
Attribute VB_Name = "frm_caja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_alta_Click()
'data_caja.RecordSource = "CAJA"
'data_caja.Refresh
habcaja
btn_cancela.Enabled = True
btn_graba.Enabled = True
btn_imprime.Enabled = False
btn_alta.Enabled = False
btn_sig.Enabled = False
btn_ant.Enabled = False
btn_saldos.Enabled = False
labusurea.Caption = ""
txt_fecha.Text = ""
txt_hora.Text = ""
txt_base.Text = ""
labusurea.Caption = ""
txt_rubro.Text = ""
dbcbonomrub.Text = ""
txt_imp.Text = ""
txt_nrobol.Text = ""
txt_obs1.Text = ""
txt_obs2.Text = ""
cboiva.ListIndex = 0
txt_impiva.Text = 0
txt_fecha.Text = Format(Date, "dd/mm/yyyy")
txt_hora.Text = Format(Time, "HH:mm")
txt_base.Text = data_parsec.Recordset("base")
labusuario.Caption = WElusuario
dbcbonomrub.SetFocus
data_caja.Recordset.AddNew


End Sub

Private Sub btn_ant_Click()
On Error GoTo Queerr

data_caja.Recordset.MovePrevious
If data_caja.Recordset.BOF = False Then
    txt_fecha.Text = Format(data_caja.Recordset("fecha"), "dd/mm/yyyy")
    txt_hora.Text = Format(data_caja.Recordset("hora"), "HH:mm")
    txt_base.Text = data_caja.Recordset("base")
    labusurea.Caption = data_caja.Recordset("usuario")
    txt_rubro.Text = data_caja.Recordset("numero")
    If IsNull(data_caja.Recordset("nombre")) = True Then
       dbcbonomrub.Text = ""
    Else
       dbcbonomrub.Text = data_caja.Recordset("nombre")
    End If
    txt_imp.Text = data_caja.Recordset("imp_fact")
    If IsNull(data_caja.Recordset("documento")) = True Then
        txt_nrobol.Text = 0
    Else
        txt_nrobol.Text = data_caja.Recordset("documento")
    End If
'    txt_nrobol.Text = data_caja.Recordset("documento")
    If IsNull(data_caja.Recordset("opiva")) = False Then
       If data_caja.Recordset("opiva") = 2 Then
          cboiva.ListIndex = 2
          txt_impiva.Text = Format(data_caja.Recordset("imp_iva"), "Standard")
       Else
          If data_caja.Recordset("opiva") = 1 Then
             cboiva.ListIndex = 1
             txt_impiva.Text = Format(data_caja.Recordset("imp_iva"), "Standard")
          Else
             cboiva.ListIndex = 0
             txt_impiva.Text = 0
          End If
       End If
    End If
    If IsNull(data_caja.Recordset("observ")) = False Then
       txt_obs1.Text = data_caja.Recordset("observ")
    Else
       txt_obs1.Text = ""
    End If
    If IsNull(data_caja.Recordset("nom_serv")) = False Then
       txt_obs2.Text = data_caja.Recordset("nom_serv")
       txt_obs2.Text = txt_obs2.Text + " Socio: " + data_caja.Recordset("nom_socio")
    Else
       txt_obs2.Text = ""
    End If
Else
    MsgBox "Comienzo de archivo", vbInformation, "Caja"
    data_caja.Recordset.MoveNext
End If

Exit Sub

Queerr:
       If Err.Number = 3146 Then
          MsgBox "Error al visualizar"
       Else
          MsgBox "Error al visualizar registro"
       End If

End Sub

Private Sub btn_cancela_Click()
data_caja.Recordset.CancelUpdate
btn_cancela.Enabled = False
btn_graba.Enabled = False
btn_imprime.Enabled = True
btn_alta.Enabled = True
btn_sig.Enabled = True
btn_ant.Enabled = True
btn_saldos.Enabled = True
labusurea.Caption = ""
txt_fecha.Text = ""
txt_hora.Text = ""
txt_base.Text = ""
labusurea.Caption = ""
txt_rubro.Text = ""
dbcbonomrub.Text = ""
txt_imp.Text = ""
txt_nrobol.Text = ""
txt_obs1.Text = ""
txt_obs2.Text = ""
cboiva.ListIndex = 0
txt_impiva.Text = 0
descaja
btn_alta.SetFocus
'data_caja.RecordSource = "Select * from caja order by fecha"
'data_caja.Refresh
'data_caja.Recordset.MoveLast

End Sub

Private Sub btn_cierra_Click()
Unload Me

End Sub

Private Sub btn_graba_Click()
Dim Xlafedelacaja2 As Date
Xlafedelacaja2 = Date - 31

If dbcbonomrub.Text <> "" And txt_rubro.Text <> "" Then
   If txt_rubro.Text = 112001 Or _
      txt_rubro.Text = 112002 Or _
      txt_rubro.Text = 112022 Or _
      txt_rubro.Text = 112004 Or _
      txt_rubro.Text = 112026 Or _
      txt_rubro.Text = 112006 Or _
      txt_rubro.Text = 112032 Or _
      txt_rubro.Text = 112033 Or _
      txt_rubro.Text = 112042 Or _
      txt_rubro.Text = 112044 Or _
      txt_rubro.Text = 112011 Or _
      txt_rubro.Text = 112110 Or _
      txt_rubro.Text = 112113 Or _
      txt_rubro.Text = 112116 Or _
      txt_rubro.Text = 211408 Or _
      txt_rubro.Text = 211314 Then
      If txt_imp.Text < 0 Then
           data_caja.Recordset("fecha") = Format(txt_fecha.Text, "dd-mm-yyyy")
           data_caja.Recordset("hora") = txt_hora.Text
           data_caja.Recordset("base") = txt_base.Text
           data_caja.Recordset("numero") = txt_rubro.Text
           data_caja.Recordset("moneda") = "$"
           data_caja.Recordset("nombre") = dbcbonomrub.Text
           data_caja.Recordset("movimiento") = data_rubros.Recordset("movimiento")
           data_caja.Recordset("imp_fact") = Format(txt_imp.Text, "Standard")
           If cboiva.ListIndex = 0 Then
              txt_impiva.Text = 0
              data_caja.Recordset("imp_iva") = 0
              data_caja.Recordset("opiva") = 0
           Else
              If cboiva.ListIndex = 1 Then
                 txt_impiva.Text = txt_imp.Text / 1.1
                 txt_impiva.Text = txt_impiva.Text * 0.1
                 data_caja.Recordset("imp_iva") = Format(txt_impiva.Text, "Standard")
                 data_caja.Recordset("opiva") = 1
              Else
                 If cboiva.ListIndex = 2 Then
                    txt_impiva.Text = txt_imp.Text / 1.22
                    txt_impiva.Text = txt_impiva.Text * 0.22
                    data_caja.Recordset("imp_iva") = Format(txt_impiva.Text, "Standard")
                    data_caja.Recordset("opiva") = 2
                 End If
              End If
           End If
           If txt_nrobol.Text <> "" Then
              data_caja.Recordset("documento") = txt_nrobol.Text
           Else
''              data_caja.Recordset("documento") = data_parsec.Recordset("rojo") + 1
              data_caja.Recordset("documento") = 0
           End If
           data_caja.Recordset("observ") = txt_obs1.Text
           data_caja.Recordset("saldo") = Format(txt_imp.Text, "Standard")
           data_caja.Recordset("usuario") = labusuario.Caption
           data_caja.Recordset("saldo_user") = Format(txt_imp.Text, "Standard")
           data_caja.Recordset.Update
''           data_parsec.Recordset.Edit
''           data_parsec.Recordset("rojo") = data_parsec.Recordset("rojo") + 1
''           data_parsec.Recordset.Update
           
'           data_caja.Recordset.MoveLast
           data_caja.Refresh
'           vercaja
           descaja
           btn_cancela.Enabled = False
           btn_graba.Enabled = False
           btn_imprime.Enabled = True
           btn_alta.Enabled = True
           btn_sig.Enabled = True
           btn_ant.Enabled = True
           btn_saldos.Enabled = True
           btn_alta.SetFocus
'           data_caja.RecordSource = "Select * from caja where base =" & data_parsec.Recordset("base") & " and fecha >=#" & Format(Xlafedelacaja2, "yyyy/mm/dd") & "# order by fecha"
'           data_caja.Refresh
'           data_caja.Recordset.MoveLast
      Else
          MsgBox "Debe de ingresar importe negativo", vbCritical, "Mensaje"
          txt_imp.SetFocus
      End If
   Else
       data_caja.Recordset("fecha") = Format(txt_fecha.Text, "dd-mm-yyyy")
       data_caja.Recordset("hora") = txt_hora.Text
       data_caja.Recordset("base") = txt_base.Text
       data_caja.Recordset("numero") = txt_rubro.Text
       data_caja.Recordset("moneda") = "$"
       data_caja.Recordset("nombre") = dbcbonomrub.Text
       data_caja.Recordset("movimiento") = data_rubros.Recordset("movimiento")
       data_caja.Recordset("imp_fact") = Format(txt_imp.Text, "Standard")
       If cboiva.ListIndex = 0 Then
          txt_impiva.Text = 0
          data_caja.Recordset("imp_iva") = 0
          data_caja.Recordset("opiva") = 0
       Else
          If cboiva.ListIndex = 1 Then
             txt_impiva.Text = txt_imp.Text / 1.1
             txt_impiva.Text = txt_impiva.Text * 0.1
             data_caja.Recordset("imp_iva") = Format(txt_impiva.Text, "Standard")
             data_caja.Recordset("opiva") = 1
          Else
             If cboiva.ListIndex = 2 Then
                txt_impiva.Text = txt_imp.Text / 1.22
                txt_impiva.Text = txt_impiva.Text * 0.22
                data_caja.Recordset("imp_iva") = Format(txt_impiva.Text, "Standard")
                data_caja.Recordset("opiva") = 2
             End If
          End If
       End If
       If txt_nrobol.Text <> "" Then
          data_caja.Recordset("documento") = txt_nrobol.Text
       Else
'''          data_caja.Recordset("documento") = data_parsec.Recordset("rojo") + 1
          data_caja.Recordset("documento") = 0
       End If
       data_caja.Recordset("observ") = txt_obs1.Text
       data_caja.Recordset("saldo") = Format(txt_imp.Text, "Standard")
       data_caja.Recordset("usuario") = labusuario.Caption
       data_caja.Recordset("saldo_user") = Format(txt_imp.Text, "Standard")
       data_caja.Recordset.Update
''       data_parsec.Recordset.Edit
''       data_parsec.Recordset("rojo") = data_parsec.Recordset("rojo") + 1
''       data_parsec.Recordset.Update
       
'       data_caja.Recordset.MoveLast
       data_caja.Refresh
'       vercaja
       descaja
       btn_cancela.Enabled = False
       btn_graba.Enabled = False
       btn_imprime.Enabled = True
       btn_alta.Enabled = True
       btn_sig.Enabled = True
       btn_ant.Enabled = True
       btn_saldos.Enabled = True
       btn_alta.SetFocus
'       data_caja.RecordSource = "Select * from caja where base =" & data_parsec.Recordset("base") & " and fecha >=#" & Format(Xlafedelacaja2, "yyyy/mm/dd") & "# order by fecha"
'       data_caja.Refresh
'       data_caja.Recordset.MoveLast
   End If
Else
   MsgBox "Verifique el rubro", vbInformation, "Caja"
   dbcbonomrub.SetFocus
End If

End Sub

Private Sub btn_imprime_Click()
btn_imprime.Enabled = False
Dim Lafecdecaja As Date
Lafecdecaja = Date - 31
'If txt_fecha.Text <> "" Then
   
'   data_caja.RecordSource = "Select * from caja where fecha = #" & Format(txt_fecha.Text, "yyyy/mm/dd") & "# And usuario ='" & labusuario.Caption & "' order by fecha"
'   data_caja.Refresh
'Else
   data_caja.RecordSource = "Select * from caja where fecha = '" & Format(Date, "yyyy-mm-dd") & "' And usuario ='" & WElusuario & "' order by fecha"
   data_caja.Refresh
'End If
If data_caja.Recordset.RecordCount > 0 Then
   data_caja.Recordset.MoveFirst
   Do While Not data_caja.Recordset.EOF
      If data_caja.Recordset("movimiento") = "EGRESO" Then
         Xsaldcaj = Xsaldcaj - data_caja.Recordset("imp_fact")
      Else
         If data_caja.Recordset("movimiento") = "INGRESO" Then
            Xsaldcaj = Xsaldcaj + data_caja.Recordset("imp_fact")
         End If
      End If
'      data_caja.Recordset.Edit
'      data_caja.Recordset("saldo_user") = Xsaldcaj
'      data_caja.Recordset.Update
      data_caja.Recordset.MoveNext
   Loop
'   data_caja.Recordset.MoveFirst
'   Do While Not data_caja.Recordset.EOF
'      data_caja.Recordset.Edit
'      data_caja.Recordset("saldo") = Xsaldcaj
'      data_caja.Recordset.Update
'      data_caja.Recordset.MoveNext
'   Loop
   labsaldo.Visible = True
   labsaldo.Caption = Format(Xsaldcaj, "Currency")
Else
   MsgBox "No existen registros", vbInformation, "Caja"
End If
'data_caja.RecordSource = "Select * from caja order by fecha"
data_caja.RecordSource = "Select * from caja where base =" & data_parsec.Recordset("base") & " and fecha >='" & Format(Lafecdecaja, "yyyy-mm-dd") & "' order by fecha DESC"

data_caja.Refresh
'data_caja.Recordset.MoveLast
btn_imprime.Enabled = True

frm_infcaja.Show vbModal

End Sub

Private Sub btn_saldos_Click()
Dim Xsaldcaj As Long
Dim Xlafedelacaja As Date
Xlafedelacaja = Date - 31
data_caja.RecordSource = "Select * from caja where fecha = '" & Format(txt_fecha.Text, "yyyy-mm-dd") & "' And usuario ='" & labusuario.Caption & "' And base =" & data_parsec.Recordset("base") & " order by fecha"
data_caja.Refresh
If data_caja.Recordset.RecordCount > 0 Then
   data_caja.Recordset.MoveFirst
   Do While Not data_caja.Recordset.EOF
      If data_caja.Recordset("movimiento") = "EGRESO" Then
         Xsaldcaj = Xsaldcaj - data_caja.Recordset("imp_fact")
      Else
         If data_caja.Recordset("movimiento") = "INGRESO" Then
            Xsaldcaj = Xsaldcaj + data_caja.Recordset("imp_fact")
         End If
      End If
      If Format(data_caja.Recordset("saldo_user"), "Standard") = Format(Xsaldcaj, "Standard") Then
         data_caja.Recordset.MoveNext
      Else
'         data_caja.Recordset.Edit
         data_caja.Recordset("saldo_user") = Xsaldcaj
         data_caja.Recordset.Update
         data_caja.Recordset.MoveNext
      End If
   Loop
   data_caja.Recordset.MoveFirst
   Do While Not data_caja.Recordset.EOF
      If Format(data_caja.Recordset("saldo"), "Standard") = Format(Xsaldcaj, "Standard") Then
         data_caja.Recordset.MoveNext
      Else
'         data_caja.Recordset.Edit
         data_caja.Recordset("saldo") = Xsaldcaj
         data_caja.Recordset.Update
         data_caja.Recordset.MoveNext
      End If
   Loop
   labsaldo.Visible = True
   labsaldo.Caption = Format(Xsaldcaj, "Currency")
Else
   MsgBox "No existen registros", vbInformation, "Caja"
End If
data_caja.RecordSource = "Select * from caja where base =" & data_parsec.Recordset("base") & " and fecha >='" & Format(Xlafedelacaja, "yyyy-mm-dd") & "' order by fecha DESC"

'data_caja.RecordSource = "Select * from caja order by fecha"
data_caja.Refresh
'data_caja.Recordset.MoveLast

End Sub

Private Sub btn_sig_Click()
On Error GoTo Queerr

data_caja.Recordset.MoveNext
If data_caja.Recordset.EOF = False Then
    txt_fecha.Text = Format(data_caja.Recordset("fecha"), "dd/mm/yyyy")
    txt_hora.Text = Format(data_caja.Recordset("hora"), "HH:mm")
    txt_base.Text = data_caja.Recordset("base")
    labusurea.Caption = data_caja.Recordset("usuario")
    txt_rubro.Text = data_caja.Recordset("numero")
    dbcbonomrub.Text = data_caja.Recordset("nombre")
    txt_imp.Text = data_caja.Recordset("imp_fact")
    If IsNull(data_caja.Recordset("documento")) = True Then
        txt_nrobol.Text = 0
    Else
        txt_nrobol.Text = data_caja.Recordset("documento")
    End If
'    txt_nrobol.Text = data_caja.Recordset("documento")
    If IsNull(data_caja.Recordset("opiva")) = False Then
       If data_caja.Recordset("opiva") = 2 Then
          cboiva.ListIndex = 2
          txt_impiva.Text = Format(data_caja.Recordset("imp_iva"), "Standard")
       Else
          If data_caja.Recordset("opiva") = 1 Then
             cboiva.ListIndex = 1
             txt_impiva.Text = Format(data_caja.Recordset("imp_iva"), "Standard")
          Else
             cboiva.ListIndex = 0
             txt_impiva.Text = 0
          End If
       End If
    End If
    If IsNull(data_caja.Recordset("observ")) = False Then
       txt_obs1.Text = data_caja.Recordset("observ")
    Else
       txt_obs1.Text = ""
    End If
    If IsNull(data_caja.Recordset("nom_serv")) = False Then
       txt_obs2.Text = data_caja.Recordset("nom_serv")
       txt_obs2.Text = txt_obs2.Text + " Socio: " + data_caja.Recordset("nom_socio")
    Else
       txt_obs2.Text = ""
    End If
Else
    MsgBox "Final de registros", vbInformation, "Caja"
    data_caja.Recordset.MovePrevious
End If

Exit Sub

Queerr:
       If Err.Number = 3146 Then
          MsgBox "Error al visualizar"
       Else
          MsgBox "Error al visualizar registro"
       End If
       
End Sub

Private Sub cboiva_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nrobol.SetFocus
End If

End Sub

Private Sub cboiva_LostFocus()
If cboiva.ListIndex = 1 Then
   txt_impiva.Text = txt_imp.Text / 1.1
   txt_impiva.Text = txt_impiva.Text * 0.1
   txt_impiva.Text = Format(txt_impiva.Text, "Standard")
Else
   If cboiva.ListIndex = 2 Then
      txt_impiva.Text = txt_imp.Text / 1.22
      txt_impiva.Text = txt_impiva.Text * 0.22
      txt_impiva.Text = Format(txt_impiva.Text, "Standard")
   End If
End If

End Sub

Private Sub dbcbonomrub_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If dbcbonomrub.Text <> "" Then
        dbcbonomrub.ListField = "nombre"
        dbcbonomrub.BoundColumn = "nombre"
        If IsNumeric(dbcbonomrub.Text) Then
'           data_rubros.Recordset.FindFirst "numero =" & dbcbonomrub.Text
           data_rubros.RecordSource = "Select * from cod_caja where numero =" & dbcbonomrub.Text
           data_rubros.Refresh
           If data_rubros.Recordset.RecordCount > 0 Then
              dbcbonomrub.Text = data_rubros.Recordset("nombre")
              txt_rubro.Text = data_rubros.Recordset("numero")
              txt_imp.SetFocus
              dbcbonomrub.Height = 500
              dbcbonomrub.ListField = ""
              dbcbonomrub.BoundColumn = ""
           Else
              data_rubros.RecordSource = "select * from cod_caja where numero >=" & dbcbonomrub.Text
              data_rubros.Refresh
              dbcbonomrub.Height = 2350
           End If
        Else
'           data_rubros.Recordset.FindFirst "nombre ='" & dbcbonomrub.Text & "'"
           data_rubros.RecordSource = "Select * from cod_caja where nombre='" & dbcbonomrub.Text & "'"
           data_rubros.Refresh
           If data_rubros.Recordset.RecordCount > 0 Then
              dbcbonomrub.Text = data_rubros.Recordset("nombre")
              txt_rubro.Text = data_rubros.Recordset("numero")
              dbcbonomrub.Height = 500
              dbcbonomrub.ListField = ""
              dbcbonomrub.BoundColumn = ""
              txt_imp.SetFocus
           Else
              data_rubros.RecordSource = "select * from cod_caja where nombre >='" & dbcbonomrub.Text & "' order by nombre"
              data_rubros.Refresh
              dbcbonomrub.Height = 2350
           End If
        End If
   Else
       btn_cancela.SetFocus
   End If
   If txt_rubro.Text <> "" Then
      If txt_rubro.Text = 513018 Or _
         txt_rubro.Text = 211408 Then
         MsgBox "Rubro incorrecto", vbCritical, "Mensaje"
         dbcbonomrub.Text = ""
         txt_rubro.Text = ""
         dbcbonomrub.SetFocus
         
      End If
   End If
End If

End Sub

Private Sub Form_Activate()
btn_alta.SetFocus

End Sub

Private Sub Form_Load()
Dim Xlafedelacaja As Date
Xlafedelacaja = Date - 10
data_caja.ConnectionString = "dsn=" & Xconexrmt
'data_caja.DatabaseName = App.Path & "\sapp.mdb"
'data_caja.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_caja.RecordSource = "caja"
'data_caja.Refresh
data_parsec.DatabaseName = App.path & "\parse.mdb"
data_parsec.RecordSource = "PARSEC0"
data_parsec.Refresh

data_rubros.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_rubros.RecordSource = "cod_caja"
data_rubros.Refresh

data_usuar.DatabaseName = "C:\WINDOWS\usapp.mdb"
data_usuar.RecordSource = "usuarioact"
data_usuar.Refresh

dbcbonomrub.ListField = ""
dbcbonomrub.BoundColumn = ""
data_caja.RecordSource = "Select * from caja where base =" & data_parsec.Recordset("base") & " and fecha >='" & Format(Xlafedelacaja, "yyyy-mm-dd") & "' order by fecha DESC"
data_caja.Refresh
If data_caja.Recordset.RecordCount > 0 Then
'    data_caja.Recordset.MoveLast
    txt_fecha.Text = Format(data_caja.Recordset("fecha"), "dd/mm/yyyy")
    txt_hora.Text = Format(data_caja.Recordset("hora"), "HH:mm")
    txt_base.Text = data_caja.Recordset("base")
    labusurea.Caption = data_caja.Recordset("usuario")
    If IsNull(data_caja.Recordset("numero")) = False Then
       txt_rubro.Text = data_caja.Recordset("numero")
    Else
       txt_rubro.Text = 0
    End If
    If IsNull(data_caja.Recordset("nombre")) = False Then
       dbcbonomrub.Text = data_caja.Recordset("nombre")
    Else
       dbcbonomrub.Text = ""
    End If
    txt_imp.Text = data_caja.Recordset("imp_fact")
    If IsNull(data_caja.Recordset("documento")) = True Then
       txt_nrobol.Text = 0
    Else
       txt_nrobol.Text = data_caja.Recordset("documento")
    End If
    
    If IsNull(data_caja.Recordset("opiva")) = False Then
       If data_caja.Recordset("opiva") = 2 Then
          cboiva.ListIndex = 2
          txt_impiva.Text = Format(data_caja.Recordset("imp_iva"), "Standard")
       Else
          If data_caja.Recordset("opiva") = 1 Then
             cboiva.ListIndex = 1
             txt_impiva.Text = Format(data_caja.Recordset("imp_iva"), "Standard")
          Else
             cboiva.ListIndex = 0
             txt_impiva.Text = 0
          End If
       End If
    End If
    If IsNull(data_caja.Recordset("observ")) = False Then
       txt_obs1.Text = data_caja.Recordset("observ")
    Else
       txt_obs1.Text = ""
    End If
    If IsNull(data_caja.Recordset("nom_serv")) = False Then
       txt_obs2.Text = data_caja.Recordset("nom_serv")
       If IsNull(data_caja.Recordset("nom_socio")) = False Then
          txt_obs2.Text = txt_obs2.Text + " Socio: " + data_caja.Recordset("nom_socio")
       Else
          txt_obs2.Text = txt_obs2.Text + " Socio: " + "NN"
       End If
    Else
       txt_obs2.Text = ""
    End If
Else
    MsgBox "No existen registros con ésta BASE, REALICE ALGUNA FACTURACION DESDE LA FICHA DE SOCIOS y luego vuelva a ingresar.", vbInformation, "CAJA"
End If

data_usuar.Recordset.MoveFirst
labusuario.Caption = WElusuario

descaja


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub txt_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_rubro.SetFocus
End If

End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_hora.SetFocus
End If

End Sub

Private Sub txt_hora_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_base.SetFocus
End If

End Sub

Private Sub txt_imp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboiva.ListIndex = 0
   cboiva.SetFocus
End If

End Sub

Private Sub txt_nrobol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_obs1.SetFocus
End If

End Sub

Private Sub txt_obs1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   btn_graba.SetFocus
End If

End Sub

Private Sub txt_obs2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   btn_graba.SetFocus
End If

End Sub

Private Sub txt_rubro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_imp.SetFocus
End If

End Sub

Public Function habcaja()
'txt_fecha.Enabled = True
'txt_hora.Enabled = True
'txt_base.Enabled = True
txt_rubro.Enabled = True
dbcbonomrub.Enabled = True
txt_imp.Enabled = True
txt_impiva.Enabled = True
cboiva.Enabled = True
txt_nrobol.Enabled = True
txt_obs1.Enabled = True
txt_obs2.Enabled = True

End Function

Public Function descaja()
txt_fecha.Enabled = False
txt_hora.Enabled = False
txt_base.Enabled = False
txt_rubro.Enabled = False
dbcbonomrub.Enabled = False
txt_imp.Enabled = False
txt_impiva.Enabled = False
cboiva.Enabled = False
txt_nrobol.Enabled = False
txt_obs1.Enabled = False
txt_obs2.Enabled = False

End Function

Public Function vercaja()
If data_caja.Recordset.EOF = True Then
   data_caja.Recordset.MovePrevious
   txt_fecha.Text = Format(data_caja.Recordset("fecha"), "dd/mm/yyyy")
Else
   data_caja.Recordset.MovePrevious
   txt_fecha.Text = Format(data_caja.Recordset("fecha"), "dd/mm/yyyy")
End If

txt_hora.Text = Format(data_caja.Recordset("hora"), "HH:mm")
txt_base.Text = data_caja.Recordset("base")
labusurea.Caption = data_caja.Recordset("usuario")
txt_rubro.Text = data_caja.Recordset("numero")
dbcbonomrub.Text = data_caja.Recordset("nombre")
txt_imp.Text = data_caja.Recordset("imp_fact")
If IsNull(data_caja.Recordset("documento")) = True Then
   txt_nrobol.Text = 0
Else
   txt_nrobol.Text = data_caja.Recordset("documento")
End If
If IsNull(data_caja.Recordset("opiva")) = False Then
   If data_caja.Recordset("opiva") = 2 Then
      cboiva.ListIndex = 2
      txt_impiva.Text = Format(data_caja.Recordset("imp_iva"), "Standard")
   Else
      If data_caja.Recordset("opiva") = 1 Then
         cboiva.ListIndex = 1
         txt_impiva.Text = Format(data_caja.Recordset("imp_iva"), "Standard")
      Else
         cboiva.ListIndex = 0
         txt_impiva.Text = 0
      End If
   End If
End If
If IsNull(data_caja.Recordset("observ")) = False Then
   txt_obs1.Text = data_caja.Recordset("observ")
Else
   txt_obs1.Text = ""
End If
If IsNull(data_caja.Recordset("nom_serv")) = False Then
   txt_obs2.Text = data_caja.Recordset("nom_serv")
   txt_obs2.Text = txt_obs2.Text + " Socio: " + data_caja.Recordset("nom_socio")
Else
   txt_obs2.Text = ""
End If
'data_usuar.Recordset.MoveFirst
labusuario.Caption = WElusuario

End Function
