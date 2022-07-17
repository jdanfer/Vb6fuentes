VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_buscasocesp 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de socios"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10470
   Icon            =   "frm_buscasocesp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   10470
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_deuda 
      Height          =   375
      Left            =   3120
      Top             =   4440
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
      DataSourceName  =   "sappnew"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data_deuda"
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   240
      Top             =   4440
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
      DataSourceName  =   "sappnew"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data1"
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
   Begin MSFlexGridLib.MSFlexGrid DBGrid1 
      Height          =   3375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5953
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.Data data_u 
      Caption         =   "data_u"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
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
      Left            =   9720
      Picture         =   "frm_buscasocesp.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txt_busc 
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
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   5535
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
      ItemData        =   "frm_buscasocesp.frx":09CC
      Left            =   2160
      List            =   "frm_buscasocesp.frx":09D9
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Buscar por:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   7440
      Picture         =   "frm_buscasocesp.frx":09FA
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1335
   End
End
Attribute VB_Name = "frm_buscasocesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_busc.SetFocus
End If

End Sub

Private Sub Command1_Click()
Xdeb = 0

Unload Me

End Sub

Private Sub DBGrid1_DblClick()
'data_deuda.Connect = "odbc;dsn=" & Xconexrmt & ";"

Dim Xsideuda As Integer
Xsideuda = 0
If Xdeb = 15 Then
   data1.RecordSource = "Select * from clientes where cl_codigo =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 0)
   data1.Refresh
   If data1.Recordset.RecordCount > 0 Then
        frm_especialistas.t_mat.Text = data1.Recordset("cl_codigo")
        If IsNull(data1.Recordset("cl_apellid")) = False Then
           frm_especialistas.t_nompac.Text = data1.Recordset("cl_apellid")
        Else
           frm_especialistas.t_nompac.Text = ""
        End If
        If IsNull(data1.Recordset("cl_cedula")) = False Then
           frm_especialistas.t_ced.Text = Trim(Str(data1.Recordset("cl_cedula")))
        End If
        If IsNull(data1.Recordset("cl_codced")) = False Then
           frm_especialistas.t_codced.Text = Trim(Str(data1.Recordset("cl_codced")))
        End If
        If IsNull(data1.Recordset("cl_codconv")) = False Then
           frm_especialistas.t_conv.Text = data1.Recordset("cl_codconv")
        End If
        If IsNull(data1.Recordset("cl_telefon")) = False Then
           frm_especialistas.t_tellinea.Text = data1.Recordset("cl_telefon")
        End If
        If IsNull(data1.Recordset("cl_dpto")) = False Then
           frm_especialistas.t_celu.Text = data1.Recordset("cl_dpto")
        End If
        If IsNull(data1.Recordset("cl_fnac")) = False Then
           frm_especialistas.mfnac.Text = data1.Recordset("cl_fnac")
        End If
   End If
   Unload Me
Else
End If
Xdeb = 0

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   DBGrid1_DblClick
End If

End Sub

Private Sub Form_Load()
'SelectLimit 20
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_deuda.ConnectionString = "dsn=" & Xconexrmt
data1.ConnectionString = "dsn=" & Xconexrmt
'Data1.RecordSource = "clientes"
'Data1.Refresh
'SelectLimit 0
Combo1.ListIndex = 0
DBGrid1.Rows = 2
DBGrid1.Cols = 7
DBGrid1.TextMatrix(0, 0) = "MATRICULA"
DBGrid1.ColWidth(0) = 1300
DBGrid1.TextMatrix(0, 1) = "NOMBRES"
DBGrid1.ColWidth(1) = 2900
DBGrid1.TextMatrix(0, 2) = "TELEFONO"
DBGrid1.ColWidth(2) = 1600
DBGrid1.TextMatrix(0, 3) = "CEDULA"
DBGrid1.ColWidth(3) = 1500
DBGrid1.TextMatrix(0, 4) = "DIG"
DBGrid1.ColWidth(4) = 400
DBGrid1.TextMatrix(0, 5) = "COD.CONV"
DBGrid1.ColWidth(5) = 1200
DBGrid1.TextMatrix(0, 6) = "NOMBRE CONV."
DBGrid1.ColWidth(6) = 2400

End Sub

Private Sub Form_Resize()
With Image1
    .Top = 0
    .Left = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub txt_busc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Combo1.ListIndex = 0 Then
      data1.RecordSource = "select * from clientes where cl_apellid >='" & txt_busc.Text & "' order by cl_apellid limit 80"
      data1.Refresh
   Else
      If Combo1.ListIndex = 1 Then
         data1.RecordSource = "select * from clientes where cl_cedula =" & Val(txt_busc.Text) & " order by cl_cedula"
         data1.Refresh
      Else
         If Combo1.ListIndex = 2 Then
            data1.RecordSource = "select * from clientes where cl_telefon >=" & Val(txt_busc.Text) & " order by cl_telefon limit 50"
            data1.Refresh
         End If
      End If
   End If
    DBGrid1.Rows = 2
    DBGrid1.Cols = 7
    DBGrid1.TextMatrix(0, 0) = "MATRICULA"
    DBGrid1.ColWidth(0) = 1300
    DBGrid1.TextMatrix(0, 1) = "NOMBRES"
    DBGrid1.ColWidth(1) = 2900
    DBGrid1.TextMatrix(0, 2) = "TELEFONO"
    DBGrid1.ColWidth(2) = 1600
    DBGrid1.TextMatrix(0, 3) = "CEDULA"
    DBGrid1.ColWidth(3) = 1500
    DBGrid1.TextMatrix(0, 4) = "DIG"
    DBGrid1.ColWidth(4) = 400
    DBGrid1.TextMatrix(0, 5) = "COD.CONV"
    DBGrid1.ColWidth(5) = 1200
    DBGrid1.TextMatrix(0, 6) = "NOMBRE CONV."
    DBGrid1.ColWidth(6) = 2400
   
    Dim Xcann As Integer
    Xcann = 1
    If data1.Recordset.RecordCount > 0 Then
        data1.Recordset.MoveFirst
        Do While Not data1.Recordset.EOF
           DBGrid1.TextMatrix(Xcann, 0) = data1.Recordset("cl_codigo")
           If IsNull(data1.Recordset("cl_apellid")) = False Then
              DBGrid1.TextMatrix(Xcann, 1) = data1.Recordset("cl_apellid")
           End If
           If IsNull(data1.Recordset("cl_telefon")) = False Then
              DBGrid1.TextMatrix(Xcann, 2) = data1.Recordset("cl_telefon")
           End If
           If IsNull(data1.Recordset("cl_cedula")) = False Then
              DBGrid1.TextMatrix(Xcann, 3) = data1.Recordset("cl_cedula")
           End If
           If IsNull(data1.Recordset("cl_codced")) = False Then
              DBGrid1.TextMatrix(Xcann, 4) = data1.Recordset("cl_codced")
           End If
           If IsNull(data1.Recordset("cl_codconv")) = False Then
              DBGrid1.TextMatrix(Xcann, 5) = data1.Recordset("cl_codconv")
           End If
           If IsNull(data1.Recordset("cl_nomconv")) = False Then
              DBGrid1.TextMatrix(Xcann, 6) = data1.Recordset("cl_nomconv")
           End If
           DBGrid1.Rows = DBGrid1.Rows + 1
           data1.Recordset.MoveNext
           Xcann = Xcann + 1
        Loop
    End If
   
   DBGrid1.SetFocus
End If

End Sub
