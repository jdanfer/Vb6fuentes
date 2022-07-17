VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_conscliencu 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar clientes"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   375
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
   Icon            =   "frm_conscliencu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8310
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   1920
      Top             =   4080
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
   Begin MSFlexGridLib.MSFlexGrid DBGrid1 
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5318
      _Version        =   393216
      BackColorBkg    =   12615680
      SelectionMode   =   1
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7440
      Picture         =   "frm_conscliencu.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4695
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "frm_conscliencu.frx":09CC
      Left            =   2400
      List            =   "frm_conscliencu.frx":09D6
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Doble click para editar"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "BUSCAR POR..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   5280
      Picture         =   "frm_conscliencu.frx":09EB
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1575
   End
End
Attribute VB_Name = "frm_conscliencu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
Data1.RecordSource = "Select * from clientes where cl_codigo =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 0)
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
    frm_encuestas.t_mat.Text = Data1.Recordset("cl_codigo")
    frm_encuestas.t_nom.Text = Data1.Recordset("cl_apellid")
End If
Unload Me

End Sub

Private Sub Form_Load()
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.ConnectionString = "dsn=" & Xconexrmt & ";"

DBGrid1.Rows = 2
DBGrid1.Cols = 7
DBGrid1.TextMatrix(0, 0) = "MATRICULA"
DBGrid1.ColWidth(0) = 1500
DBGrid1.TextMatrix(0, 1) = "NOMBRES"
DBGrid1.ColWidth(1) = 2900
DBGrid1.TextMatrix(0, 2) = "TELÉFONO"
DBGrid1.ColWidth(2) = 1900
DBGrid1.TextMatrix(0, 3) = "CEDULA"
DBGrid1.ColWidth(3) = 1200
DBGrid1.TextMatrix(0, 4) = "DG"
DBGrid1.ColWidth(4) = 200
DBGrid1.TextMatrix(0, 5) = "FEC.ING."
DBGrid1.ColWidth(5) = 1500
DBGrid1.TextMatrix(0, 6) = "CONVENIO"
DBGrid1.ColWidth(6) = 1500


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Text1.Text <> "" Then
      If Combo1.ListIndex = 0 Then
         Data1.RecordSource = "Select * from clientes where cl_apellid >='" & Text1.Text & "' order by cl_apellid limit 100"
         Data1.Refresh
      Else
         If Combo1.ListIndex = 1 Then
            Data1.RecordSource = "Select * from clientes where cl_cedula =" & Text1.Text & " order by cl_cedula"
            Data1.Refresh
         End If
      End If
   End If
    DBGrid1.Rows = 2
    DBGrid1.Cols = 7
    DBGrid1.TextMatrix(0, 0) = "MATRICULA"
    DBGrid1.ColWidth(0) = 1500
    DBGrid1.TextMatrix(0, 1) = "NOMBRES"
    DBGrid1.ColWidth(1) = 2900
    DBGrid1.TextMatrix(0, 2) = "TELÉFONO"
    DBGrid1.ColWidth(2) = 1900
    DBGrid1.TextMatrix(0, 3) = "CEDULA"
    DBGrid1.ColWidth(3) = 1200
    DBGrid1.TextMatrix(0, 4) = "DG"
    DBGrid1.ColWidth(4) = 200
    DBGrid1.TextMatrix(0, 5) = "FEC.ING."
    DBGrid1.ColWidth(5) = 1500
    DBGrid1.TextMatrix(0, 6) = "CONVENIO"
    DBGrid1.ColWidth(6) = 1500
   Dim Xcann As Integer
    Xcann = 1
    If Data1.Recordset.RecordCount > 0 Then
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
           If IsNull(Data1.Recordset("cl_codigo")) = False Then
              DBGrid1.TextMatrix(Xcann, 0) = Data1.Recordset("cl_codigo")
           End If
           If IsNull(Data1.Recordset("cl_apellid")) = False Then
              DBGrid1.TextMatrix(Xcann, 1) = Data1.Recordset("cl_apellid")
           End If
           If IsNull(Data1.Recordset("cl_telefon")) = False Then
              DBGrid1.TextMatrix(Xcann, 2) = Data1.Recordset("cl_telefon")
           End If
           If IsNull(Data1.Recordset("cl_cedula")) = False Then
              DBGrid1.TextMatrix(Xcann, 3) = Data1.Recordset("cl_cedula")
           End If
           If IsNull(Data1.Recordset("cl_codced")) = False Then
              DBGrid1.TextMatrix(Xcann, 4) = Data1.Recordset("cl_codced")
           End If
           If IsNull(Data1.Recordset("cl_fecing")) = False Then
              DBGrid1.TextMatrix(Xcann, 5) = Data1.Recordset("cl_fecing")
           End If
           If IsNull(Data1.Recordset("cl_codconv")) = False Then
              DBGrid1.TextMatrix(Xcann, 6) = Data1.Recordset("cl_codconv")
           End If
           DBGrid1.Rows = DBGrid1.Rows + 1
           Data1.Recordset.MoveNext
           Xcann = Xcann + 1
        Loop
    End If
   
   DBGrid1.SetFocus
End If

End Sub
