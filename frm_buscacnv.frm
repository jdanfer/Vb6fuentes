VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_buscacnv 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Consulta"
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2160
      Top             =   3480
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Por Razón Social"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   360
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   495
      Left            =   8040
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSFlexGridLib.MSFlexGrid DBgrid1 
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6588
      _Version        =   393216
      BackColorBkg    =   12615680
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9840
      Picture         =   "frm_buscacnv.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cerrar"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txt_busca 
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
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Por código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Buscar por descripción"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   8055
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   960
      Picture         =   "frm_buscacnv.frx":058A
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1095
   End
End
Attribute VB_Name = "frm_buscacnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
Data1.RecordSource = "Select * from convenio where cnv_codigo ='" & DBgrid1.TextMatrix(DBgrid1.RowSel, 0) & "'"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
    If frm_inffaccnv.Visible = True Then
       frm_inffaccnv.t_cli.Text = Data1.Recordset("cnv_cuenta")
       ControlproxEmi
       Unload Me
    Else
       frmabm.txt_codcnv.Text = Data1.Recordset("cnv_codigo")
       frmabm.txt_nomcnv.Text = Data1.Recordset("cnv_desc")
       If IsNull(Data1.Recordset("cnv_entre")) = False Then
          If Trim(Data1.Recordset("cnv_entre")) <> "" Then
             frmabm.t_rs.Text = Data1.Recordset("cnv_entre")
          Else
             frmabm.t_rs.Text = ""
          End If
       Else
          frmabm.t_rs.Text = ""
       End If
       ControlproxEmi
       Unload Me
    '   frm_buscacnv.Hide
    End If
End If

End Sub

Private Sub DBgrid1_EnterCell()
'MsgBox "ES:" & DBgrid1.TextMatrix(DBgrid1.RowSel, 0) & "'"

If DBgrid1.TextMatrix(DBgrid1.RowSel, 0) <> "" Then
   Adodc1.RecordSource = "Select * from convenio where cnv_codigo ='" & DBgrid1.TextMatrix(DBgrid1.RowSel, 0) & "'"
   Adodc1.Refresh
   If Adodc1.Recordset.RecordCount > 0 Then
      Adodc1.Recordset.MoveFirst
      If IsNull(Adodc1.Recordset("cnv_motbaj")) = False Then
         Label1.Caption = Adodc1.Recordset("cnv_motbaj")
      Else
         Label1.Caption = ""
      End If
   End If
End If

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   DBGrid1_DblClick
End If

End Sub

Private Sub Form_Activate()
'txt_busca.SetFocus
Option1.SetFocus

End Sub

Private Sub Form_Deactivate()
frmabm.txt_codcnv.SetFocus
'frm_buscacnv.Hide
End Sub

Private Sub Form_Load()
Dim Xss As String
Xss = "SI"
Adodc1.ConnectionString = "dsn=" & Xconexrmt

'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data1.RecordSource = "convenio"
'data1.Refresh
Data1.ConnectionString = "dsn=" & Xconexrmt
If Xconv = "" Then
   Xconv = "PART"
End If

If XWeltipoU = "ADMINISTRADOR" Then
   Data1.RecordSource = "Select * from convenio where cnv_codigo >= '" & UCase(Xconv) & "' and cnv_fbaja is null and cnv_umpago not in (1) order by cnv_codigo"
   Data1.Refresh
Else
   Data1.RecordSource = "Select * from convenio where cnv_codigo >= '" & UCase(Xconv) & "' And Cnv_alta ='" & Xss & "' and cnv_fbaja is null and cnv_umpago not in (1) order by cnv_codigo"
   Data1.Refresh
End If
DBgrid1.rows = 2
DBgrid1.Cols = 3
DBgrid1.TextMatrix(0, 0) = "CODIGO"
DBgrid1.ColWidth(0) = 1500
DBgrid1.TextMatrix(0, 1) = "DESCRIPCION"
DBgrid1.ColWidth(1) = 4900
DBgrid1.TextMatrix(0, 2) = "RAZON SOCIAL"
DBgrid1.ColWidth(2) = 4900

Dim Xcann As Integer
Xcann = 1
If Data1.Recordset.RecordCount > 0 Then
    Data1.Recordset.MoveFirst
    Do While Not Data1.Recordset.EOF
       If IsNull(Data1.Recordset("cnv_codigo")) = False Then
          DBgrid1.TextMatrix(Xcann, 0) = Data1.Recordset("cnv_codigo")
       End If
       If IsNull(Data1.Recordset("cnv_desc")) = False Then
          DBgrid1.TextMatrix(Xcann, 1) = Data1.Recordset("cnv_desc")
       End If
       If IsNull(Data1.Recordset("cnv_entre")) = False Then
          DBgrid1.TextMatrix(Xcann, 2) = Data1.Recordset("cnv_entre")
       End If
       
       DBgrid1.rows = DBgrid1.rows + 1
       Data1.Recordset.MoveNext
       Xcann = Xcann + 1
    Loop
End If

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_busca.SetFocus
End If

End Sub

Private Sub Option2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_busca.SetFocus
End If

End Sub

Private Sub TXT_BUSCA_KeyPress(KeyAscii As Integer)
Dim Xsdos As String
Xsdos = "SI"
If KeyAscii = 13 Then
    Data1.Refresh
    DBgrid1.rows = 2
    DBgrid1.Cols = 3
    DBgrid1.TextMatrix(0, 0) = "CODIGO"
    DBgrid1.ColWidth(0) = 1500
    DBgrid1.TextMatrix(0, 1) = "DESCRIPCION"
    DBgrid1.ColWidth(1) = 4900
    DBgrid1.TextMatrix(0, 2) = "RAZON SOCIAL"
    DBgrid1.ColWidth(2) = 4900
    Dim Xcann As Integer
    Xcann = 1
    If Data1.Recordset.RecordCount > 0 Then
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
           If IsNull(Data1.Recordset("cnv_codigo")) = False Then
              DBgrid1.TextMatrix(Xcann, 0) = Data1.Recordset("cnv_codigo")
           End If
           If IsNull(Data1.Recordset("cnv_desc")) = False Then
              DBgrid1.TextMatrix(Xcann, 1) = Data1.Recordset("cnv_desc")
           End If
           If IsNull(Data1.Recordset("cnv_entre")) = False Then
              DBgrid1.TextMatrix(Xcann, 2) = Data1.Recordset("cnv_entre")
           End If
           DBgrid1.rows = DBgrid1.rows + 1
           Data1.Recordset.MoveNext
           Xcann = Xcann + 1
        Loop
    End If
    DBgrid1.SetFocus
Else
   KeyAscii = Asc(UCase(chr(KeyAscii)))
   If Option1.Value = True Then
      If frm_usuario.data_usuario.Recordset("tipo") = "ADMINISTRADOR" Then
         Data1.RecordSource = "select * from convenio Where cnv_desc >= '" & UCase(txt_busca.Text) & "' and cnv_fbaja is null and cnv_umpago not in (1) Order by cnv_desc limit 150"
      Else
         Data1.RecordSource = "select * from convenio Where cnv_desc >= '" & UCase(txt_busca.Text) & "' And cnv_alta ='" & Xsdos & "' and cnv_fbaja is null and cnv_umpago not in (1) Order by cnv_desc limit 150"
      End If
   Else
      If Option2.Value = True Then
         If frm_usuario.data_usuario.Recordset("tipo") = "ADMINISTRADOR" Then
            Data1.RecordSource = "select * from convenio Where cnv_codigo >= '" & UCase(txt_busca.Text) & "' and cnv_fbaja is null and cnv_umpago not in (1) Order by cnv_codigo limit 70"
         Else
            Data1.RecordSource = "select * from convenio Where cnv_codigo >= '" & UCase(txt_busca.Text) & "' And cnv_alta ='" & Xsdos & "' and cnv_fbaja is null and cnv_umpago not in (1) Order by cnv_codigo limit 70"
         End If
      Else
         If frm_usuario.data_usuario.Recordset("tipo") = "ADMINISTRADOR" Then
            Data1.RecordSource = "select * from convenio Where cnv_entre >= '" & UCase(txt_busca.Text) & "' and cnv_fbaja is null and cnv_umpago not in (1) Order by cnv_entre limit 70"
         Else
            Data1.RecordSource = "select * from convenio Where cnv_entre >= '" & UCase(txt_busca.Text) & "' And cnv_alta ='" & Xsdos & "' and cnv_fbaja is null and cnv_umpago not in (1) Order by cnv_entre limit 70"
         End If
      End If
   End If
End If

End Sub

Public Sub ControlproxEmi()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xmmpemi, Xaapemi As Integer
Dim Xarmoelmesd, Xarmoelmesh As String

If Month(Date) > 9 Then
  Xarmoelmesh = "01/" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
Else
  Xarmoelmesh = "01/0" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
End If

ConectarBD
ConbdSapp.Open
             
If Month(Date) = 12 Then
   Xmmpemi = 1
   Xaapemi = Year(Date) + 1
Else
   Xmmpemi = Month(Date) + 1
   Xaapemi = Year(Date)
End If

Xsqlpromo = "Select * from convenio where cnv_codigo ='" & DBgrid1.TextMatrix(DBgrid1.RowSel, 0) & "' and cnv_emite ='" & "SI" & "' and cnv_cant_r in (1,2) and cnv_hasta >='" & Format(Date, "yyyy/mm/dd") & "'"

With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   If Day(Date) > 25 Then
        If frmabm.t_pmemi.Text <> "" Then
           If Val(frmabm.t_pmemi.Text) <= 0 Then
              If Xmmpemi = 12 Then
                 frmabm.t_pmemi.Text = 1
                 frmabm.t_paemi.Text = Xaapemi + 1
              Else
                 frmabm.t_pmemi.Text = Xmmpemi + 1
                 frmabm.t_paemi.Text = Xaapemi
              End If
           Else
              If frmabm.t_pmemi.Text > 9 Then
                 Xarmoelmesd = "01/" & Trim(str(frmabm.t_pmemi.Text)) & "/" & Trim(str(frmabm.t_paemi.Text))
              Else
                 Xarmoelmesd = "01/0" & Trim(str(frmabm.t_pmemi.Text)) & "/" & Trim(str(frmabm.t_paemi.Text))
              End If
              If Format(Xarmoelmesd, "yyyy/mm/dd") <= Format(Xarmoelmesh, "yyyy/mm/dd") Then
                 If Xmmpemi = 12 Then
                    frmabm.t_pmemi.Text = 1
                    frmabm.t_paemi.Text = Xaapemi + 1
                 Else
                    frmabm.t_pmemi.Text = Xmmpemi + 1
                    frmabm.t_paemi.Text = Xaapemi
                 End If
              End If
           End If
        Else
           If Xmmpemi = 12 Then
              frmabm.t_pmemi.Text = 1
              frmabm.t_paemi.Text = Xaapemi + 1
           Else
              frmabm.t_pmemi.Text = Xmmpemi + 1
              frmabm.t_paemi.Text = Xaapemi
           End If
        End If
   Else
        If frmabm.t_pmemi.Text <> "" Then
           If Val(frmabm.t_pmemi.Text) <= 0 Then
              frmabm.t_pmemi.Text = Xmmpemi
              frmabm.t_paemi.Text = Xaapemi
           Else
              If frmabm.t_pmemi.Text > 9 Then
                 Xarmoelmesd = "01/" & Trim(str(frmabm.t_pmemi.Text)) & "/" & Trim(str(frmabm.t_paemi.Text))
              Else
                 Xarmoelmesd = "01/0" & Trim(str(frmabm.t_pmemi.Text)) & "/" & Trim(str(frmabm.t_paemi.Text))
              End If
              If Format(Xarmoelmesd, "yyyy/mm/dd") <= Format(Xarmoelmesh, "yyyy/mm/dd") Then
                 frmabm.t_pmemi.Text = Xmmpemi
                 frmabm.t_paemi.Text = Xaapemi
              End If
           End If
        Else
           frmabm.t_pmemi.Text = Xmmpemi
           frmabm.t_paemi.Text = Xaapemi
        End If
   End If
Else
   frmabm.t_pmemi.Text = 0
   frmabm.t_paemi.Text = 0
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

