VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_busstock 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar datos..."
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8325
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_busstock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   8325
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data1 
      Height          =   495
      Left            =   2520
      Top             =   2280
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4683
      _Version        =   393216
      BackColorBkg    =   12615680
      SelectionMode   =   1
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7560
      Picture         =   "frm_busstock.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox t_bus 
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "frm_busstock.frx":09CC
      Left            =   2040
      List            =   "frm_busstock.frx":09D6
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Doble click selecciona el registro."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Buscar por..."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   5160
      Picture         =   "frm_busstock.frx":09EF
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1575
   End
End
Attribute VB_Name = "frm_busstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()

'frm_ctrolstock.data1.Recordset.FindFirst "id =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 0)

Data1.RecordSource = "Select * from stock where id =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 0)
Data1.Refresh

If Data1.Recordset.RecordCount > 0 Then
   frm_ctrolstock.t_cod.Text = Data1.Recordset("id")
   If IsNull(Data1.Recordset("descrip")) = False Then
      frm_ctrolstock.t_desc.Text = Data1.Recordset("descrip")
   Else
      frm_ctrolstock.t_desc.Text = ""
   End If
   If IsNull(Data1.Recordset("minimo")) = False Then
      frm_ctrolstock.t_min.Text = Data1.Recordset("minimo")
   Else
      frm_ctrolstock.t_min.Text = 0
   End If
   If IsNull(Data1.Recordset("basico")) = False Then
      frm_ctrolstock.t_bas.Text = Data1.Recordset("basico")
   Else
      frm_ctrolstock.t_bas.Text = 0
   End If
   If IsNull(Data1.Recordset("alerta")) = False Then
      frm_ctrolstock.t_alerta.Text = Data1.Recordset("alerta")
   Else
      frm_ctrolstock.t_alerta.Text = 0
   End If
   If IsNull(Data1.Recordset("actual")) = False Then
      frm_ctrolstock.t_act.Text = Data1.Recordset("actual")
   Else
      frm_ctrolstock.t_act.Text = 0
   End If
   If IsNull(Data1.Recordset("preuni")) = False Then
      frm_ctrolstock.t_prec.Text = Data1.Recordset("preuni")
   Else
      frm_ctrolstock.t_prec.Text = 0
   End If
   If IsNull(Data1.Recordset("actual")) = False Then
      frm_ctrolstock.t_act.Text = Data1.Recordset("actual")
   Else
      frm_ctrolstock.t_act.Text = 0
   End If
   If IsNull(Data1.Recordset("grupo")) = False Then
      frm_ctrolstock.Combo1.ListIndex = Data1.Recordset("grupo")
   Else
      frm_ctrolstock.Combo1.ListIndex = 0
   End If
   If IsNull(Data1.Recordset("ingreso")) = False Then
      frm_ctrolstock.mfing.Text = Data1.Recordset("ingreso")
   Else
      frm_ctrolstock.mfing.Text = "__/__/____"
   End If
   If IsNull(Data1.Recordset("ultact")) = False Then
      frm_ctrolstock.mfultact.Text = Data1.Recordset("ultact")
   Else
      frm_ctrolstock.mfultact.Text = "__/__/____"
   End If
   If IsNull(Data1.Recordset("vence")) = False Then
      frm_ctrolstock.mfvence.Text = Data1.Recordset("vence")
   Else
      frm_ctrolstock.mfvence.Text = "__/__/____"
   End If
   If IsNull(Data1.Recordset("obs")) = False Then
      frm_ctrolstock.t_obs.Text = Data1.Recordset("obs")
   Else
      frm_ctrolstock.t_obs.Text = ""
   End If
Else
   MsgBox "No se encuentra el código VERIFIQUE!!!", vbCritical, "Stock"
End If
Unload Me
   
End Sub

Private Sub Form_Load()
'Data1.DatabaseName = App.Path & "\" & Trim(Xlabdd)
'Data1.RecordSource = "stock"
'Data1.Refresh
Combo1.ListIndex = 0
DBGrid1.rows = 2
DBGrid1.Cols = 4
DBGrid1.TextMatrix(0, 0) = "CODIGO"
DBGrid1.ColWidth(0) = 1200
DBGrid1.TextMatrix(0, 1) = "DESCRIPCION"
DBGrid1.ColWidth(1) = 3500
DBGrid1.TextMatrix(0, 2) = "Stock Actual"
DBGrid1.ColWidth(2) = 1900
DBGrid1.TextMatrix(0, 3) = "Precio U."
DBGrid1.ColWidth(3) = 1900
Data1.ConnectionString = "dsn=" & Xconexrmt
If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
   Data1.RecordSource = "Select * from stock where grupo =" & 3 & " order by descrip"
Else
   Data1.RecordSource = "Select * from stock where grupo not in (3) order by descrip"
End If
Data1.Refresh
Dim Xcann As Integer
Xcann = 1
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      If IsNull(Data1.Recordset("id")) = False Then
         DBGrid1.TextMatrix(Xcann, 0) = Data1.Recordset("id")
      End If
      If IsNull(Data1.Recordset("descrip")) = False Then
         DBGrid1.TextMatrix(Xcann, 1) = Data1.Recordset("descrip")
      End If
      If IsNull(Data1.Recordset("actual")) = False Then
         DBGrid1.TextMatrix(Xcann, 2) = Data1.Recordset("actual")
      End If
      If IsNull(Data1.Recordset("preuni")) = False Then
         DBGrid1.TextMatrix(Xcann, 3) = Data1.Recordset("preuni")
      End If
      DBGrid1.rows = DBGrid1.rows + 1
      Data1.Recordset.MoveNext
      Xcann = Xcann + 1
   Loop
End If

End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Height = Me.Height
     .Width = Me.Width
End With

End Sub

Private Sub t_bus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Combo1.ListIndex = 0 Then
       If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
          Data1.RecordSource = "Select * from stock where descrip >='" & t_bus.Text & "' and grupo =" & 3 & " order by descrip"
       Else
          Data1.RecordSource = "Select * from stock where descrip >='" & t_bus.Text & "' and grupo not in (3) order by descrip"
       End If
       Data1.Refresh
    Else
       If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
          Data1.RecordSource = "Select * from stock where id >=" & t_bus.Text & " and grupo =" & 3 & " order by id"
       Else
          Data1.RecordSource = "Select * from stock where id >=" & t_bus.Text & " and grupo not in (3) order by id"
       End If
       Data1.Refresh
    End If
    DBGrid1.Clear
    DBGrid1.rows = 2
    DBGrid1.Cols = 4
    DBGrid1.TextMatrix(0, 0) = "CODIGO"
    DBGrid1.ColWidth(0) = 1200
    DBGrid1.TextMatrix(0, 1) = "DESCRIPCION"
    DBGrid1.ColWidth(1) = 3500
    DBGrid1.TextMatrix(0, 2) = "Stock Actual"
    DBGrid1.ColWidth(2) = 1900
    DBGrid1.TextMatrix(0, 3) = "Precio U."
    DBGrid1.ColWidth(3) = 1900
    Dim Xcann As Integer
    Xcann = 1
    If Data1.Recordset.RecordCount > 0 Then
       Data1.Recordset.MoveFirst
       Do While Not Data1.Recordset.EOF
          If IsNull(Data1.Recordset("id")) = False Then
             DBGrid1.TextMatrix(Xcann, 0) = Data1.Recordset("id")
          End If
          If IsNull(Data1.Recordset("descrip")) = False Then
             DBGrid1.TextMatrix(Xcann, 1) = Data1.Recordset("descrip")
          End If
          If IsNull(Data1.Recordset("actual")) = False Then
             DBGrid1.TextMatrix(Xcann, 2) = Data1.Recordset("actual")
          End If
          If IsNull(Data1.Recordset("preuni")) = False Then
             DBGrid1.TextMatrix(Xcann, 3) = Data1.Recordset("preuni")
          End If
          DBGrid1.rows = DBGrid1.rows + 1
          Data1.Recordset.MoveNext
          Xcann = Xcann + 1
       Loop
    End If
    DBGrid1.SetFocus
    
End If

End Sub
