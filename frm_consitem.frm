VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_consitem 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de items..."
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8025
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_consitem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   8025
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid DBGrid1 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4260
      _Version        =   393216
      BackColorBkg    =   12615680
      SelectionMode   =   1
   End
   Begin MSAdodcLib.Adodc data_it 
      Height          =   495
      Left            =   4440
      Top             =   3120
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "data_it"
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      Picture         =   "frm_consitem.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox t_cons 
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Doble click para seleccionar."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Consultar por DESCRIPCION:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   2640
      Picture         =   "frm_consitem.frx":09CC
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   975
   End
End
Attribute VB_Name = "frm_consitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
'DBGrid1.TextMatrix(DBGrid1.RowSel, 0)

data_it.RecordSource = "Select * from stock where id =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 0)
data_it.Refresh
If data_it.Recordset("actual") <= 0 Then
   MsgBox "Sin Stock, VERIFIQUE!", vbCritical
End If
frm_reggasto.labdesc.Caption = data_it.Recordset("descrip")
frm_reggasto.t_cod.Text = data_it.Recordset("id")

Unload Me


End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   DBGrid1_DblClick
End If

End Sub

Private Sub Form_Load()
'data_it.Connect = "ODBC;DSN=stock;"
data_it.ConnectionString = "dsn=" & Xconexrmt
DBGrid1.rows = 2
DBGrid1.Cols = 4
DBGrid1.TextMatrix(0, 0) = "CODIGO"
DBGrid1.ColWidth(0) = 1200
DBGrid1.TextMatrix(0, 1) = "DESCRIPCIÓN"
DBGrid1.ColWidth(1) = 3200
DBGrid1.TextMatrix(0, 2) = "STOCK"
DBGrid1.ColWidth(2) = 1200
DBGrid1.TextMatrix(0, 3) = "PRECIO"
DBGrid1.ColWidth(3) = 1200


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub t_cons_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
      data_it.RecordSource = "Select * from stock where descrip >='" & t_cons.Text & "' and grupo =" & 3 & " order by descrip"
   Else
      data_it.RecordSource = "Select * from stock where descrip >='" & t_cons.Text & "' and grupo not in (3) order by descrip"
   End If
   data_it.Refresh
   
   DBGrid1.rows = 2
   DBGrid1.Cols = 4
   DBGrid1.TextMatrix(0, 0) = "CODIGO"
   DBGrid1.ColWidth(0) = 1200
   DBGrid1.TextMatrix(0, 1) = "DESCRIPCIÓN"
   DBGrid1.ColWidth(1) = 3200
   DBGrid1.TextMatrix(0, 2) = "STOCK"
   DBGrid1.ColWidth(2) = 1200
   DBGrid1.TextMatrix(0, 3) = "PRECIO"
   DBGrid1.ColWidth(3) = 1200
    Dim Xcann As Integer
     Xcann = 1
     If data_it.Recordset.RecordCount > 0 Then
         data_it.Recordset.MoveFirst
         Do While Not data_it.Recordset.EOF
            If IsNull(data_it.Recordset("id")) = False Then
               DBGrid1.TextMatrix(Xcann, 0) = data_it.Recordset("id")
            End If
            If IsNull(data_it.Recordset("descrip")) = False Then
               DBGrid1.TextMatrix(Xcann, 1) = data_it.Recordset("descrip")
            End If
            If IsNull(data_it.Recordset("actual")) = False Then
               DBGrid1.TextMatrix(Xcann, 2) = data_it.Recordset("actual")
            End If
            If IsNull(data_it.Recordset("preuni")) = False Then
               DBGrid1.TextMatrix(Xcann, 3) = data_it.Recordset("preuni")
            End If
            
            DBGrid1.rows = DBGrid1.rows + 1
            data_it.Recordset.MoveNext
            Xcann = Xcann + 1
         Loop
     End If
   
   DBGrid1.SetFocus
End If

End Sub
