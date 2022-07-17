VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_veodeuda 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Ver deuda"
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_deudas 
      Height          =   375
      Left            =   3840
      Top             =   3840
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "data_deudas"
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7440
      Picture         =   "frm_veodeuda.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   3720
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3135
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   7
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label labno 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label labma 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1320
      Picture         =   "frm_veodeuda.frx":058A
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   2295
   End
End
Attribute VB_Name = "frm_veodeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Xmsel = 0
Xasel = 0
Unload Me

End Sub

Private Sub Form_Activate()
Dim xcuenta As Integer
Dim Xsaldoo As Long
Dim Xmesd, Xanod As Integer
xcuenta = 1
grid.ColWidth(0) = 1200
grid.ColWidth(0) = 1200
grid.ColWidth(0) = 1200
grid.ColWidth(0) = 1200
grid.ColWidth(4) = 2400
grid.ColWidth(5) = 1200
grid.ColWidth(6) = 1200

grid.TextMatrix(0, 0) = "FECHA"
grid.TextMatrix(0, 1) = "CUOTA"
grid.TextMatrix(0, 2) = "IMPORTE"
grid.TextMatrix(0, 3) = "SALDOS"
grid.TextMatrix(0, 4) = "DESCRIPCION"
grid.TextMatrix(0, 5) = "NUMERO"
grid.TextMatrix(0, 6) = "LINEA"

data_deudas.ConnectionString = "dsn=" & Xconexrmt
If XAlta = 599 Then
   labma.Caption = Xtot
   labno.Caption = frm_largador.txt_nomb.Text
Else
   labma.Caption = frmabm.txt_mat.Caption
   labno.Caption = frmabm.txt_apellid.Text
End If
If XAlta = 599 Then
   data_deudas.RecordSource = "Select * from deudas where cliente =" & Val(labma.Caption) & " and fecha_pago is null order by ano,mes"
Else
    If frm_factura.data_estudio.Recordset("codest") = 999 Then
       data_deudas.RecordSource = "Select * from deudas where cliente =" & Val(labma.Caption) & " and fecha_pago is null and mes not in (0) order by ano,mes"
    Else
       data_deudas.RecordSource = "Select * from deudas where cliente =" & Val(labma.Caption) & " and fecha_pago is null and mes in (0) order by ano,mes"
    End If
End If

data_deudas.Refresh
If data_deudas.Recordset.RecordCount > 0 Then
   Do While Not data_deudas.Recordset.EOF
      If IsNull(data_deudas.Recordset("fecha_pago")) = True Then
         Xsaldo = Xsaldo + data_deudas.Recordset("total")
         If IsNull(data_deudas.Recordset("fecha")) = False Then
            grid.TextMatrix(xcuenta, 0) = data_deudas.Recordset("fecha")
            Xmesd = Month(data_deudas.Recordset("fecha"))
            Xanod = Year(data_deudas.Recordset("fecha"))
         Else
            grid.TextMatrix(xcuenta, 0) = Date
            Xmesd = Month(Date)
            Xanod = Year(Date)
         End If
         If IsNull(data_deudas.Recordset("mes")) = False Then
            grid.TextMatrix(xcuenta, 1) = Trim(str(data_deudas.Recordset("mes"))) + "/" + Trim(str(data_deudas.Recordset("ano")))
         Else
            grid.TextMatrix(xcuenta, 1) = Trim(str(Xmesd)) + "/" + Trim(str(Xanod))
         End If
         If IsNull(data_deudas.Recordset("total")) = False Then
            grid.TextMatrix(xcuenta, 2) = data_deudas.Recordset("total")
         Else
            grid.TextMatrix(xcuenta, 2) = 0
         End If
         grid.TextMatrix(xcuenta, 3) = Format(Xsaldo, "Standard")
         If IsNull(data_deudas.Recordset("origen")) = False Then
            grid.TextMatrix(xcuenta, 4) = data_deudas.Recordset("origen")
         Else
            grid.TextMatrix(xcuenta, 4) = "DEUDA"
         End If
         If IsNull(data_deudas.Recordset("documento")) = False Then
            grid.TextMatrix(xcuenta, 5) = data_deudas.Recordset("documento")
         Else
            grid.TextMatrix(xcuenta, 5) = "0"
         End If
         If IsNull(data_deudas.Recordset("nro_vende")) = False Then
            grid.TextMatrix(xcuenta, 6) = data_deudas.Recordset("nro_vende")
         Else
            grid.TextMatrix(xcuenta, 6) = "0"
         End If
         xcuenta = xcuenta + 1
         grid.AddItem ""
      End If
      data_deudas.Recordset.MoveNext
   Loop
   If xcuenta <= 1 Then
      MsgBox "No hay deuda generada", vbInformation, "Facturación"
      Unload Me
   End If
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

Private Sub grid_KeyPress(KeyAscii As Integer)
If XAlta = 599 Then

Else
    If KeyAscii = 13 Then
       Xelnrodeuda = 0
       If Len(Trim(grid.TextMatrix(grid.Row, grid.Col))) = 7 Then
          frm_factura.txt_ano.Text = Mid(Trim((grid.TextMatrix(grid.Row, 1))), 4, 4)
          frm_factura.txt_mes.Text = Mid(Trim((grid.TextMatrix(grid.Row, 1))), 1, 2)
          frm_factura.txt_precio.Text = grid.TextMatrix(grid.Row, 2)
          Xelnrodeuda = Val(grid.TextMatrix(grid.Row, 5))
          Xelnrolind = Val(grid.TextMatrix(grid.Row, 6))
       Else
          frm_factura.txt_ano.Text = Mid(Trim((grid.TextMatrix(grid.Row, 1))), 3, 4)
          frm_factura.txt_mes.Text = Mid(Trim((grid.TextMatrix(grid.Row, 1))), 1, 1)
          frm_factura.txt_precio.Text = grid.TextMatrix(grid.Row, 2)
          Xelnrodeuda = Val(grid.TextMatrix(grid.Row, 5))
          Xelnrolind = Val(grid.TextMatrix(grid.Row, 6))
       End If
    '   frm_veodeuda.Hide
       Unload Me
       frm_factura.btn_graba.SetFocus
       frm_factura.txt_precio.Enabled = False
    End If
End If

End Sub

Private Sub listadeuda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'   frm_veodeuda.Hide
   Unload Me
End If

End Sub
