VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_consfarma 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas de medicación solicitada / entregada"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_consfarma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   10560
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid DBGrid1 
      Height          =   3255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   5741
      _Version        =   393216
      BackColorBkg    =   12615680
      SelectionMode   =   1
   End
   Begin MSAdodcLib.Adodc data_cons 
      Height          =   495
      Left            =   1920
      Top             =   4440
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "data_cons"
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   9720
      Picture         =   "frm_consfarma.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Buscar"
      Top             =   600
      Width           =   615
   End
   Begin MSMask.MaskEdBox mfh 
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mfd 
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9720
      Picture         =   "frm_consfarma.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salida"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "frm_consfarma.frx":0F56
      Left            =   2760
      List            =   "frm_consfarma.frx":0F63
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Left            =   2280
      TabIndex        =   8
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Total de registros:"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Consultar por:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   5280
      Picture         =   "frm_consfarma.frx":0F80
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "frm_consfarma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfd.SetFocus
End If

End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
frm_consfarma.MousePointer = 11
Command2.Enabled = False

If mfd.Text = "__/__/____" And mfh.Text = "__/__/____" Then
   If Text1.Text <> "" Then
      If Combo1.ListIndex = 0 Then
         data_cons.RecordSource = "Select * from linmmdd where nro_flia =" & 6 & " and ced_socio =" & Text1.Text & " order by fecha"
         data_cons.Refresh
         If data_cons.Recordset.RecordCount > 0 Then
            data_cons.Recordset.MoveLast
            Label3.Caption = data_cons.Recordset.RecordCount
         Else
            Label3.Caption = 0
         End If
      Else
         If Combo1.ListIndex = 1 Then
            data_cons.RecordSource = "Select * from linmmdd where nro_flia =" & 6 & " and cod_cli =" & Text1.Text & " order by fecha"
            data_cons.Refresh
            If data_cons.Recordset.RecordCount > 0 Then
               data_cons.Recordset.MoveLast
               Label3.Caption = data_cons.Recordset.RecordCount
            Else
               Label3.Caption = 0
            End If
         Else
            If Combo1.ListIndex = 2 Then
               data_cons.RecordSource = "Select * from linmmdd where nro_flia =" & 6 & " and base =" & Text1.Text & " order by fecha"
               data_cons.Refresh
               If data_cons.Recordset.RecordCount > 0 Then
                  data_cons.Recordset.MoveLast
                  Label3.Caption = data_cons.Recordset.RecordCount
               Else
                  Label3.Caption = 0
               End If
            End If
         End If
      End If
   End If
Else
   If Text1.Text <> "" Then
      If Combo1.ListIndex = 0 Then
         data_cons.RecordSource = "Select * from linmmdd where nro_flia =" & 6 & " and ced_socio =" & Text1.Text & " and fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by fecha"
         data_cons.Refresh
         If data_cons.Recordset.RecordCount > 0 Then
            data_cons.Recordset.MoveLast
            Label3.Caption = data_cons.Recordset.RecordCount
         Else
            Label3.Caption = 0
         End If
      Else
         If Combo1.ListIndex = 1 Then
            data_cons.RecordSource = "Select * from linmmdd where nro_flia =" & 6 & " and cod_cli =" & Text1.Text & " and fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by fecha"
            data_cons.Refresh
            If data_cons.Recordset.RecordCount > 0 Then
               data_cons.Recordset.MoveLast
               Label3.Caption = data_cons.Recordset.RecordCount
            Else
               Label3.Caption = 0
            End If
         Else
            If Combo1.ListIndex = 2 Then
               data_cons.RecordSource = "Select * from linmmdd where nro_flia =" & 6 & " and base =" & Text1.Text & " and fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by fecha"
               data_cons.Refresh
               If data_cons.Recordset.RecordCount > 0 Then
                  data_cons.Recordset.MoveLast
                  Label3.Caption = data_cons.Recordset.RecordCount
               Else
                  Label3.Caption = 0
               End If
            End If
         End If
      End If
   Else
      data_cons.RecordSource = "Select * from linmmdd where nro_flia =" & 6 & " and fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by fecha"
      data_cons.Refresh
      If data_cons.Recordset.RecordCount > 0 Then
         data_cons.Recordset.MoveLast
         Label3.Caption = data_cons.Recordset.RecordCount
      Else
         Label3.Caption = 0
      End If
   End If

End If
frm_consfarma.MousePointer = 0
Command2.Enabled = True

DBGrid1.Rows = 2
DBGrid1.Cols = 10
DBGrid1.TextMatrix(0, 0) = "FECHA"
DBGrid1.ColWidth(0) = 1200
DBGrid1.TextMatrix(0, 1) = "MATRICULA"
DBGrid1.ColWidth(1) = 1500
DBGrid1.TextMatrix(0, 2) = "NOMBRE"
DBGrid1.ColWidth(2) = 2900
DBGrid1.TextMatrix(0, 3) = "COD.PROD"
DBGrid1.ColWidth(3) = 1200
DBGrid1.TextMatrix(0, 4) = "DESCRIPCION"
DBGrid1.ColWidth(4) = 2900
DBGrid1.TextMatrix(0, 5) = "MEDICACIÓN"
DBGrid1.ColWidth(5) = 2900
DBGrid1.TextMatrix(0, 6) = "CONVENIO"
DBGrid1.ColWidth(6) = 1200
DBGrid1.TextMatrix(0, 7) = "CEDULA"
DBGrid1.ColWidth(7) = 1500
DBGrid1.TextMatrix(0, 8) = "BASE"
DBGrid1.ColWidth(8) = 500
DBGrid1.TextMatrix(0, 9) = "MEDICACIÓN ENT."
DBGrid1.ColWidth(9) = 2900

Dim Xcann As Integer
 Xcann = 1
 If data_cons.Recordset.RecordCount > 0 Then
     data_cons.Recordset.MoveFirst
     Do While Not data_cons.Recordset.EOF
        If IsNull(data_cons.Recordset("fecha")) = False Then
           DBGrid1.TextMatrix(Xcann, 0) = data_cons.Recordset("fecha")
        End If
        If IsNull(data_cons.Recordset("cod_cli")) = False Then
           DBGrid1.TextMatrix(Xcann, 1) = data_cons.Recordset("cod_cli")
        End If
        If IsNull(data_cons.Recordset("nom_cli")) = False Then
           DBGrid1.TextMatrix(Xcann, 2) = data_cons.Recordset("nom_cli")
        End If
        If IsNull(data_cons.Recordset("cod_prod")) = False Then
           DBGrid1.TextMatrix(Xcann, 3) = data_cons.Recordset("cod_prod")
        End If
        If IsNull(data_cons.Recordset("nom_prod")) = False Then
           DBGrid1.TextMatrix(Xcann, 4) = data_cons.Recordset("nom_prod")
        End If
        If IsNull(data_cons.Recordset("nom_medic")) = False Then
           DBGrid1.TextMatrix(Xcann, 5) = data_cons.Recordset("nom_medic")
        End If
        If IsNull(data_cons.Recordset("convenio")) = False Then
           DBGrid1.TextMatrix(Xcann, 6) = data_cons.Recordset("convenio")
        End If
        If IsNull(data_cons.Recordset("ced_socio")) = False Then
           DBGrid1.TextMatrix(Xcann, 7) = data_cons.Recordset("ced_socio")
        End If
        If IsNull(data_cons.Recordset("base")) = False Then
           DBGrid1.TextMatrix(Xcann, 8) = data_cons.Recordset("base")
        End If
        If IsNull(data_cons.Recordset("zona")) = False Then
           DBGrid1.TextMatrix(Xcann, 9) = data_cons.Recordset("zona")
        End If
        
        DBGrid1.Rows = DBGrid1.Rows + 1
        data_cons.Recordset.MoveNext
        Xcann = Xcann + 1
     Loop
 End If

         
End Sub

Private Sub Form_Load()
'data_cons.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cons.ConnectionString = "dsn=" & Xconexrmt
DBGrid1.Rows = 2
DBGrid1.Cols = 10
DBGrid1.TextMatrix(0, 0) = "FECHA"
DBGrid1.ColWidth(0) = 1200
DBGrid1.TextMatrix(0, 1) = "MATRICULA"
DBGrid1.ColWidth(1) = 1500
DBGrid1.TextMatrix(0, 2) = "NOMBRE"
DBGrid1.ColWidth(2) = 2900
DBGrid1.TextMatrix(0, 3) = "COD.PROD"
DBGrid1.ColWidth(3) = 1200
DBGrid1.TextMatrix(0, 4) = "DESCRIPCION"
DBGrid1.ColWidth(4) = 2900
DBGrid1.TextMatrix(0, 5) = "MEDICACIÓN"
DBGrid1.ColWidth(5) = 2900
DBGrid1.TextMatrix(0, 6) = "CONVENIO"
DBGrid1.ColWidth(6) = 1200
DBGrid1.TextMatrix(0, 7) = "CEDULA"
DBGrid1.ColWidth(7) = 1500
DBGrid1.TextMatrix(0, 8) = "BASE"
DBGrid1.ColWidth(8) = 500
DBGrid1.TextMatrix(0, 9) = "MEDICACIÓN ENT."
DBGrid1.ColWidth(9) = 2900

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

Private Sub mfh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Text1.SetFocus
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command2.SetFocus
End If

End Sub
