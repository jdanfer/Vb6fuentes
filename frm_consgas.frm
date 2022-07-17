VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_consgas 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar gastos registrados"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_consgas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   9600
   StartUpPosition =   1  'CenterOwner
   Begin MSMask.MaskEdBox md 
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid DBGrid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4683
      _Version        =   393216
      BackColorBkg    =   12615680
      SelectionMode   =   1
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   495
      Left            =   1680
      Top             =   3600
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      Picture         =   "frm_consgas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox t_bus 
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "frm_consgas.frx":09CC
      Left            =   2040
      List            =   "frm_consgas.frx":09DC
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha desde:"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Doble click selecciona el registro."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Buscar por..."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   6600
      Picture         =   "frm_consgas.frx":0A0D
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1575
   End
End
Attribute VB_Name = "frm_consgas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
Data1.RecordSource = "Select * from gastos where id =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 6)
Data1.Refresh

If IsNull(Data1.Recordset("id")) = False Then
   frm_reggasto.Text1.Text = Data1.Recordset("id")
Else
   frm_reggasto.Text1.Text = 0
End If
If IsNull(Data1.Recordset("prec")) = False Then
   frm_reggasto.labnrop.Caption = Data1.Recordset("prec")
Else
   frm_reggasto.labnrop.Caption = 0
End If
If IsNull(Data1.Recordset("codcli")) = False Then
   frm_reggasto.t_cli.Text = Data1.Recordset("codcli")
Else
   frm_reggasto.t_cli.Text = 0
End If
If IsNull(Data1.Recordset("codprod")) = False Then
   frm_reggasto.t_cod.Text = Data1.Recordset("codprod")
Else
   frm_reggasto.t_cod.Text = 0
End If
If IsNull(Data1.Recordset("nomcli")) = False Then
   frm_reggasto.labcli.Caption = Data1.Recordset("nomcli")
Else
   frm_reggasto.labcli.Caption = ""
End If
If IsNull(Data1.Recordset("descrip")) = False Then
   frm_reggasto.labdesc.Caption = Data1.Recordset("descrip")
Else
   frm_reggasto.labdesc.Caption = ""
End If
If IsNull(Data1.Recordset("cant")) = False Then
   frm_reggasto.t_cant.Text = Data1.Recordset("cant")
   frm_reggasto.t_cantant.Text = Data1.Recordset("cant")
Else
   frm_reggasto.t_cant.Text = 0
   frm_reggasto.t_cantant.Text = 0
End If
If IsNull(Data1.Recordset("obs")) = False Then
   frm_reggasto.t_obs.Text = Data1.Recordset("obs")
Else
   frm_reggasto.t_obs.Text = ""
End If
If IsNull(Data1.Recordset("fecha")) = False Then
   frm_reggasto.mfec.Text = Data1.Recordset("fecha")
Else
   frm_reggasto.mfec.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("hora")) = False Then
   frm_reggasto.mhor.Text = Format(Data1.Recordset("hora"), "HH:mm")
Else
   frm_reggasto.mhor.Text = "__:__"
End If
If IsNull(Data1.Recordset("usuario")) = False Then
   frm_reggasto.labus.Caption = Data1.Recordset("usuario")
Else
   frm_reggasto.labus.Caption = ""
End If
frm_reggasto.data_verent.RecordSource = "Select * from gastos where id =" & Data1.Recordset("id")
frm_reggasto.data_verent.Refresh
'frm_reggasto.data_gasto.Recordset.FindFirst "id =" & data1.Recordset("id")

Unload Me

End Sub

Private Sub Form_Load()
'Data1.DatabaseName = App.Path & "\" & Trim(Xlabdd)
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data1.RecordSource = "Select * from gastos order by fecha"
'data1.Refresh
Data1.ConnectionString = "dsn=" & Xconexrmt
DBGrid1.rows = 2
DBGrid1.Cols = 7
DBGrid1.TextMatrix(0, 0) = "FECHA"
DBGrid1.ColWidth(0) = 1200
DBGrid1.TextMatrix(0, 1) = "CODIGO"
DBGrid1.ColWidth(1) = 1200
DBGrid1.TextMatrix(0, 2) = "DESCRIPCIÓN"
DBGrid1.ColWidth(2) = 3200
DBGrid1.TextMatrix(0, 3) = "CLIENTE"
DBGrid1.ColWidth(3) = 1200
DBGrid1.TextMatrix(0, 4) = "CANT."
DBGrid1.ColWidth(4) = 1200
DBGrid1.TextMatrix(0, 5) = "OBSERVACIÓN"
DBGrid1.ColWidth(5) = 3900
DBGrid1.TextMatrix(0, 6) = "ID"
DBGrid1.ColWidth(6) = 400

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If t_bus.Text = "" Then
   Else
      If Combo1.ListIndex = 0 Then
         If md.Text <> "__/__/____" Then
            Data1.RecordSource = "Select * from gastos where descrip >='" & t_bus.Text & "' and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' order by descrip,fecha"
         Else
            Data1.RecordSource = "Select * from gastos where descrip >='" & t_bus.Text & "' order by descrip limit 1000"
         End If
         Data1.Refresh
      Else
         If Combo1.ListIndex = 2 Then
            If md.Text <> "__/__/____" Then
               Data1.RecordSource = "Select * from gastos where codcli =" & t_bus.Text & " and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' order by fecha"
            Else
               Data1.RecordSource = "Select * from gastos where codcli =" & t_bus.Text & " order by fecha"
            End If
            Data1.Refresh
         End If
      End If
   End If
   DBGrid1.rows = 2
   DBGrid1.Cols = 7
   DBGrid1.TextMatrix(0, 0) = "FECHA"
   DBGrid1.ColWidth(0) = 1200
   DBGrid1.TextMatrix(0, 1) = "CODIGO"
   DBGrid1.ColWidth(1) = 1200
   DBGrid1.TextMatrix(0, 2) = "DESCRIPCIÓN"
   DBGrid1.ColWidth(2) = 3200
   DBGrid1.TextMatrix(0, 3) = "CLIENTE"
   DBGrid1.ColWidth(3) = 1200
   DBGrid1.TextMatrix(0, 4) = "CANT."
   DBGrid1.ColWidth(4) = 1200
   DBGrid1.TextMatrix(0, 5) = "OBSERVACIÓN"
   DBGrid1.ColWidth(5) = 3900
   DBGrid1.TextMatrix(0, 6) = "ID"
   DBGrid1.ColWidth(6) = 400
    
    Dim Xcann As Integer
     Xcann = 1
     If Data1.Recordset.RecordCount > 0 Then
         Data1.Recordset.MoveFirst
         Do While Not Data1.Recordset.EOF
            If IsNull(Data1.Recordset("fecha")) = False Then
               DBGrid1.TextMatrix(Xcann, 0) = Data1.Recordset("fecha")
            End If
            If IsNull(Data1.Recordset("codprod")) = False Then
               DBGrid1.TextMatrix(Xcann, 1) = Data1.Recordset("codprod")
            End If
            If IsNull(Data1.Recordset("descrip")) = False Then
               DBGrid1.TextMatrix(Xcann, 2) = Data1.Recordset("descrip")
            End If
            If IsNull(Data1.Recordset("codcli")) = False Then
               DBGrid1.TextMatrix(Xcann, 3) = Data1.Recordset("codcli")
            End If
            If IsNull(Data1.Recordset("cant")) = False Then
               DBGrid1.TextMatrix(Xcann, 4) = Data1.Recordset("cant")
            End If
            If IsNull(Data1.Recordset("obs")) = False Then
               DBGrid1.TextMatrix(Xcann, 5) = Data1.Recordset("obs")
            End If
            If IsNull(Data1.Recordset("id")) = False Then
               DBGrid1.TextMatrix(Xcann, 6) = Data1.Recordset("id")
            End If
            
            DBGrid1.rows = DBGrid1.rows + 1
            Data1.Recordset.MoveNext
            Xcann = Xcann + 1
         Loop
     End If
   
   DBGrid1.SetFocus
End If

End Sub

Private Sub t_bus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If t_bus.Text = "" Then
   Else
      If Combo1.ListIndex = 1 Or Combo1.ListIndex = 3 Then
         If Combo1.ListIndex = 3 Then
            Data1.RecordSource = "Select * from gastos where codprod =" & t_bus.Text & " order by fecha"
         Else
            Data1.RecordSource = "Select * from gastos where fecha >='" & Format(t_bus.Text, "yyyy-mm-dd") & "' order by fecha"
         End If
         Data1.Refresh
        DBGrid1.rows = 2
        DBGrid1.Cols = 7
        DBGrid1.TextMatrix(0, 0) = "FECHA"
        DBGrid1.ColWidth(0) = 1200
        DBGrid1.TextMatrix(0, 1) = "CODIGO"
        DBGrid1.ColWidth(1) = 1200
        DBGrid1.TextMatrix(0, 2) = "DESCRIPCIÓN"
        DBGrid1.ColWidth(2) = 3200
        DBGrid1.TextMatrix(0, 3) = "CLIENTE"
        DBGrid1.ColWidth(3) = 1200
        DBGrid1.TextMatrix(0, 4) = "CANT."
        DBGrid1.ColWidth(4) = 1200
        DBGrid1.TextMatrix(0, 5) = "OBSERVACIÓN"
        DBGrid1.ColWidth(5) = 3900
        DBGrid1.TextMatrix(0, 6) = "ID"
        DBGrid1.ColWidth(6) = 400
         
         Dim Xcann As Integer
          Xcann = 1
          If Data1.Recordset.RecordCount > 0 Then
              Data1.Recordset.MoveFirst
              Do While Not Data1.Recordset.EOF
                 If IsNull(Data1.Recordset("fecha")) = False Then
                    DBGrid1.TextMatrix(Xcann, 0) = Data1.Recordset("fecha")
                 End If
                 If IsNull(Data1.Recordset("codprod")) = False Then
                    DBGrid1.TextMatrix(Xcann, 1) = Data1.Recordset("codprod")
                 End If
                 If IsNull(Data1.Recordset("descrip")) = False Then
                    DBGrid1.TextMatrix(Xcann, 2) = Data1.Recordset("descrip")
                 End If
                 If IsNull(Data1.Recordset("codcli")) = False Then
                    DBGrid1.TextMatrix(Xcann, 3) = Data1.Recordset("codcli")
                 End If
                 If IsNull(Data1.Recordset("cant")) = False Then
                    DBGrid1.TextMatrix(Xcann, 4) = Data1.Recordset("cant")
                 End If
                 If IsNull(Data1.Recordset("obs")) = False Then
                    DBGrid1.TextMatrix(Xcann, 5) = Data1.Recordset("obs")
                 End If
                 If IsNull(Data1.Recordset("id")) = False Then
                    DBGrid1.TextMatrix(Xcann, 6) = Data1.Recordset("id")
                 End If
                 
                 DBGrid1.rows = DBGrid1.rows + 1
                 Data1.Recordset.MoveNext
                 Xcann = Xcann + 1
              Loop
          End If
        
        DBGrid1.SetFocus
      
      
      Else
         md.SetFocus
      End If
   End If
End If

End Sub
