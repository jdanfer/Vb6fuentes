VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_busllamado 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar llamados"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9990
   Icon            =   "frm_busllamado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9990
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc data_chof 
      Height          =   495
      Left            =   1200
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data_chof"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   4080
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   375
      Left            =   6360
      Top             =   600
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
      DataSourceName  =   "sappnew"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data4"
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   375
      Left            =   6720
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "data3"
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
      Left            =   120
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Height          =   5415
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9551
      _Version        =   393216
      BackColorBkg    =   12615680
      FocusRect       =   2
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
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Historial"
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
      Left            =   4920
      TabIndex        =   8
      Top             =   720
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Actual"
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
      Left            =   2400
      TabIndex        =   7
      Top             =   720
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.CommandButton bok 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Picture         =   "frm_busllamado.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Procesar búsqueda"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSMask.MaskEdBox mh 
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
   Begin MSMask.MaskEdBox md 
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0E0FF&
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
      ItemData        =   "frm_busllamado.frx":09CC
      Left            =   2400
      List            =   "frm_busllamado.frx":09D9
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Buscar en:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "BUSCAR POR..."
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
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   120
      Picture         =   "frm_busllamado.frx":09F4
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frm_busllamado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bok_Click()
On Error GoTo Quebusca
bok.Enabled = False

If Combo1.ListIndex = 0 Then
   If md.Text <> "__/__/____" Then
      If mh.Text <> "__/__/____" Then
         If Option2.Value = True Then
'            data1.DatabaseName = App.Path & "\llamado.mdb"
            Data1.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.path & "\llamado.mdb"
            Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha<=#" & Format(mh.Text, "yyyy/mm/dd") & "# order by fecha,hora"
            Data1.Refresh
         Else
            Data1.ConnectionString = "DSN=" & Xconexrmt
            Data1.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' And fecha<='" & Format(mh.Text, "yyyy/mm/dd") & "' order by fecha,hora"
            Data1.Refresh
         End If
      End If
   End If
Else
   If Combo1.ListIndex = 1 Then
      If Text1.Text <> "" Then
         If Option2.Value = True Then
            Data1.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.path & "\llamado.mdb"
            Data1.RecordSource = "Select * from llamado where nombre >='" & Text1.Text & "' order by nombre,fecha,hora limit 300"
            Data1.Refresh
         Else
            Data1.ConnectionString = "DSN=" & Xconexrmt
            Data1.RecordSource = "Select * from llamado where nombre >='" & Text1.Text & "' order by nombre,fecha,hora limit 300"
            Data1.Refresh
         End If
      End If
   Else
      If Combo1.ListIndex = 2 Then
         If Text1.Text <> "" Then
            If Option2.Value = True Then
               Data1.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.path & "\llamado.mdb"
               Data1.RecordSource = "Select * from llamado where ci =" & Text1.Text & " order by fecha,hora"
               Data1.Refresh
            Else
               Data1.ConnectionString = "DSN=" & Xconexrmt
               Data1.RecordSource = "Select * from llamado where ci =" & Text1.Text & " order by fecha DESC,hora"
               Data1.Refresh
            End If
         End If
      End If
   End If
End If
DBGrid1.Clear
DBGrid1.rows = 2
DBGrid1.Cols = 14
DBGrid1.TextMatrix(0, 0) = "FECHA"
DBGrid1.ColWidth(0) = 1200
DBGrid1.TextMatrix(0, 1) = "HORA"
DBGrid1.ColWidth(1) = 900
DBGrid1.TextMatrix(0, 2) = "NOMBRE"
DBGrid1.ColWidth(2) = 2900
DBGrid1.TextMatrix(0, 3) = "MATRICULA"
DBGrid1.ColWidth(3) = 1500
DBGrid1.TextMatrix(0, 4) = "CATEG."
DBGrid1.ColWidth(4) = 1200
DBGrid1.TextMatrix(0, 5) = "EDAD"
DBGrid1.ColWidth(5) = 500
DBGrid1.TextMatrix(0, 6) = "DIRECCION"
DBGrid1.ColWidth(6) = 2900
DBGrid1.TextMatrix(0, 7) = "TELEFONO"
DBGrid1.ColWidth(7) = 1500
DBGrid1.TextMatrix(0, 8) = "MOT.CONSULTA"
DBGrid1.ColWidth(8) = 3900
DBGrid1.TextMatrix(0, 9) = "CODIGO"
DBGrid1.ColWidth(9) = 400
DBGrid1.TextMatrix(0, 10) = "MOVIL"
DBGrid1.ColWidth(10) = 400
DBGrid1.TextMatrix(0, 11) = "FEC.PAS"
DBGrid1.ColWidth(11) = 1200
DBGrid1.TextMatrix(0, 12) = "HORA PAS"
DBGrid1.ColWidth(12) = 500
DBGrid1.TextMatrix(0, 13) = "ID"
DBGrid1.ColWidth(13) = 1200

Dim Xcann As Integer
Xcann = 1
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      If IsNull(Data1.Recordset("fecha")) = False Then
         DBGrid1.TextMatrix(Xcann, 0) = Data1.Recordset("fecha")
      End If
      If IsNull(Data1.Recordset("hora")) = False Then
         DBGrid1.TextMatrix(Xcann, 1) = Data1.Recordset("hora")
      End If
      If IsNull(Data1.Recordset("nombre")) = False Then
         DBGrid1.TextMatrix(Xcann, 2) = Data1.Recordset("nombre")
      End If
      If IsNull(Data1.Recordset("matric")) = False Then
         DBGrid1.TextMatrix(Xcann, 3) = Data1.Recordset("matric")
      End If
      If IsNull(Data1.Recordset("categ")) = False Then
         DBGrid1.TextMatrix(Xcann, 4) = Data1.Recordset("categ")
      End If
      If IsNull(Data1.Recordset("edad")) = False Then
         DBGrid1.TextMatrix(Xcann, 5) = Data1.Recordset("edad")
      End If
      If IsNull(Data1.Recordset("referen")) = False Then
         DBGrid1.TextMatrix(Xcann, 6) = Data1.Recordset("referen")
      End If
      If IsNull(Data1.Recordset("telef")) = False Then
         DBGrid1.TextMatrix(Xcann, 7) = Data1.Recordset("telef")
      End If
      If IsNull(Data1.Recordset("obsmot")) = False Then
         DBGrid1.TextMatrix(Xcann, 8) = Data1.Recordset("obsmot")
      End If
      If IsNull(Data1.Recordset("codmot")) = False Then
         DBGrid1.TextMatrix(Xcann, 9) = Data1.Recordset("codmot")
      End If
      If IsNull(Data1.Recordset("movilpas")) = False Then
         DBGrid1.TextMatrix(Xcann, 10) = Data1.Recordset("movilpas")
      End If
      If IsNull(Data1.Recordset("fecpas")) = False Then
         DBGrid1.TextMatrix(Xcann, 11) = Data1.Recordset("fecpas")
      End If
      If IsNull(Data1.Recordset("horpas")) = False Then
         DBGrid1.TextMatrix(Xcann, 12) = Data1.Recordset("horpas")
      End If
      If IsNull(Data1.Recordset("nrolla")) = False Then
         DBGrid1.TextMatrix(Xcann, 13) = Data1.Recordset("nrolla")
      End If
       
      DBGrid1.rows = DBGrid1.rows + 1
      Data1.Recordset.MoveNext
      Xcann = Xcann + 1
   Loop
End If
bok.Enabled = True
Exit Sub

Quebusca:
         If Err.Number = 13 Then
            MsgBox "Error de datos en la búsqueda, verifique", vbInformation
            bok.Enabled = True
         Else
            MsgBox "Error de datos al buscar, verifique!", vbInformation
            bok.Enabled = True
         End If
         

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Combo1.ListIndex = 0 Then
      mh.Visible = True
      md.Visible = True
      Text1.Visible = False
      md.SetFocus
   Else
      If Combo1.ListIndex = 1 Then
         md.Visible = False
         mh.Visible = False
         Text1.Visible = True
         Text1.SetFocus
      Else
         If Combo1.ListIndex = 2 Then
            md.Visible = False
            mh.Visible = False
            Text1.Visible = True
            Text1.SetFocus
         End If
      End If
   End If
End If

End Sub

Private Sub Combo1_LostFocus()
If Combo1.ListIndex = 0 Then
   mh.Visible = True
   md.Visible = True
   Text1.Visible = False
Else
   If Combo1.ListIndex = 1 Then
      md.Visible = False
      mh.Visible = False
      Text1.Visible = True
   Else
      If Combo1.ListIndex = 2 Then
         md.Visible = False
         mh.Visible = False
         Text1.Visible = True
      
      End If
   End If
End If
   
End Sub

Private Sub DBGrid1_DblClick()
'frm_largador.data_lla.Recordset.FindFirst "nrolla =" & Data1.Recordset("nrolla")
'If Not frm_largador.data_lla.Recordset.NoMatch Then
If Option1.Value = True Then
    Adodc1.RecordSource = "Select * from llamado where nrolla =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 13)
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        frm_largador.txt_nro.Text = Adodc1.Recordset("nrolla")
        frm_largador.mfecha.Text = Adodc1.Recordset("fecha")
        frm_largador.txt_hora.Text = Format(Adodc1.Recordset("hora"), "HH:mm")
        frm_largador.txt_usua.Text = Adodc1.Recordset("usuario")
        If IsNull(Adodc1.Recordset("nomodif")) = False Then
           frm_largador.Check1.Value = Adodc1.Recordset("nomodif")
        Else
           frm_largador.Check1.Value = 0
        End If
        If IsNull(Adodc1.Recordset("matric")) = False Then
           frm_largador.txt_mat.Text = Adodc1.Recordset("matric")
        Else
           frm_largador.txt_mat.Text = 0
        End If
        If IsNull(Adodc1.Recordset("timbre")) = False Then
           frm_largador.cbotimbre.ListIndex = Adodc1.Recordset("timbre")
        Else
           frm_largador.cbotimbre.ListIndex = -1
        End If
        If IsNull(Adodc1.Recordset("valor_timbre")) = False Then
           frm_largador.t_timbre.Text = Adodc1.Recordset("valor_timbre")
        Else
           frm_largador.t_timbre.Text = ""
        End If
        If IsNull(Adodc1.Recordset("segui_covid")) = False Then
           frm_largador.chcovid.Value = Adodc1.Recordset("segui_covid")
        Else
           frm_largador.chcovid.Value = 0
        End If
        If IsNull(Adodc1.Recordset("hora_anterior")) = False Then
           frm_largador.labanthor.Caption = Adodc1.Recordset("hora_anterior")
        Else
           frm_largador.labanthor.Caption = ""
        End If
        If IsNull(Adodc1.Recordset("nombre")) = False Then
           frm_largador.txt_nomb.Text = Adodc1.Recordset("nombre")
        Else
           frm_largador.txt_nomb.Text = ""
        End If
        If IsNull(Adodc1.Recordset("edad")) = False Then
           frm_largador.txt_edad.Text = Adodc1.Recordset("edad")
        Else
           frm_largador.txt_edad.Text = ""
        End If
        If IsNull(Adodc1.Recordset("mes")) = False Then
           If Adodc1.Recordset("mes") > 10 Then
              frm_largador.txt_costo.Text = Adodc1.Recordset("mes")
           Else
              frm_largador.txt_costo.Text = 0
           End If
        Else
           frm_largador.txt_costo.Text = 0
        End If
        If Format(Adodc1.Recordset("fecha"), "yyyy/mm/dd") >= Format("31/12/2016", "yyyy/mm/dd") Then
           If IsNull(Adodc1.Recordset("realiza")) = False Then
              frm_largador.chtmut.Value = Adodc1.Recordset("realiza")
           Else
              frm_largador.chtmut.Value = 0
           End If
        Else
           frm_largador.chtmut.Value = 0
        End If
        If IsNull(Adodc1.Recordset("ano")) = False Then
           If Adodc1.Recordset("ano") > 10 Then
              frm_largador.txt_boleta.Text = Adodc1.Recordset("ano")
           Else
              frm_largador.txt_boleta.Text = 0
           End If
        Else
           frm_largador.txt_boleta.Text = 0
        End If
        
        If Adodc1.Recordset("pend") = 4 Then
           frm_largador.Command5.Visible = True
           frm_largador.Frame2.Visible = False
        Else
           frm_largador.Command5.Visible = False
           frm_largador.Frame2.Visible = True
        End If
        If IsNull(Adodc1.Recordset("unied")) = False Then
           If Adodc1.Recordset("unied") = 3 Then
              frm_largador.cboed.ListIndex = 0
           Else
              If Adodc1.Recordset("unied") = 2 Then
                 frm_largador.cboed.ListIndex = 1
              Else
                 If Adodc1.Recordset("unied") = 1 Then
                    frm_largador.cboed.ListIndex = 2
                 Else
                    frm_largador.cboed.ListIndex = 0
                 End If
              End If
           End If
        Else
           frm_largador.cboed.ListIndex = 0
        End If
        If IsNull(Adodc1.Recordset("categ")) = False Then
           frm_largador.txt_cat.Text = Adodc1.Recordset("categ")
        Else
           frm_largador.txt_cat.Text = ""
        End If
        If IsNull(Adodc1.Recordset("nomcat")) = False Then
           frm_largador.txt_nomcat.Text = Adodc1.Recordset("nomcat")
        Else
           frm_largador.txt_nomcat.Text = ""
        End If
        If IsNull(Adodc1.Recordset("ci")) = False Then
           frm_largador.txt_ced.Text = Int(Adodc1.Recordset("ci"))
        Else
           frm_largador.txt_ced.Text = 0
        End If
        If IsNull(Adodc1.Recordset("telef")) = False Then
           frm_largador.txt_tel.Text = Adodc1.Recordset("telef")
        Else
           frm_largador.txt_tel.Text = ""
        End If
        If IsNull(Adodc1.Recordset("codzon")) = False Then
           If Adodc1.Recordset("codzon") = 2 Then
              frm_largador.cbozona.ListIndex = 1
           Else
              If Adodc1.Recordset("codzon") = 3 Then
                 frm_largador.cbozona.ListIndex = 2
              Else
                 If Adodc1.Recordset("codzon") = 4 Then
                    frm_largador.cbozona.ListIndex = 3
                 Else
                    If Adodc1.Recordset("codzon") = 5 Then
                       frm_largador.cbozona.ListIndex = 4
                    Else
                       If Adodc1.Recordset("codzon") = 6 Then
                          frm_largador.cbozona.ListIndex = 5
                       Else
                          If Adodc1.Recordset("codzon") = 7 Then
                             frm_largador.cbozona.ListIndex = 6
                          Else
                             frm_largador.cbozona.ListIndex = 0
                          End If
                       End If
                    End If
                 End If
              End If
           End If
        Else
           frm_largador.cbozona.ListIndex = 0
        End If
        If IsNull(Adodc1.Recordset("base")) = False Then
           frm_largador.cbobase.Text = Adodc1.Recordset("base")
        Else
           frm_largador.cbobase.Text = 0
        End If
        If IsNull(Adodc1.Recordset("referen")) = False Then
           frm_largador.txt_direc.Text = Adodc1.Recordset("referen")
        Else
           frm_largador.txt_direc.Text = ""
        End If
        If IsNull(Adodc1.Recordset("obs")) = False Then
           frm_largador.txt_obs.Text = Adodc1.Recordset("obs")
        Else
           frm_largador.txt_obs.Text = ""
        End If
        If IsNull(Adodc1.Recordset("motcon")) = False Then
           frm_largador.txt_ante.Text = Adodc1.Recordset("motcon")
        Else
           frm_largador.txt_ante.Text = ""
        End If
        If IsNull(Adodc1.Recordset("obsmot")) = False Then
           frm_largador.txt_mot.Text = Adodc1.Recordset("obsmot")
        Else
           frm_largador.txt_mot.Text = ""
        End If
        If IsNull(Adodc1.Recordset("codmot")) = False Then
           If Adodc1.Recordset("codmot") = "R" Then
              frm_largador.cbocolor.ListIndex = 2
           Else
              If Adodc1.Recordset("codmot") = "A" Then
                 frm_largador.cbocolor.ListIndex = 1
              Else
                 If Adodc1.Recordset("codmot") = "C" Then
                    frm_largador.cbocolor.ListIndex = 3
                 Else
                    If Adodc1.Recordset("codmot") = "Z" Then
                       frm_largador.cbocolor.ListIndex = 4
                    Else
                       If Adodc1.Recordset("codmot") = "N" Then
                          frm_largador.cbocolor.ListIndex = 5
                       Else
                          frm_largador.cbocolor.ListIndex = 0
                       End If
                    End If
                 End If
              End If
           End If
        Else
           frm_largador.cbocolor.ListIndex = 0
        End If
        If frm_largador.cbocolor.Text = "VERDE" Then
           frm_largador.cbocolor.BackColor = &HC000&
        Else
           If frm_largador.cbocolor.Text = "ROJO" Then
              frm_largador.cbocolor.BackColor = &HFF&
           Else
              If frm_largador.cbocolor.Text = "AMARILLO" Then
                 frm_largador.cbocolor.BackColor = &HFFFF&
              Else
                 If frm_largador.cbocolor.Text = "CELESTE" Then
                    frm_largador.cbocolor.BackColor = &HFFFF00
                 Else
                    If frm_largador.cbocolor.Text = "AZUL" Then
                       frm_largador.cbocolor.BackColor = &HC00000
                    Else
                       If frm_largador.cbocolor.Text = "NEGRO" Then
                          frm_largador.cbocolor.BackColor = &H80000006
                       Else
                          frm_largador.cbocolor.BackColor = &HFFFFFF
                       End If
                    End If
                 End If
              End If
           End If
        End If
        If IsNull(Adodc1.Recordset("movilpas")) = False Then
           frm_largador.txt_movil.Text = Adodc1.Recordset("movilpas")
        Else
           frm_largador.txt_movil.Text = ""
        End If
        If IsNull(Adodc1.Recordset("fecpas")) = False Then
           frm_largador.mfecasig.Text = Format(Adodc1.Recordset("fecpas"), "dd/mm/yyyy")
        Else
           frm_largador.mfecasig.Text = "__/__/____"
        End If
        If IsNull(Adodc1.Recordset("horpas")) = False Then
           frm_largador.txt_horasig.Text = Format(Adodc1.Recordset("horpas"), "HH:mm")
        Else
           frm_largador.txt_horasig.Text = ""
        End If
        If IsNull(Adodc1.Recordset("aft")) = False Then
           frm_largador.Label40.Caption = "AFT:" & Adodc1.Recordset("aft")
        Else
           frm_largador.Label40.Caption = ""
        End If
        
        If IsNull(Adodc1.Recordset("fecsali")) = False Then
           frm_largador.msalida.Text = Format(Adodc1.Recordset("fecsali"), "dd/mm/yyyy")
        Else
           frm_largador.msalida.Text = "__/__/____"
        End If
        If IsNull(Adodc1.Recordset("horsali")) = False Then
           frm_largador.txt_horsal.Text = Format(Adodc1.Recordset("horsali"), "HH:mm")
        Else
           frm_largador.txt_horsal.Text = ""
        End If
        If IsNull(Adodc1.Recordset("fec_llega")) = False Then
           frm_largador.mllegada.Text = Format(Adodc1.Recordset("fec_llega"), "dd/mm/yyyy")
        Else
           frm_largador.mllegada.Text = "__/__/____"
        End If
        If IsNull(Adodc1.Recordset("hor_llega")) = False Then
           frm_largador.txt_horlle.Text = Format(Adodc1.Recordset("hor_llega"), "HH:mm")
        Else
           frm_largador.txt_horlle.Text = ""
        End If
        If IsNull(Adodc1.Recordset("fec_rea")) = False Then
           frm_largador.mtd.Text = Format(Adodc1.Recordset("fec_rea"), "dd/mm/yyyy")
        Else
           frm_largador.mtd.Text = "__/__/____"
        End If
        If IsNull(Adodc1.Recordset("hor_rea")) = False Then
           If Adodc1.Recordset("hor_rea") <> "" Then
              frm_largador.txt_hortd.Text = Format(Adodc1.Recordset("hor_rea"), "HH:mm")
           Else
              frm_largador.txt_hortd.Text = "__:__"
           End If
        Else
           frm_largador.txt_hortd.Text = "__:__"
        End If
        If IsNull(Adodc1.Recordset("diag")) = False Then
           frm_largador.txt_diag.Text = Adodc1.Recordset("diag")
        Else
           frm_largador.txt_diag.Text = ""
        End If
        If IsNull(Adodc1.Recordset("colormot")) = False Then
           If Adodc1.Recordset("colormot") = "R" Then
              frm_largador.cbocolfin.ListIndex = 2
           Else
              If Adodc1.Recordset("colormot") = "A" Then
                 frm_largador.cbocolfin.ListIndex = 1
              Else
                 If Adodc1.Recordset("colormot") = "V" Then
                    frm_largador.cbocolfin.ListIndex = 0
                 Else
                    If Adodc1.Recordset("colormot") = "N" Then
                       frm_largador.cbocolfin.ListIndex = 3
                    Else
                       frm_largador.cbocolfin.Text = ""
                    End If
                 End If
              End If
           End If
        Else
           frm_largador.cbocolfin.Text = ""
        End If
        If IsNull(Adodc1.Recordset("nommed")) = False Then
           frm_largador.dbcbomed.Text = Adodc1.Recordset("nommed")
        Else
           frm_largador.dbcbomed.ListField = ""
           frm_largador.dbcbomed.BoundColumn = ""
           frm_largador.dbcbomed.Text = ""
           frm_largador.dbcbomed.ListField = "med_nombre"
           frm_largador.dbcbomed.BoundColumn = "med_nombre"
        End If
        If IsNull(Adodc1.Recordset("codmed")) = False Then
           frm_largador.txt_codmed.Text = Adodc1.Recordset("codmed")
        Else
           frm_largador.txt_codmed.Text = 0
        End If
        If IsNull(Adodc1.Recordset("trasla")) = False Then
           If Adodc1.Recordset("trasla") > 0 Then
              If Adodc1.Recordset("trasla") > 10 Then
                 If Adodc1.Recordset("trasla") = 11 Then
                    frm_largador.cbotras.ListIndex = 8
                 Else
                    frm_largador.cbotras.ListIndex = 9
                 End If
              Else
                 frm_largador.cbotras.ListIndex = Adodc1.Recordset("trasla")
              End If
           Else
              frm_largador.cbotras.ListIndex = 0
           End If
        Else
           frm_largador.cbotras.ListIndex = 0
        End If
        If IsNull(Adodc1.Recordset("lugar")) = False Then
           frm_largador.txt_lugar.Text = Adodc1.Recordset("lugar")
        Else
           frm_largador.txt_lugar.Text = ""
        End If
        If IsNull(Adodc1.Recordset("hsald")) = False Then
           frm_largador.txt_trassal.Text = Format(Adodc1.Recordset("hsald"), "HH:mm")
        Else
           frm_largador.txt_trassal.Text = ""
        End If
        If IsNull(Adodc1.Recordset("hllega")) = False Then
           frm_largador.txt_enca.Text = Format(Adodc1.Recordset("hllega"), "HH:mm")
        Else
           frm_largador.txt_enca.Text = ""
        End If
        If IsNull(Adodc1.Recordset("hzona")) = False Then
           frm_largador.txt_enzona.Text = Format(Adodc1.Recordset("hzona"), "HH:mm")
        Else
           frm_largador.txt_enzona.Text = ""
        End If
        If IsNull(Adodc1.Recordset("movtras")) = False Then
           frm_largador.txt_movtra.Text = Adodc1.Recordset("movtras")
        Else
           frm_largador.txt_movtra.Text = ""
        End If
        If IsNull(Adodc1.Recordset("totdem")) = False Then
           frm_largador.txt_demora.Text = Format(Adodc1.Recordset("totdem"), "HH:mm")
        Else
           frm_largador.txt_demora.Text = ""
        End If
        If IsNull(Adodc1.Recordset("dcobr")) = False Then
            frm_largador.Combo1.Text = Adodc1.Recordset("dcobr")
        Else
            frm_largador.Combo1.Text = ""
        End If
        If IsNull(Adodc1.Recordset("activo")) = False Then
           frm_largador.Label3.Caption = Format(Adodc1.Recordset("activo"), "HH:mm:ss")
        Else
           frm_largador.Label3.Caption = "00:00:00"
        End If
        If IsNull(Adodc1.Recordset("timdes")) = False Then
           frm_largador.Label39.Caption = Adodc1.Recordset("timdes")
        Else
           frm_largador.Label39.Caption = "Sin Largar"
        End If
        If IsNull(Adodc1.Recordset("motmov")) = True Then
           frm_largador.txt_locali.Text = ""
        Else
           frm_largador.txt_locali.Text = Adodc1.Recordset("motmov")
        End If
        If IsNull(Adodc1.Recordset("mm")) = True Then
           frm_largador.Label41.Caption = 0
        Else
           frm_largador.Label41.Caption = Adodc1.Recordset("mm")
        End If
        If IsNull(Adodc1.Recordset("thh")) = True Then
           frm_largador.Label42.Caption = 0
        Else
           frm_largador.Label42.Caption = Adodc1.Recordset("thh")
        End If
        If IsNull(Adodc1.Recordset("tmm")) = True Then
           frm_largador.Label43.Caption = 0
        Else
           frm_largador.Label43.Caption = Adodc1.Recordset("tmm")
        End If
        If IsNull(Adodc1.Recordset("pasado")) = True Then
           frm_largador.Label44.Caption = 0
        Else
           frm_largador.Label44.Caption = Adodc1.Recordset("pasado")
        End If
        If IsNull(Adodc1.Recordset("ano")) = True Then
           frm_largador.Label45.Caption = 0
        Else
           frm_largador.Label45.Caption = Adodc1.Recordset("ano")
        End If
        If IsNull(Adodc1.Recordset("mes")) = True Then
           frm_largador.Label46.Caption = -1
        Else
           frm_largador.Label46.Caption = Adodc1.Recordset("mes")
        End If
        If IsNull(Adodc1.Recordset("timsi")) = True Then
           frm_largador.Label48.Caption = 0
        Else
           frm_largador.Label48.Caption = Adodc1.Recordset("timsi")
        End If
        If IsNull(Adodc1.Recordset("enfer")) = True Then
           frm_largador.Check2.Value = 0
        Else
           frm_largador.Check2.Value = Adodc1.Recordset("enfer")
        End If
        If IsNull(Adodc1.Recordset("motcance")) = True Then
           frm_largador.txt_quien.Text = ""
        Else
           frm_largador.txt_quien.Text = Adodc1.Recordset("motcance")
        End If
        If IsNull(Adodc1.Recordset("hh")) = True Then
           frm_largador.Combo3.ListIndex = -1
        Else
           frm_largador.Combo3.ListIndex = Adodc1.Recordset("hh")
        End If
        If IsNull(Adodc1.Recordset("cancela")) = True Then
           If IsNull(Adodc1.Recordset("hor_cance")) = False Then
              frm_largador.txt_salca.Text = Adodc1.Recordset("hor_cance")
           Else
              frm_largador.txt_salca.Text = ""
           End If
        End If
        data3.RecordSource = "Select * from resplla where nro =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 13)
        data3.Refresh
        If data3.Recordset.RecordCount > 0 Then
           If IsNull(data3.Recordset("telef")) = False Then
              If data3.Recordset("telef") = "RECIBO" Then
                 frm_largador.Combo2.ListIndex = 0
              Else
                 If data3.Recordset("telef") = "CONFORME" Then
                    frm_largador.Combo2.ListIndex = 1
                 Else
                    frm_largador.Combo2.ListIndex = -1
                 End If
              End If
           Else
              frm_largador.Combo2.ListIndex = -1
           End If
           If IsNull(data3.Recordset("movil_rea")) = False Then
              If data3.Recordset("movil_rea") > 0 Then
                 data_chof.RecordSource = "Select * from movil where nromov =" & data3.Recordset("movil_rea")
                 data_chof.Refresh
'                 frm_largador.data_chof.RecordSource = "Select * from movil where nromov =" & data3.Recordset("movil_rea")
'                 frm_largador.data_chof.Refresh
                 If data_chof.Recordset.RecordCount > 0 Then
                    frm_largador.labcodchof.Caption = data3.Recordset("movil_rea")
                    frm_largador.labnomchof.Caption = "Chof.:" & data_chof.Recordset("chofer")
                 Else
                    frm_largador.labcodchof.Caption = 0
                    frm_largador.labnomchof.Caption = ""
                 End If
              Else
                 frm_largador.labcodchof.Caption = 0
                 frm_largador.labnomchof.Caption = ""
              End If
           Else
              frm_largador.labcodchof.Caption = 0
              frm_largador.labnomchof.Caption = ""
           End If
           If IsNull(data3.Recordset("pasado")) = False Then
              frm_largador.Check4.Value = data3.Recordset("pasado")
           Else
              frm_largador.Check4.Value = 0
           End If
           If IsNull(data3.Recordset("hzona")) = False Then
              frm_largador.labcmt.Visible = True
              frm_largador.labcmt.Caption = "PASADO A CMT HORA:" & Format(data3.Recordset("hzona"), "HH:mm")
              If IsNull(data3.Recordset("mm")) = False Then
                 If data3.Recordset("mm") = 1 Then
                    frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & " NO RESUELTO H."
                    If IsNull(data3.Recordset("hsald")) = False Then
                       frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & data3.Recordset("hsald")
                    End If
                    If IsNull(data3.Recordset("totend")) = False Then
                       If data3.Recordset("totend") = "R" Then
                          frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & " RECLASIFICA A ROJO"
                       End If
                       If data3.Recordset("totend") = "A" Then
                          frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & " RECLASIFICA A AMARILLO"
                       End If
                    End If
                 End If
                 If data3.Recordset("mm") = 2 Then
                    frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & " RESUELTO HORA:"
                    If IsNull(data3.Recordset("hor_rea")) = False Then
                       frm_largador.labcmt.Caption = frm_largador.labcmt.Caption & data3.Recordset("hor_rea")
                    End If
                 End If
              End If
           Else
              frm_largador.labcmt.Caption = ""
              frm_largador.labcmt.Visible = False
           End If
           If IsNull(data3.Recordset("mes")) = False Then
              frm_largador.t_codced.Text = Int(data3.Recordset("mes"))
           Else
              frm_largador.t_codced.Text = 0
           End If
           If IsNull(data3.Recordset("movilpas")) = False Then
              frm_largador.dbcbomed2.ListField = ""
              frm_largador.dbcbomed2.BoundColumn = ""
'              frm_largador.dbcbomed2.DataSource = "" 'data_med2
              data4.RecordSource = "Select * from medicos where med_cod =" & data3.Recordset("movilpas")
              data4.Refresh
              If data4.Recordset.RecordCount > 0 Then
                 frm_largador.dbcbomed2.Text = data4.Recordset("med_nombre")
              Else
                 frm_largador.dbcbomed2.ListField = ""
                 frm_largador.dbcbomed2.BoundColumn = ""
                 frm_largador.dbcbomed2.Text = ""
                 frm_largador.dbcbomed2.ListField = "med_nombre"
                 frm_largador.dbcbomed2.BoundColumn = "med_nombre"
              End If
              frm_largador.txt_codmed2.Text = data3.Recordset("movilpas")
              frm_largador.dbcbomed2.ListField = "med_nombre"
              frm_largador.dbcbomed2.BoundColumn = "med_nombre"
'              frm_largador.dbcbomed2.DataSource = "data_med2" 'data_med2
           Else
              frm_largador.txt_codmed2.Text = 0
              frm_largador.dbcbomed2.ListField = ""
              frm_largador.dbcbomed2.BoundColumn = ""
              frm_largador.dbcbomed2.Text = ""
              frm_largador.dbcbomed2.ListField = "med_nombre"
              frm_largador.dbcbomed2.BoundColumn = "med_nombre"
           End If
           If IsNull(data3.Recordset("fec_llega")) = False Then
              frm_largador.mftrassol.Text = Format(data3.Recordset("fec_llega"), "dd/mm/yyyy")
           Else
              frm_largador.mftrassol.Text = "__/__/____"
           End If
           If IsNull(data3.Recordset("hor_llega")) = False Then
              frm_largador.mhtrassol.Text = Format(data3.Recordset("hor_llega"), "HH:mm")
           Else
              frm_largador.mhtrassol.Text = "__:__"
           End If
        Else
           frm_largador.txt_codmed2.Text = 0
           frm_largador.Check4.Value = 0
           frm_largador.dbcbomed2.Text = ""
           frm_largador.mftrassol.Text = "__/__/____"
           frm_largador.mhtrassol.Text = "__:__"
           frm_largador.t_codced.Text = 0
           frm_largador.labcmt.Caption = ""
           frm_largador.labcmt.Visible = False
        End If
        
        Unload Me
    Else
        MsgBox "Este registro solo puede ser visualizado desde Ver llamados", vbInformation, "Mensaje"
        DBGrid1.SetFocus
    End If
Else
    
    Data2.RecordSource = "Select * from llamado where nrolla =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 13)
    Data2.Refresh
    If Data2.Recordset.RecordCount > 0 Then
        frm_largador.txt_nro.Text = Data2.Recordset("nrolla")
        frm_largador.mfecha.Text = Data2.Recordset("fecha")
        frm_largador.txt_hora.Text = Format(Data2.Recordset("hora"), "HH:mm")
        frm_largador.txt_usua.Text = Data2.Recordset("usuario")
        frm_largador.Label26.Caption = "Traslado:"
        If IsNull(Data2.Recordset("matric")) = False Then
           frm_largador.txt_mat.Text = Data2.Recordset("matric")
        Else
           frm_largador.txt_mat.Text = 0
        End If
        If IsNull(Data2.Recordset("nombre")) = False Then
           frm_largador.txt_nomb.Text = Data2.Recordset("nombre")
        Else
           frm_largador.txt_nomb.Text = ""
        End If
        If IsNull(Data2.Recordset("edad")) = False Then
           frm_largador.txt_edad.Text = Data2.Recordset("edad")
        Else
           frm_largador.txt_edad.Text = ""
        End If
        If IsNull(Data2.Recordset("mes")) = False Then
           If Data2.Recordset("mes") > 10 Then
              frm_largador.txt_costo.Text = Data2.Recordset("mes")
           Else
              frm_largador.txt_costo.Text = 0
           End If
        Else
           frm_largador.txt_costo.Text = 0
        End If
        If Format(Data2.Recordset("fecha"), "yyyy/mm/dd") >= Format("01/07/2016", "yyyy/mm/dd") Then
           If frm_largador.txt_costo.Text <> "" Then
              If frm_largador.txt_costo.Text > 10 Then
                 If IsNull(Data2.Recordset("realiza")) = False Then
                    frm_largador.chtmut.Value = Data2.Recordset("realiza")
                 Else
                    frm_largador.chtmut.Value = 0
                 End If
              Else
                 frm_largador.chtmut.Value = 0
              End If
           Else
              frm_largador.chtmut.Value = 0
           End If
        Else
           frm_largador.chtmut.Value = 0
        End If
        If IsNull(Data2.Recordset("ano")) = False Then
           If Data2.Recordset("ano") > 10 Then
              frm_largador.txt_boleta.Text = Data2.Recordset("ano")
           Else
              frm_largador.txt_boleta.Text = 0
           End If
        Else
           frm_largador.txt_boleta.Text = 0
        End If
        
        If Data2.Recordset("pend") = 4 Then
           frm_largador.Command5.Visible = True
           frm_largador.Frame2.Visible = False
        Else
           frm_largador.Command5.Visible = False
           frm_largador.Frame2.Visible = True
        End If
        If IsNull(Data2.Recordset("unied")) = False Then
           If Data2.Recordset("unied") = 3 Then
              frm_largador.cboed.ListIndex = 0
           Else
              If Data2.Recordset("unied") = 2 Then
                 frm_largador.cboed.ListIndex = 1
              Else
                 If Data2.Recordset("unied") = 1 Then
                    frm_largador.cboed.ListIndex = 2
                 Else
                    frm_largador.cboed.ListIndex = 0
                 End If
              End If
           End If
        Else
           frm_largador.cboed.ListIndex = 0
        End If
        If IsNull(Data2.Recordset("categ")) = False Then
           frm_largador.txt_cat.Text = Data2.Recordset("categ")
        Else
           frm_largador.txt_cat.Text = ""
        End If
        If IsNull(Data2.Recordset("nomcat")) = False Then
           frm_largador.txt_nomcat.Text = Data2.Recordset("nomcat")
        Else
           frm_largador.txt_nomcat.Text = ""
        End If
        If IsNull(Data2.Recordset("ci")) = False Then
           frm_largador.txt_ced.Text = Int(Data2.Recordset("ci"))
        Else
           frm_largador.txt_ced.Text = 0
        End If
        If IsNull(Data2.Recordset("telef")) = False Then
           frm_largador.txt_tel.Text = Data2.Recordset("telef")
        Else
           frm_largador.txt_tel.Text = ""
        End If
        If IsNull(Data2.Recordset("codzon")) = False Then
           If Data2.Recordset("codzon") = 2 Then
              frm_largador.cbozona.ListIndex = 1
           Else
              If Data2.Recordset("codzon") = 3 Then
                 frm_largador.cbozona.ListIndex = 2
              Else
                 If Data2.Recordset("codzon") = 4 Then
                    frm_largador.cbozona.ListIndex = 3
                 Else
                    If Data2.Recordset("codzon") = 5 Then
                       frm_largador.cbozona.ListIndex = 4
                    Else
                       If Data2.Recordset("codzon") = 6 Then
                          frm_largador.cbozona.ListIndex = 5
                       Else
                          If Data2.Recordset("codzon") = 7 Then
                             frm_largador.cbozona.ListIndex = 1
                          Else
                             frm_largador.cbozona.ListIndex = 0
                          End If
                       End If
                    End If
                 End If
              End If
           End If
        Else
           frm_largador.cbozona.ListIndex = 0
        End If
        If IsNull(Data2.Recordset("base")) = False Then
           frm_largador.cbobase.Text = Data2.Recordset("base")
        Else
           frm_largador.cbobase.Text = 0
        End If
        If IsNull(Data2.Recordset("referen")) = False Then
           frm_largador.txt_direc.Text = Data2.Recordset("referen")
        Else
           frm_largador.txt_direc.Text = ""
        End If
        If IsNull(Data2.Recordset("obs")) = False Then
           frm_largador.txt_obs.Text = Data2.Recordset("obs")
        Else
           frm_largador.txt_obs.Text = ""
        End If
        If IsNull(Data2.Recordset("obsmot")) = False Then
           frm_largador.txt_ante.Text = Data2.Recordset("obsmot")
        Else
           frm_largador.txt_ante.Text = ""
        End If
        If IsNull(Data2.Recordset("motcon")) = False Then
           frm_largador.txt_mot.Text = Data2.Recordset("motcon")
        Else
           frm_largador.txt_mot.Text = ""
        End If
        If IsNull(Data2.Recordset("codmot")) = False Then
           If Data2.Recordset("codmot") = "R" Then
              frm_largador.cbocolor.ListIndex = 2
           Else
              If Data2.Recordset("codmot") = "A" Then
                 frm_largador.cbocolor.ListIndex = 1
              Else
                 If Data2.Recordset("codmot") = "C" Then
                    frm_largador.cbocolor.ListIndex = 3
                 Else
                    If Data2.Recordset("codmot") = "Z" Then
                       frm_largador.cbocolor.ListIndex = 4
                    Else
                       If Data2.Recordset("codmot") = "N" Then
                          frm_largador.cbocolor.ListIndex = 5
                       Else
                          frm_largador.cbocolor.ListIndex = 0
                       End If
                    End If
                 End If
              End If
           End If
        Else
           frm_largador.cbocolor.ListIndex = 0
        End If
        If frm_largador.cbocolor.Text = "VERDE" Then
           frm_largador.cbocolor.BackColor = &HC000&
        Else
           If frm_largador.cbocolor.Text = "ROJO" Then
              frm_largador.cbocolor.BackColor = &HFF&
           Else
              If frm_largador.cbocolor.Text = "AMARILLO" Then
                 frm_largador.cbocolor.BackColor = &HFFFF&
              Else
                 If frm_largador.cbocolor.Text = "CELESTE" Then
                    frm_largador.cbocolor.BackColor = &HFFFF00
                 Else
                    If frm_largador.cbocolor.Text = "AZUL" Then
                       frm_largador.cbocolor.BackColor = &HC00000
                    Else
                       If frm_largador.cbocolor.Text = "NEGRO" Then
                          frm_largador.cbocolor.BackColor = &H80000006
                       Else
                          frm_largador.cbocolor.BackColor = &HFFFFFF
                       End If
                    End If
                 End If
              End If
           End If
        End If
        If IsNull(Data2.Recordset("movilpas")) = False Then
           frm_largador.txt_movil.Text = Data2.Recordset("movilpas")
        Else
           frm_largador.txt_movil.Text = ""
        End If
        If IsNull(Data2.Recordset("fecpas")) = False Then
           frm_largador.mfecasig.Text = Format(Data2.Recordset("fecpas"), "dd/mm/yyyy")
        Else
           frm_largador.mfecasig.Text = "__/__/____"
        End If
        If IsNull(Data2.Recordset("horpas")) = False Then
           frm_largador.txt_horasig.Text = Format(Data2.Recordset("horpas"), "HH:mm")
        Else
           frm_largador.txt_horasig.Text = ""
        End If
        If IsNull(Data2.Recordset("fecsali")) = False Then
           frm_largador.msalida.Text = Format(Data2.Recordset("fecsali"), "dd/mm/yyyy")
        Else
           frm_largador.msalida.Text = "__/__/____"
        End If
        If IsNull(Data2.Recordset("horsali")) = False Then
           frm_largador.txt_horsal.Text = Format(Data2.Recordset("horsali"), "HH:mm")
        Else
           frm_largador.txt_horsal.Text = ""
        End If
        If IsNull(Data2.Recordset("fec_llega")) = False Then
           frm_largador.mllegada.Text = Format(Data2.Recordset("fec_llega"), "dd/mm/yyyy")
        Else
           frm_largador.mllegada.Text = "__/__/____"
        End If
        If IsNull(Data2.Recordset("hor_llega")) = False Then
           frm_largador.txt_horlle.Text = Format(Data2.Recordset("hor_llega"), "HH:mm")
        Else
           frm_largador.txt_horlle.Text = ""
        End If
        If IsNull(Data2.Recordset("fec_rea")) = False Then
           frm_largador.mtd.Text = Format(Data2.Recordset("fec_rea"), "dd/mm/yyyy")
        Else
           frm_largador.mtd.Text = "__/__/____"
        End If
        If IsNull(Data2.Recordset("hor_rea")) = False Then
           If Data2.Recordset("hor_rea") <> "" Then
              frm_largador.txt_hortd.Text = Format(Data2.Recordset("hor_rea"), "HH:mm")
           Else
              frm_largador.txt_hortd.Text = "__:__"
           End If
        Else
           frm_largador.txt_hortd.Text = "__:__"
        End If
        If IsNull(Data2.Recordset("diag")) = False Then
           frm_largador.txt_diag.Text = Data2.Recordset("diag")
        Else
           frm_largador.txt_diag.Text = ""
        End If
        If IsNull(Data2.Recordset("colormot")) = False Then
           If Data2.Recordset("colormot") = "R" Then
              frm_largador.cbocolfin.ListIndex = 2
           Else
              If Data2.Recordset("colormot") = "A" Then
                 frm_largador.cbocolfin.ListIndex = 1
              Else
                 If Data2.Recordset("colormot") = "V" Then
                    frm_largador.cbocolfin.ListIndex = 0
                 Else
                    If Data2.Recordset("colormot") = "N" Then
                       frm_largador.cbocolfin.ListIndex = 3
                    Else
                       frm_largador.cbocolfin.Text = ""
                    End If
                 End If
              End If
           End If
        Else
           frm_largador.cbocolfin.Text = ""
        End If
        If IsNull(Data2.Recordset("nommed")) = False Then
           frm_largador.dbcbomed.Text = Data2.Recordset("nommed")
        Else
           frm_largador.dbcbomed.ListField = ""
           frm_largador.dbcbomed.BoundColumn = ""
           frm_largador.dbcbomed.Text = ""
           frm_largador.dbcbomed.ListField = "med_nombre"
           frm_largador.dbcbomed.BoundColumn = "med_nombre"
        End If
        If IsNull(Data2.Recordset("codmed")) = False Then
           frm_largador.txt_codmed.Text = Data2.Recordset("codmed")
        Else
           frm_largador.txt_codmed.Text = 0
        End If
        If IsNull(Data2.Recordset("trasla")) = False Then
           If Data2.Recordset("trasla") > 0 Then
              frm_largador.cbotras.ListIndex = Data2.Recordset("trasla")
           Else
              frm_largador.cbotras.ListIndex = 0
           End If
        Else
           frm_largador.cbotras.ListIndex = 0
        End If
        If IsNull(Data2.Recordset("lugar")) = False Then
           frm_largador.txt_lugar.Text = Data2.Recordset("lugar")
        Else
           frm_largador.txt_lugar.Text = ""
        End If
        If IsNull(Data2.Recordset("hsald")) = False Then
           frm_largador.txt_trassal.Text = Format(Data2.Recordset("hsald"), "HH:mm")
        Else
           frm_largador.txt_trassal.Text = ""
        End If
        If IsNull(Data2.Recordset("hllega")) = False Then
           frm_largador.txt_enca.Text = Format(Data2.Recordset("hllega"), "HH:mm")
        Else
           frm_largador.txt_enca.Text = ""
        End If
        If IsNull(Data2.Recordset("hzona")) = False Then
           frm_largador.txt_enzona.Text = Format(Data2.Recordset("hzona"), "HH:mm")
        Else
           frm_largador.txt_enzona.Text = ""
        End If
        If IsNull(Data2.Recordset("movtras")) = False Then
           frm_largador.txt_movtra.Text = Data2.Recordset("movtras")
        Else
           frm_largador.txt_movtra.Text = ""
        End If
        If IsNull(Data2.Recordset("totdem")) = False Then
           frm_largador.txt_demora.Text = Format(Data2.Recordset("totdem"), "HH:mm")
        Else
           frm_largador.txt_demora.Text = ""
        End If
        If IsNull(Data2.Recordset("dcobr")) = False Then
            frm_largador.Combo1.Text = Data2.Recordset("dcobr")
        Else
            frm_largador.Combo1.Text = ""
        End If
        If IsNull(Data2.Recordset("activo")) = False Then
           frm_largador.Label3.Caption = Format(Data2.Recordset("activo"), "HH:mm:ss")
        Else
           frm_largador.Label3.Caption = "00:00:00"
        End If
        If IsNull(Data2.Recordset("timdes")) = False Then
           frm_largador.Label39.Caption = Data2.Recordset("timdes")
        Else
           frm_largador.Label39.Caption = "Sin Largar"
        End If
        If IsNull(Data2.Recordset("motmov")) = True Then
           frm_largador.txt_locali.Text = ""
        Else
           frm_largador.txt_locali.Text = Data2.Recordset("motmov")
        End If
        If IsNull(Data2.Recordset("mm")) = True Then
           frm_largador.Label41.Caption = 0
        Else
           frm_largador.Label41.Caption = Data2.Recordset("mm")
        End If
        If IsNull(Data2.Recordset("thh")) = True Then
           frm_largador.Label42.Caption = 0
        Else
           frm_largador.Label42.Caption = Data2.Recordset("thh")
        End If
        If IsNull(Data2.Recordset("tmm")) = True Then
           frm_largador.Label43.Caption = 0
        Else
           frm_largador.Label43.Caption = Data2.Recordset("tmm")
        End If
        If IsNull(Data2.Recordset("pasado")) = True Then
           frm_largador.Label44.Caption = 0
        Else
           frm_largador.Label44.Caption = Data2.Recordset("pasado")
        End If
        If IsNull(Data2.Recordset("ano")) = True Then
           frm_largador.Label45.Caption = 0
        Else
           frm_largador.Label45.Caption = Data2.Recordset("ano")
        End If
        If IsNull(Data2.Recordset("mes")) = True Then
           frm_largador.Label46.Caption = -1
        Else
           frm_largador.Label46.Caption = Data2.Recordset("mes")
        End If
        If IsNull(Data2.Recordset("timsi")) = True Then
           frm_largador.Label48.Caption = 0
        Else
           frm_largador.Label48.Caption = Data2.Recordset("timsi")
        End If
        If IsNull(Data2.Recordset("enfer")) = True Then
           frm_largador.Check2.Value = 0
        Else
           frm_largador.Check2.Value = Data2.Recordset("enfer")
        End If
        If IsNull(Data2.Recordset("motcance")) = True Then
           frm_largador.txt_quien.Text = ""
        Else
           frm_largador.txt_quien.Text = Data2.Recordset("motcance")
        End If
        If IsNull(Data2.Recordset("hh")) = True Then
           frm_largador.Combo3.ListIndex = -1
        Else
           frm_largador.Combo3.ListIndex = Data2.Recordset("hh")
        End If
        If IsNull(Data2.Recordset("cancela")) = True Then
           If IsNull(Data2.Recordset("hor_cance")) = False Then
              frm_largador.txt_salca.Text = Data2.Recordset("hor_cance")
           Else
              frm_largador.txt_salca.Text = ""
           End If
        End If
    
    End If
    
    MsgBox "Opción de respaldos sin habilitar", vbInformation
    Unload Me
End If


End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   DBGrid1_DblClick
End If

End Sub

Private Sub Form_Load()
Dim Xfecante As Date
Xfecante = Date - 2
data3.ConnectionString = "dsn=" & Xconexrmt
Adodc1.ConnectionString = "dsn=" & Xconexrmt

data4.ConnectionString = "dsn=" & Xconexrmt
Data1.ConnectionString = "dsn=" & Xconexrmt
Data2.DatabaseName = App.path & "\llamado.mdb"

data_chof.ConnectionString = "dsn=" & Xconexrmt
Data1.RecordSource = "Select * from llamado where fecha >='" & Format(Xfecante, "yyyy/mm/dd") & "' order by fecha desc, hora"
Data1.Refresh
Combo1.ListIndex = 0
data4.RecordSource = "medicos"
data4.Refresh
Option1.Value = True
DBGrid1.rows = 2
DBGrid1.Cols = 14
DBGrid1.TextMatrix(0, 0) = "FECHA"
DBGrid1.ColWidth(0) = 1200
DBGrid1.TextMatrix(0, 1) = "HORA"
DBGrid1.ColWidth(1) = 900
DBGrid1.TextMatrix(0, 2) = "NOMBRE"
DBGrid1.ColWidth(2) = 2900
DBGrid1.TextMatrix(0, 3) = "MATRICULA"
DBGrid1.ColWidth(3) = 1500
DBGrid1.TextMatrix(0, 4) = "CATEG."
DBGrid1.ColWidth(4) = 1200
DBGrid1.TextMatrix(0, 5) = "EDAD"
DBGrid1.ColWidth(5) = 500
DBGrid1.TextMatrix(0, 6) = "DIRECCION"
DBGrid1.ColWidth(6) = 2900
DBGrid1.TextMatrix(0, 7) = "TELEFONO"
DBGrid1.ColWidth(7) = 1500
DBGrid1.TextMatrix(0, 8) = "MOT.CONSULTA"
DBGrid1.ColWidth(8) = 3900
DBGrid1.TextMatrix(0, 9) = "CODIGO"
DBGrid1.ColWidth(9) = 600
DBGrid1.TextMatrix(0, 10) = "MOVIL"
DBGrid1.ColWidth(10) = 500
DBGrid1.TextMatrix(0, 11) = "FEC.PAS"
DBGrid1.ColWidth(11) = 1200
DBGrid1.TextMatrix(0, 12) = "HORA PAS"
DBGrid1.ColWidth(12) = 700
DBGrid1.TextMatrix(0, 13) = "ID"
DBGrid1.ColWidth(13) = 1200

Dim Xcann As Integer
Xcann = 1
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      If IsNull(Data1.Recordset("fecha")) = False Then
         DBGrid1.TextMatrix(Xcann, 0) = Data1.Recordset("fecha")
      End If
      If IsNull(Data1.Recordset("hora")) = False Then
         DBGrid1.TextMatrix(Xcann, 1) = Data1.Recordset("hora")
      End If
      If IsNull(Data1.Recordset("nombre")) = False Then
         DBGrid1.TextMatrix(Xcann, 2) = Data1.Recordset("nombre")
      End If
      If IsNull(Data1.Recordset("matric")) = False Then
         DBGrid1.TextMatrix(Xcann, 3) = Data1.Recordset("matric")
      End If
      If IsNull(Data1.Recordset("categ")) = False Then
         DBGrid1.TextMatrix(Xcann, 4) = Data1.Recordset("categ")
      End If
      If IsNull(Data1.Recordset("edad")) = False Then
         DBGrid1.TextMatrix(Xcann, 5) = Data1.Recordset("edad")
      End If
      If IsNull(Data1.Recordset("referen")) = False Then
         DBGrid1.TextMatrix(Xcann, 6) = Data1.Recordset("referen")
      End If
      If IsNull(Data1.Recordset("telef")) = False Then
         DBGrid1.TextMatrix(Xcann, 7) = Data1.Recordset("telef")
      End If
      If IsNull(Data1.Recordset("obsmot")) = False Then
         DBGrid1.TextMatrix(Xcann, 8) = Data1.Recordset("obsmot")
      End If
      If IsNull(Data1.Recordset("codmot")) = False Then
         DBGrid1.TextMatrix(Xcann, 9) = Data1.Recordset("codmot")
      End If
      If IsNull(Data1.Recordset("movilpas")) = False Then
         DBGrid1.TextMatrix(Xcann, 10) = Data1.Recordset("movilpas")
      End If
      If IsNull(Data1.Recordset("fecpas")) = False Then
         DBGrid1.TextMatrix(Xcann, 11) = Data1.Recordset("fecpas")
      End If
      If IsNull(Data1.Recordset("horpas")) = False Then
         DBGrid1.TextMatrix(Xcann, 12) = Data1.Recordset("horpas")
      End If
      If IsNull(Data1.Recordset("nrolla")) = False Then
         DBGrid1.TextMatrix(Xcann, 13) = Data1.Recordset("nrolla")
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

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bok.SetFocus
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bok.SetFocus
End If

End Sub
