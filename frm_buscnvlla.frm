VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_buscnvlla 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de convenios"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   Icon            =   "frm_buscnvlla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   9555
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3480
      Top             =   1680
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
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF0000&
      Caption         =   "Buscar por Razón Social"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "odbc;dsn=sappnew;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "convenio"
      Top             =   3000
      Visible         =   0   'False
      Width           =   3180
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
      Left            =   8760
      Picture         =   "frm_buscnvlla.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   3600
      Width           =   615
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_buscnvlla.frx":09CC
      Height          =   3135
      Left            =   120
      OleObjectBlob   =   "frm_buscnvlla.frx":09E0
      TabIndex        =   2
      Top             =   480
      Width           =   9255
   End
   Begin VB.TextBox txt_de 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label2 
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
      Left            =   1200
      TabIndex        =   5
      Top             =   3600
      Width           =   7335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Descripción:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   4680
      Picture         =   "frm_buscnvlla.frx":171F
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2055
   End
End
Attribute VB_Name = "frm_buscnvlla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
frm_largador.txt_cat.Text = Data1.Recordset("cnv_codigo")
frm_largador.txt_nomcat.Text = Data1.Recordset("cnv_desc")
Unload Me

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   DBGrid1_DblClick
End If

End Sub

Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If DBGrid1.Text <> "" Then
    Adodc1.RecordSource = "Select * from convenio where cnv_codigo ='" & DBGrid1.Text & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
       Adodc1.Recordset.MoveFirst
       If IsNull(Adodc1.Recordset("cnv_motbaj")) = False Then
          Label2.Caption = Adodc1.Recordset("cnv_motbaj")
       Else
          Label2.Caption = ""
       End If
    End If
End If


End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "select top 100, * from convenio where cnv_codigo >='" & frm_largador.txt_cat.Text & "' and cnv_alta ='" & "SI" & "' and cnv_umpago not in (1) and cnv_fbaja is null order by cnv_codigo"
Data1.Refresh
Adodc1.ConnectionString = "dsn=" & Xconexrmt


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Private Sub txt_de_KeyPress(KeyAscii As Integer)
If Check1.Value = 1 Then
   Data1.RecordSource = "select top 110, * from convenio where cnv_entre >='" & txt_de.Text & "' and cnv_alta ='" & "SI" & "' and cnv_umpago not in (1) and cnv_fbaja is null order by cnv_entre"
Else
   Data1.RecordSource = "select top 110, * from convenio where cnv_desc >='" & txt_de.Text & "' and cnv_alta ='" & "SI" & "' and cnv_umpago not in (1) and cnv_fbaja is null order by cnv_desc"
End If
Data1.Refresh
If KeyAscii = 13 Then
   DBGrid1.SetFocus
End If

End Sub
