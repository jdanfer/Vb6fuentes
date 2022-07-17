VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_impteso 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Caja"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   Icon            =   "frm_impteso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4815
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc data_teso 
      Height          =   375
      Left            =   360
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
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data_teso"
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
   Begin VB.TextBox txt_rub 
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
      Left            =   2400
      TabIndex        =   7
      Top             =   1680
      Width           =   2055
   End
   Begin Crystal.CrystalReport crteso 
      Left            =   4320
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton bcance 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Picture         =   "frm_impteso.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton bacep 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "frm_impteso.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Procesar"
      Top             =   2400
      Width           =   735
   End
   Begin MSMask.MaskEdBox mhasta 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mdesde 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   4800
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "RUBRO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   4800
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "FECHA HASTA:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "FECHA DESDE:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   1800
      Picture         =   "frm_impteso.frx":0F56
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "frm_impteso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bacep_Click()
Dim Saldoini, Saldofin As Double
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")

MiBaseact.Execute "Delete * from inf_teso"

data_inf.RecordSource = "inf_teso"
data_inf.Refresh

If IsDate(mdesde.Text) = True Then
   If IsDate(mhasta.Text) = True Then
      If txt_rub.Text = "" Then
         data_teso.RecordSource = "Select * from tesorero where fecha >= '" & Format(mdesde.Text, "yyyy-mm-dd") & "' And fecha <= '" & Format(mhasta.Text, "yyyy-mm-dd") & "' And usuario ='" & WNombase & "' order by nromov"
         data_teso.Refresh
      Else
         data_teso.RecordSource = "Select * from tesorero where cod_rub =" & txt_rub.Text & " And fecha >= '" & Format(mdesde.Text, "yyyy-mm-dd") & "' And fecha <= '" & Format(mhasta.Text, "yyyy-mm-dd") & "' And usuario ='" & WNombase & "' order by nromov"
         data_teso.Refresh
      End If
      If data_teso.Recordset.RecordCount > 0 Then
         frm_impteso.MousePointer = 11
         data_teso.Recordset.MoveFirst
         If data_teso.Recordset("concep") = "E" Then
            Saldoini = data_teso.Recordset("saldos") - data_teso.Recordset("monto")
         Else
            Saldoini = data_teso.Recordset("saldos") + data_teso.Recordset("monto")
         End If
         Saldofin = Saldoini
         Do While Not data_teso.Recordset.EOF
            data_inf.Recordset.AddNew
            data_inf.Recordset("fecha") = Format(data_teso.Recordset("fecha"), "dd/mm/yyyy")
            data_inf.Recordset("hora") = Format(data_teso.Recordset("hora"), "HH:mm")
            data_inf.Recordset("cod_rub") = data_teso.Recordset("cod_rub")
            data_inf.Recordset("nom_rub") = data_teso.Recordset("nom_rub")
            data_inf.Recordset("moneda") = data_teso.Recordset("moneda")
            data_inf.Recordset("monto") = data_teso.Recordset("monto")
            data_inf.Recordset("obs") = data_teso.Recordset("obs")
            data_inf.Recordset("cod_debe") = data_teso.Recordset("cod_debe")
            data_inf.Recordset("cod_haber") = data_teso.Recordset("cod_haber")
            data_inf.Recordset("saldos") = data_teso.Recordset("saldos")
            data_inf.Recordset("concep") = data_teso.Recordset("concep")
            data_inf.Recordset("saldou") = data_teso.Recordset("saldou")
            data_inf.Recordset("base") = data_teso.Recordset("base")
            data_inf.Recordset("descon") = data_teso.Recordset("descon")
            data_inf.Recordset("iva") = data_teso.Recordset("iva")
            data_inf.Recordset("impiva") = data_teso.Recordset("impiva")
            data_inf.Recordset("usuario") = data_teso.Recordset("usuario")
            data_inf.Recordset("saldos") = Saldoini
            data_inf.Recordset.Update
            If data_teso.Recordset("concep") = "E" Then
               Saldofin = Saldofin + data_teso.Recordset("monto")
            Else
               Saldofin = Saldofin - data_teso.Recordset("monto")
            End If
            data_teso.Recordset.MoveNext
         Loop
         data_inf.Recordset.MoveFirst
         Do While Not data_inf.Recordset.EOF
            data_inf.Recordset.Edit
            data_inf.Recordset("saldou") = Saldofin
            data_inf.Recordset.Update
            data_inf.Recordset.MoveNext
         Loop
         data_inf.RecordSource = "Select * from inf_teso order by fecha"
         data_inf.Refresh
         data_inf.Recordset.MoveFirst
         frm_impteso.MousePointer = 0
         If txt_rub.Text = "" Then
            crteso.ReportFileName = App.Path & "\infteso.rpt"
            crteso.ReportTitle = "CAJA DE TESORERIA DESDE " + Format(mdesde.Text, "dd/mm/yyyy") + " HASTA " + Format(mhasta.Text, "dd/mm/yyyy")
            crteso.Action = 1
         Else
            crteso.ReportFileName = App.Path & "\inftesocob.rpt"
            crteso.ReportTitle = "INFORME DESDE " + Format(mdesde.Text, "dd/mm/yyyy") + " HASTA " + Format(mhasta.Text, "dd/mm/yyyy")
            crteso.Action = 1
         End If
      Else
         MsgBox "No existen registros", vbInformation, "Mensaje"
         
      End If
   Else
      MsgBox "Verifique FECHA", vbCritical, "Mensaje"
      mhasta.SetFocus
   End If
Else
   MsgBox "Verifique FECHA", vbCritical, "Mensaje"
   mdesde.SetFocus
End If

End Sub

Private Sub bcance_Click()
Unload Me

End Sub

Private Sub Form_Load()
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_inf.DatabaseName = App.Path & "\informes.mdb"
data_inf.RecordSource = "inf_teso"
data_inf.Refresh
'data_teso.DatabaseName = App.Path & "\sapp.mdb"
'data_teso.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_teso.ConnectionString = "DSN=" & Xconexrmt
crteso.ReportFileName = App.Path & "\infteso.rpt"
txt_rub.Text = ""
mdesde.Text = Format(Date, "dd/mm/yyyy")
mhasta.Text = Format(Date, "dd/mm/yyyy")

End Sub

Private Sub Form_Resize()
With Image1
    .Top = 0
    .Left = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mdesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhasta.SetFocus
End If

End Sub

Private Sub mhasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_rub.SetFocus
End If

End Sub

Private Sub txt_rub_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bacep.SetFocus
End If
 
End Sub
