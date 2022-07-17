VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_histconve 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial precios convenios"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7245
   Icon            =   "frm_histconve.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7245
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_lineas 
      Height          =   375
      Left            =   120
      Top             =   3000
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "data_lineas"
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
      Height          =   375
      Left            =   6480
      Picture         =   "frm_histconve.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cerrar"
      Top             =   3240
      Width           =   495
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4895
      SortKey         =   1
      View            =   3
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod.Conv."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "F.Desde"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "F.Hasta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Importe"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Usuario"
         Object.Width           =   2540
      EndProperty
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Convenio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   840
      Picture         =   "frm_histconve.frx":0B14
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   3015
   End
End
Attribute VB_Name = "frm_histconve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
Dim Xcount As Integer
Xcount = 1
data_lineas.ConnectionString = "dsn=" & Xconexrmt
ListView1.ListItems.Clear
Label2.Caption = frm_convenios.txt_desc.Text
data_lineas.RecordSource = "Select * from cnv_prec where cnv_codigo ='" & frm_convenios.txt_cod.Text & "' order by cnv_desde DESC"
data_lineas.Refresh
If data_lineas.Recordset.RecordCount <> 0 Then
   data_lineas.Recordset.MoveFirst
   Do While Not data_lineas.Recordset.EOF
       If IsNull(data_lineas.Recordset("cnv_codigo")) = False Then
          ListView1.ListItems.Add Xcount, , data_lineas.Recordset("cnv_codigo")
       Else
          ListView1.ListItems.Add Xcount, , " "
       End If
       If IsNull(data_lineas.Recordset("cnv_desde")) = False Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(data_lineas.Recordset("cnv_desde"), "dd/mm/yyyy")
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , " "
       End If
       If IsNull(data_lineas.Recordset("cnv_hasta")) = False Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(data_lineas.Recordset("cnv_hasta"), "dd/mm/yyyy")
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , " "
       End If
       If IsNull(data_lineas.Recordset("precio")) = False Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("precio")
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
       End If
       If IsNull(data_lineas.Recordset("usuario")) = False Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_lineas.Recordset("usuario")
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "Sin Datos"
       End If
       
       data_lineas.Recordset.MoveNext
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

