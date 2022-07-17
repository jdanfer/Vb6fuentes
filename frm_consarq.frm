VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_consarq 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta historial de arqueos"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8955
   Icon            =   "frm_consarq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Matricula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cobr."
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Arqueo"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Mes"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Año"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Nro.Documento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF8080&
      Caption         =   "Consultar en arqueo actual"
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
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   7560
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
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
   Begin VB.TextBox t_mat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8160
      Picture         =   "frm_consarq.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Buscar..."
      Top             =   480
      Width           =   495
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frm_consarq.frx":0B14
      Left            =   3120
      List            =   "frm_consarq.frx":0B16
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frm_consarq.frx":0B18
      Left            =   2400
      List            =   "frm_consarq.frx":0B1A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Matrícula (opcional)"
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
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Arqueo a consultar:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   1560
      Picture         =   "frm_consarq.frx":0B1C
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2895
   End
End
Attribute VB_Name = "frm_consarq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim Xcount As Long
Dim Xelarqueo As String
On Error GoTo Vererr

frm_consarq.MousePointer = 11
Command1.Enabled = False
If Check1.Value = 1 Then
   If t_mat.Text <> "" Then
      Adodc1.RecordSource = "Select * from arqueo where matricula =" & t_mat.Text & " order by nrorec"
      Adodc1.Refresh
   Else
      Adodc1.RecordSource = "Select * from arqueo order by nrorec"
      Adodc1.Refresh
   End If
Else
    If Val(Combo1.Text) < 10 Then
       Xelarqueo = "arq0" & Trim(Combo1.Text) & Mid(Trim(Combo2.Text), 3, 2)
    Else
       Xelarqueo = "arq" & Trim(Combo1.Text) & Mid(Trim(Combo2.Text), 3, 2)
    End If
    If t_mat.Text <> "" Then
       Adodc1.RecordSource = "Select * from " & Xelarqueo & " where matricula =" & t_mat.Text & " order by nrorec"
       Adodc1.Refresh
    Else
       Adodc1.RecordSource = "Select * from " & Xelarqueo & " order by nrorec"
       Adodc1.Refresh
    End If
End If

Xcount = 1
ListView1.ListItems.Clear
DoEvents

If Adodc1.Recordset.RecordCount > 0 Then
   Adodc1.Recordset.MoveFirst
   Do While Not Adodc1.Recordset.EOF
       If IsNull(Adodc1.Recordset("matricula")) = False Then
          ListView1.ListItems.Add Xcount, , Adodc1.Recordset("matricula")
       Else
          ListView1.ListItems.Add Xcount, , "0"
       End If
       If IsNull(Adodc1.Recordset("nombre")) = False Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Adodc1.Recordset("nombre")
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , " "
       End If
       If IsNull(Adodc1.Recordset("cob")) = True Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Adodc1.Recordset("cob")
       End If
       If IsNull(Adodc1.Recordset("arqueo")) = True Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "X"
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Adodc1.Recordset("arqueo")
       End If
       If IsNull(Adodc1.Recordset("mes")) = True Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Adodc1.Recordset("mes")
       End If
       If IsNull(Adodc1.Recordset("ano")) = True Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Adodc1.Recordset("ano")
       End If
       If IsNull(Adodc1.Recordset("nrorec")) = True Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Adodc1.Recordset("nrorec")
       End If
       If IsNull(Adodc1.Recordset("total")) = True Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Adodc1.Recordset("total")
       End If
       If IsNull(Adodc1.Recordset("fecha")) = True Then
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "XX/XX/XXXX"
       Else
          ListView1.ListItems.Item(Xcount).ListSubItems.Add , , Format(Adodc1.Recordset("fecha"), "dd/mm/yyyy")
       End If
       Adodc1.Recordset.MoveNext
       Xcount = Xcount + 1
   Loop
Else
    frm_consarq.MousePointer = 0
    MsgBox "No existe historial", vbInformation, "Ver historial"
End If
frm_consarq.MousePointer = 0
Command1.Enabled = True

Exit Sub

Vererr:
       If Err.Number = 91 Then
          frm_consarq.MousePointer = 0
          MsgBox "Error al consultar datos, verifique", vbInformation
       Else
          frm_consarq.MousePointer = 0
          MsgBox "Error al consultar datos, verifique si existe arqueo", vbInformation
       End If

End Sub

Private Sub Form_Load()
Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "4"
Combo1.AddItem "5"
Combo1.AddItem "6"
Combo1.AddItem "7"
Combo1.AddItem "8"
Combo1.AddItem "9"
Combo1.AddItem "10"
Combo1.AddItem "11"
Combo1.AddItem "12"
Dim Xanio As Integer
Xanio = Year(Date)
Xanio = Xanio - 4
For Xanio = Xanio To Year(Date)
    Combo2.AddItem Xanio
Next
Adodc1.ConnectionString = "dsn=" & Xconexrmt


End Sub

Private Sub Form_Resize()
With Image1
   .Top = 0
   .Left = 0
   .Width = Me.Width
   .Height = Me.Height
End With

End Sub

