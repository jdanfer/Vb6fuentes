VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ctrolsolhc 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copias de HC pendientes de ser retiradas por el usuario"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_ctrolsolhc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   10215
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data2 
      Height          =   375
      Left            =   6360
      Top             =   3600
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "data2"
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   840
      Top             =   1920
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
      BackColor       =   &H008080FF&
      Caption         =   "Copia de HC retorna a Registros"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9600
      Picture         =   "frm_ctrolsolhc.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   3480
      Width           =   495
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "Nombre"
         Object.Width           =   6244
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "b"
         Text            =   "Servicio"
         Object.Width           =   6421
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "c"
         Text            =   "Cédula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "d"
         Text            =   "No.Fact."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "e"
         Text            =   "Fecha Fact"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "f"
         Text            =   "BASE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "g"
         Text            =   "Matrícula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Fecha Envío"
         Object.Width           =   2540
      EndProperty
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
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Copias de HC pendientes de ser retiradas por el usuario:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   4680
      Picture         =   "frm_ctrolsolhc.frx":09CC
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1335
   End
End
Attribute VB_Name = "frm_ctrolsolhc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
Dim XNfac, XNmat, XNcod As Long
Dim XNfe As Date
Command2.Enabled = False
If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then
'   MsgBox "Esta"
'   MsgBox "ES:" & ListView1.SelectedItem.ListSubItems(1).Text
   XNcod = ListView1.SelectedItem.ListSubItems(2).Text
   XNfac = ListView1.SelectedItem.ListSubItems(3).Text
   XNfe = ListView1.SelectedItem.ListSubItems(4).Text
   XNmat = ListView1.SelectedItem.ListSubItems(5).Text
   data2.RecordSource = "Select * from linmmdd where fecha ='" & Format(XNfe, "yyyy-mm-dd") & "' And factura =" & XNfac & " And cod_prod =" & 991
   data2.Refresh
   If data2.Recordset.RecordCount > 0 Then
      If IsNull(data2.Recordset("servicio")) = True Then
'         data2.Recordset.Edit
         data2.Recordset("servicio") = 2
         data2.Recordset("vto") = Date
         data2.Recordset.Update
      Else
         If data2.Recordset("servicio") = 2 Then
         Else
'            data2.Recordset.Edit
            data2.Recordset("servicio") = 2
            data2.Recordset("vto") = Date
            data2.Recordset.Update
         End If
      End If
   End If
Else
   XNcod = ListView1.SelectedItem.ListSubItems(2).Text
   XNfac = ListView1.SelectedItem.ListSubItems(3).Text
   XNfe = ListView1.SelectedItem.ListSubItems(4).Text
   XNmat = ListView1.SelectedItem.ListSubItems(5).Text
   data2.RecordSource = "Select * from linmmdd where fecha ='" & Format(XNfe, "yyyy-mm-dd") & "' And factura =" & XNfac & " And cod_prod =" & 991
   data2.Refresh
   If data2.Recordset.RecordCount > 0 Then
      If IsNull(data2.Recordset("servicio")) = True Then
'         data2.Recordset.Edit
         data2.Recordset("servicio") = 0
         data2.Recordset.Update
      Else
         If data2.Recordset("servicio") = 0 Then
         Else
'            data2.Recordset.Edit
            data2.Recordset("servicio") = 0
            data2.Recordset.Update
         End If
      End If
   End If

End If
MsgBox "Procesado terminado"
Command2.Enabled = True

End Sub

Private Sub Form_Load()
Dim Xcount As Long
Dim Xedtiene As Long
Dim Xdef As Date
Xdef = Date - 90
Xcount = 1
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data2.Connect = "odbc;dsn=" & Xconexrmt & ";"
data2.ConnectionString = "dsn=" & Xconexrmt
data1.ConnectionString = "dsn=" & Xconexrmt

If WElusuario = "SDOMINGUEZ" Or WElusuario = "PATRICIAH" Or WElusuario = "SPEREZ" Or WElusuario = "JFERNAN" Or WElusuario = "MCOSTA" Then
   data1.RecordSource = "Select * from linmmdd where cod_prod =" & 991 & " And fecha >='" & Format(Xdef, "yyyy-mm-dd") & "' order by fecha"
   data1.Refresh
   Command2.Enabled = True
Else
   data1.RecordSource = "Select * from linmmdd where cod_prod =" & 991 & " And fecha >='" & Format(Xdef, "yyyy-mm-dd") & "' And base =" & frmabm.data_parsec.Recordset("base") & " order by fecha"
   data1.Refresh
   Command2.Enabled = False
End If
ListView1.ListItems.Clear
If data1.Recordset.RecordCount > 0 Then
   data1.Recordset.MoveFirst
   Do While Not data1.Recordset.EOF
      If IsNull(data1.Recordset("servicio")) = True Then
         data1.Recordset.MoveNext
      Else
         If data1.Recordset("servicio") = 1 Or data1.Recordset("servicio") = 2 Then
            data1.Recordset.MoveNext
         Else
            If IsNull(data1.Recordset("cod_cli")) = False Then
               ListView1.ListItems.Add Xcount, , data1.Recordset("nom_cli")
            Else
               ListView1.ListItems.Add Xcount, , "NN"
            End If
            If IsNull(data1.Recordset("nom_prod")) = True Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "SIN DATOS"
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data1.Recordset("nom_prod")
            End If
            If IsNull(data1.Recordset("ced_socio")) = True Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data1.Recordset("ced_socio")
            End If
            If IsNull(data1.Recordset("factura")) = True Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data1.Recordset("factura")
            End If
            If IsNull(data1.Recordset("fecha")) = True Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data1.Recordset("fecha")
            End If
            If IsNull(data1.Recordset("base")) = True Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data1.Recordset("base")
            End If
            If IsNull(data1.Recordset("cod_cli")) = True Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , "0"
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data1.Recordset("cod_cli")
            End If
            If IsNull(data1.Recordset("realizada")) = True Then
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , ""
            Else
               ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data1.Recordset("realizada")
            End If
            
            Xcount = Xcount + 1
            data1.Recordset.MoveNext
         End If
      End If
   Loop
   Label3.Caption = Xcount - 1
Else
   MsgBox "No existen solicitudes", vbInformation, "Ver historial"
   Label3.Caption = 0
End If
'End If

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'MsgBox "El item está cheq"
Dim XNfac, XNmat, XNcod As Long
Dim XNfe As Date

If ListView1.ListItems.Item(ListView1.SelectedItem.Index).Checked = True Then
'   MsgBox "Esta"
'   MsgBox "ES:" & ListView1.SelectedItem.ListSubItems(1).Text
   XNcod = ListView1.SelectedItem.ListSubItems(2).Text
   XNfac = ListView1.SelectedItem.ListSubItems(3).Text
   XNfe = ListView1.SelectedItem.ListSubItems(4).Text
   XNmat = ListView1.SelectedItem.ListSubItems(5).Text
   data2.RecordSource = "Select * from linmmdd where fecha ='" & Format(XNfe, "yyyy-mm-dd") & "' And factura =" & XNfac & " And cod_prod =" & 991
   data2.Refresh
   If data2.Recordset.RecordCount > 0 Then
      If IsNull(data2.Recordset("servicio")) = True Then
'         data2.Recordset.Edit
         data2.Recordset("servicio") = 1
         data2.Recordset("vto") = Date
         data2.Recordset.Update
      Else
         If data2.Recordset("servicio") = 1 Then
         Else
'            data2.Recordset.Edit
            data2.Recordset("servicio") = 1
            data2.Recordset("vto") = Date
            data2.Recordset.Update
         End If
      End If
   End If
Else
   XNcod = ListView1.SelectedItem.ListSubItems(2).Text
   XNfac = ListView1.SelectedItem.ListSubItems(3).Text
   XNfe = ListView1.SelectedItem.ListSubItems(4).Text
   XNmat = ListView1.SelectedItem.ListSubItems(5).Text
   data2.RecordSource = "Select * from linmmdd where fecha ='" & Format(XNfe, "yyyy-mm-dd") & "' And factura =" & XNfac & " And cod_prod =" & 991
   data2.Refresh
   If data2.Recordset.RecordCount > 0 Then
      If IsNull(data2.Recordset("servicio")) = True Then
'         data2.Recordset.Edit
         data2.Recordset("servicio") = 0
         data2.Recordset.Update
      Else
         If data2.Recordset("servicio") = 0 Then
         Else
'            data2.Recordset.Edit
            data2.Recordset("servicio") = 0
            data2.Recordset.Update
         End If
      End If
   End If

End If


End Sub

