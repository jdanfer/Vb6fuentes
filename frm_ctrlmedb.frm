VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ctrlmedb 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medicación enviada a la base para su entrega"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11265
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_ctrlmedb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   11265
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_lin 
      Height          =   375
      Left            =   7680
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
      Caption         =   "data_lin"
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
   Begin MSAdodcLib.Adodc data_actu 
      Height          =   330
      Left            =   7560
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      Caption         =   "data_actu"
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
   Begin Crystal.CrystalReport cr1 
      Left            =   4200
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox mfh 
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mfd 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10560
      Picture         =   "frm_ctrlmedb.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   3720
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Medicación pendiente de entrega"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10935
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1920
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSAdodcLib.Adodc data_lindos 
         Height          =   375
         Left            =   6480
         Top             =   1200
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
         Caption         =   "data_lindos"
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
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Informe pedidos"
         Height          =   495
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Emite todo lo solicitado en un rango de fechas"
         Top             =   2880
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Informe entregas"
         Height          =   495
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Emite todo lo entregado en un rango de fechas."
         Top             =   2880
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Entregado"
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2880
         Width           =   2535
      End
      Begin MSComctlLib.ListView lis1 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   4471
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "a"
            Text            =   "FECHA"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "b"
            Text            =   "HORA"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "c"
            Text            =   "NOMBRE"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "d"
            Text            =   "CONVENIO"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "e"
            Text            =   "MEDICAMENTO"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Key             =   "f"
            Text            =   "BASE"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Key             =   "g"
            Text            =   "CEDULA"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Key             =   "h"
            Text            =   "FACTURA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Key             =   "i"
            Text            =   "LINEA"
            Object.Width           =   353
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Key             =   "j"
            Text            =   "MATRICULA"
            Object.Width           =   1940
         EndProperty
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "BASE:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ingrese rango de fechas:"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   9240
      TabIndex        =   5
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Usuario actual:"
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   5880
      Picture         =   "frm_ctrlmedb.frx":09CC
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "frm_ctrlmedb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xind As Long
Dim Xmatme, Xfacme, Xlinme As Long
Dim Xmensajeme As String
Dim Xcountt As Long
Dim Xdeff As Date
Xdeff = Date - 31
Dim Xnommedi As String
Xcountt = 1
Xmensajeme = MsgBox("Desea procesar los registros marcados cómo ENTREGADOS?", vbInformation + vbYesNo, "Control")
Xind = 0

If Xmensajeme = vbYes Then
   Xnommedi = InputBox("Ingrese descripción de medicamento entregado", "Medicamento entregado")
   If Xnommedi = "" Then
      MsgBox "Debe ingresar nombre de medicamento para procesar, NO SE REALIZA EL PROCESO DE ENTREGA", vbCritical, "SAPP"
   Else
       For Xind = 1 To lis1.ListItems.count
           lis1.ListItems(Xind).Selected = True
           If lis1.ListItems.Item(lis1.SelectedItem.index).Checked = True Then
    '       MsgBox "Chequeado"
              Xmatme = lis1.SelectedItem.ListSubItems(9).Text
              Xlinme = lis1.SelectedItem.ListSubItems(8).Text
              Xfacme = lis1.SelectedItem.ListSubItems(7).Text
              data_actu.RecordSource = "Select * from linmmdd where cod_cli =" & Xmatme & " and factura =" & Xfacme & " and linea =" & Xlinme
              data_actu.Refresh
              If data_actu.Recordset.RecordCount > 0 Then
    '             data_actu.Recordset.Edit
                 If data_actu.Recordset("dias") = 3 Then
                    data_actu.Recordset("dias") = 8
                 Else
                    data_actu.Recordset("dias") = 11
                 End If
                 data_actu.Recordset("realizada") = Date
                 data_actu.Recordset("numero") = Welnrou
                 data_actu.Recordset("zona") = Mid(Xnommedi, 1, 25)
    '             data_actu.Recordset("margen_prd") = Val(Mid(Format(Time, "HH:mm"), 1, 2))
    '             data_actu.Recordset("pre_prod") = Val(Mid(Format(Time, "HH:mm"), 4, 2))
                 data_actu.Recordset.Update
                 data_actu.Refresh
              End If
           End If
       Next Xind
       If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "AGUILLEN" Or frm_menu.data_parse.Recordset("base") = 16 Or frm_menu.data_parse.Recordset("base") = 91 Then
          data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,3,5) order by fecha,base"
          data_lin.Refresh
       Else
          data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,3,5) and base =" & frm_menu.data_parse.Recordset("base") & " order by fecha,factura"
          data_lin.Refresh
       End If
       lis1.ListItems.Clear
       If data_lin.Recordset.RecordCount > 0 Then
          data_lin.Recordset.MoveFirst
          Do While Not data_lin.Recordset.EOF
             If data_lin.Recordset("cod_prod") = 60103 Or _
                data_lin.Recordset("cod_prod") = 60105 Or _
                data_lin.Recordset("cod_prod") = 60106 Or _
                data_lin.Recordset("cod_prod") = 60107 Or _
                data_lin.Recordset("cod_prod") = 60108 Or _
                data_lin.Recordset("cod_prod") = 60109 Then
                If IsNull(data_lin.Recordset("fecha")) = False Then
                   lis1.ListItems.Add Xcountt, , data_lin.Recordset("fecha")
                Else
                   lis1.ListItems.Add Xcountt, , "01/01/2010"
                End If
                If IsNull(data_lin.Recordset("hora")) = False Then
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("hora")
                Else
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
                End If
                If IsNull(data_lin.Recordset("nom_cli")) = False Then
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("nom_cli")
                Else
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "NN"
                End If
                If IsNull(data_lin.Recordset("convenio")) = False Then
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("convenio")
                Else
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "S/C"
                End If
                If IsNull(data_lin.Recordset("nom_medic")) = False Then
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("nom_medic")
                Else
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "NO INGRESADO"
                End If
                If IsNull(data_lin.Recordset("base")) = False Then
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("base")
                Else
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "99"
                End If
                If IsNull(data_lin.Recordset("ced_socio")) = False Then
                   If IsNull(data_lin.Recordset("fact")) = False Then
                      lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("ced_socio") & "-" & data_lin.Recordset("fact")
                   Else
                      lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("ced_socio") & "-0"
                   End If
                Else
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
                End If
                If IsNull(data_lin.Recordset("factura")) = False Then
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("factura")
                Else
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
                End If
                If IsNull(data_lin.Recordset("linea")) = False Then
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("linea")
                Else
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
                End If
                If IsNull(data_lin.Recordset("cod_cli")) = False Then
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("cod_cli")
                Else
                   lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
                End If
                Xcountt = Xcountt + 1
             End If
             data_lin.Recordset.MoveNext
          Loop
       Else
          MsgBox "No hay registros", vbInformation, "Ver historial"
       End If
       Xcountt = 1
   End If
End If

'        Sumar = Sumar + CDbl(ListView1.ListItems(i).SubItems(1))

End Sub


Private Sub Command2_Click()

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"
data_inf.DatabaseName = App.path & "\informes.mdb"
data_inf.RecordSource = "infvtas"
data_inf.Refresh

frm_ctrlmedb.MousePointer = 11
If mfd.Text <> "__/__/____" Then
   If mfh.Text <> "__/__/____" Then
      If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "AGUILLEN" Then
         data_lin.RecordSource = "Select * from linmmdd where realizada >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and realizada <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And dias =" & 8 & " And nro_flia=" & 6 & " order by fecha,base"
         data_lin.Refresh
      Else
         data_lin.RecordSource = "Select * from linmmdd where realizada >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and realizada <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And dias =" & 8 & " And nro_flia=" & 6 & " and base =" & frm_menu.data_parse.Recordset("base") & " order by cod_prod,fecha"
         data_lin.Refresh
      End If
      If data_lin.Recordset.RecordCount > 0 Then
         data_lin.Recordset.MoveFirst
         Do While Not data_lin.Recordset.EOF
            data_inf.Recordset.AddNew
            data_inf.Recordset("fecha") = data_lin.Recordset("realizada")
            data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
            data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
            data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
            data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
            data_inf.Recordset("base") = data_lin.Recordset("base")
            data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
            data_inf.Recordset.Update
            data_lin.Recordset.MoveNext
         Loop
      End If
      data_inf.RecordSource = "Select * from infvtas"
      data_inf.Refresh
      frm_ctrlmedb.MousePointer = 0
      cr1.ReportFileName = App.path & "\infvtamedb.rpt"
      cr1.ReportTitle = "INFORME DE ENTREGA DE MEDICACION DESDE: " & mfd.Text & " HASTA: " & mfh.Text
      cr1.Action = 1
   
   End If
End If
frm_ctrlmedb.MousePointer = 0

End Sub

Private Sub Command3_Click()

'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"

data_inf.DatabaseName = App.path & "\informes.mdb"
data_inf.RecordSource = "infvtas"
data_inf.Refresh

frm_ctrlmedb.MousePointer = 11
If mfd.Text <> "__/__/____" Then
   If mfh.Text <> "__/__/____" Then
      If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "AGUILLEN" Then
         data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And nro_flia=" & 6 & " order by fecha,base"
         data_lin.Refresh
      Else
         data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' And nro_flia=" & 6 & " and base =" & frm_menu.data_parse.Recordset("base") & " order by cod_prod,fecha"
         data_lin.Refresh
      End If
      If data_lin.Recordset.RecordCount > 0 Then
         data_lin.Recordset.MoveFirst
         Do While Not data_lin.Recordset.EOF
            data_inf.Recordset.AddNew
            data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
            data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
            data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
            data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
            data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
            data_inf.Recordset("base") = data_lin.Recordset("base")
            data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
            data_inf.Recordset.Update
            data_lin.Recordset.MoveNext
         Loop
      End If
      data_inf.RecordSource = "Select * from infvtas"
      data_inf.Refresh
      frm_ctrlmedb.MousePointer = 0
      cr1.ReportFileName = App.path & "\infvtamedb.rpt"
      cr1.ReportTitle = "INFORME DE VENTA MEDICACION DESDE: " & mfd.Text & " HASTA: " & mfh.Text
      cr1.Action = 1
   
   End If
End If
frm_ctrlmedb.MousePointer = 0

End Sub

Private Sub Command6_Click()
Unload Me

End Sub

Private Sub Form_Load()
Dim Xcountt As Long
Dim Xdeff As Date
Xdeff = Date - 30
data_lin.ConnectionString = "dsn=" & Xconexrmt
data_actu.ConnectionString = "dsn=" & Xconexrmt
data_lindos.ConnectionString = "dsn=" & Xconexrmt

Label2.Caption = WElusuario
Label5.Caption = frm_menu.data_parse.Recordset("base")
Xcountt = 1
'data_actu.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_lindos.Connect = "odbc;dsn=" & Xconexrmt & ";"
If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "AGUILLEN" Then
   data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,3,5) order by fecha,base"
   data_lin.Refresh
Else
   data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,3,5) and base =" & frm_menu.data_parse.Recordset("base") & " order by fecha,factura"
   data_lin.Refresh
End If

lis1.ListItems.Clear
If data_lin.Recordset.RecordCount > 0 Then
   data_lin.Recordset.MoveFirst
   Do While Not data_lin.Recordset.EOF
      If data_lin.Recordset("cod_prod") = 60103 Or _
         data_lin.Recordset("cod_prod") = 60105 Or _
         data_lin.Recordset("cod_prod") = 60106 Or _
         data_lin.Recordset("cod_prod") = 60107 Or _
         data_lin.Recordset("cod_prod") = 60108 Or _
         data_lin.Recordset("cod_prod") = 60109 Then
         If IsNull(data_lin.Recordset("fecha")) = False Then
            lis1.ListItems.Add Xcountt, , data_lin.Recordset("fecha")
         Else
            lis1.ListItems.Add Xcountt, , "01/01/2010"
         End If
         If IsNull(data_lin.Recordset("hora")) = False Then
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("hora")
         Else
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
         End If
         If IsNull(data_lin.Recordset("nom_cli")) = False Then
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("nom_cli")
         Else
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "NN"
         End If
         If IsNull(data_lin.Recordset("convenio")) = False Then
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("convenio")
         Else
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "S/C"
         End If
         If IsNull(data_lin.Recordset("nom_medic")) = False Then
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("nom_medic")
         Else
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "NO INGRESADO"
         End If
         If IsNull(data_lin.Recordset("base")) = False Then
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("base")
         Else
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "99"
         End If
         If IsNull(data_lin.Recordset("ced_socio")) = False Then
            If IsNull(data_lin.Recordset("fact")) = False Then
               lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("ced_socio") & "-" & data_lin.Recordset("fact")
            Else
               lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("ced_socio") & "-0"
            End If
         Else
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
         End If
         If IsNull(data_lin.Recordset("factura")) = False Then
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("factura")
         Else
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
         End If
         If IsNull(data_lin.Recordset("linea")) = False Then
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("linea")
         Else
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
         End If
         If IsNull(data_lin.Recordset("cod_cli")) = False Then
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , data_lin.Recordset("cod_cli")
         Else
            lis1.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
         End If
         
         Xcountt = Xcountt + 1
      End If
      data_lin.Recordset.MoveNext
   Loop
Else
   MsgBox "No hay registros", vbInformation, "Ver historial"
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

Private Sub lis1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'MsgBox "El item está cheq"

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

