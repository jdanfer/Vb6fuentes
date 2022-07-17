VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ctrlfarm 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Controles de entrega de medicación"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   420
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
   Icon            =   "frm_ctrlfarm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11265
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pendientes"
      Height          =   495
      Left            =   5640
      Picture         =   "frm_ctrlfarm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Medicación pendiente de retirar que ya está vencida."
      Top             =   6960
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc data_actu 
      Height          =   735
      Left            =   4920
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
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
   Begin VB.CommandButton Command15 
      Caption         =   "Etiqueta1"
      Height          =   375
      Left            =   7680
      TabIndex        =   21
      Top             =   7440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Etiqueta"
      Height          =   375
      Left            =   9840
      TabIndex        =   20
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   4920
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0000FF00&
      Caption         =   "Imprimir Etiqueta"
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6960
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   7560
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Consultas"
      Height          =   495
      Left            =   3840
      Picture         =   "frm_ctrlfarm.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Informes"
      Height          =   495
      Left            =   2040
      Picture         =   "frm_ctrlfarm.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10560
      Picture         =   "frm_ctrlfarm.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   6960
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Actualizar"
      Height          =   495
      Left            =   120
      Picture         =   "frm_ctrlfarm.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Pedidos pasados a farmacia central"
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   10935
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   375
         Left            =   1440
         Top             =   1320
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
      Begin MSAdodcLib.Adodc data_lindos 
         Height          =   495
         Left            =   7560
         Top             =   1200
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Solicitado a mut."
         Height          =   495
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Medicación solicitada a mutualista"
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Devolución..."
         Height          =   495
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Medicación no entregada"
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enviado a base..."
         Height          =   495
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Se envía directamente a la base dónde se solicitó"
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enviado a Sede..."
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Se envía a la sede secundaria correspondiente"
         Top             =   2640
         Width           =   1935
      End
      Begin MSComctlLib.ListView lis2 
         Height          =   2295
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
            Object.Width           =   2291
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
         Height          =   615
         Left            =   4800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1560
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Devolución..."
         Height          =   495
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enviado a base..."
         Height          =   495
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2880
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Pasar a Farm. Central"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Entregado en sede"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2880
         Width           =   2175
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
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   9240
      TabIndex        =   11
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Usuario actual:"
      Height          =   255
      Left            =   7440
      TabIndex        =   10
      Top             =   0
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Picture         =   "frm_ctrlfarm.frx":1FF4
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   1455
   End
End
Attribute VB_Name = "frm_ctrlfarm"
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
Dim Feclist As Date

Xdeff = Date - 31
Dim Xnommedi As String
Xcountt = 1
Xmensajeme = MsgBox("Desea procesar los registros marcados cómo ENTREGADOS?", vbInformation + vbYesNo, "Control")
Xind = 0
DoEvents
frm_ctrlfarm.MousePointer = 11
If Xmensajeme = vbYes Then
   For Xind = 1 To lis1.ListItems.count
       pb1.Visible = True
       pb1.Max = lis1.ListItems.count
       pb1.Value = 0
       lis1.ListItems(Xind).Selected = True
       If lis1.ListItems.Item(lis1.SelectedItem.index).Checked = True Then
'       MsgBox "Chequeado"
          Feclist = CDate(lis1.SelectedItem.Text)
          Xmatme = lis1.SelectedItem.ListSubItems(9).Text
          Xlinme = lis1.SelectedItem.ListSubItems(8).Text
          Xfacme = lis1.SelectedItem.ListSubItems(7).Text
          pb1.Visible = True
          data_actu.RecordSource = "Select * from linmmdd where cod_cli =" & Xmatme & " and factura =" & Xfacme & " and linea =" & Xlinme & " and fecha ='" & Format(Feclist, "yyyy-mm-dd") & "'"
          data_actu.Refresh
          Xnommedi = InputBox("Ingrese descripción de medicamento entregado", "Medicamento entregado")
          If Xnommedi = "" Then
             MsgBox "Debe ingresar nombre de medicamento para procesar, NO SE REALIZA EL PROCESO DE ENTREGA", vbCritical, "SAPP"
          Else
            If data_actu.Recordset.RecordCount > 0 Then
'               data_actu.Recordset.Edit
               If data_actu.Recordset("dias") = 4 Then
                  data_actu.Recordset("dias") = 9
               Else
                  data_actu.Recordset("dias") = 1
               End If
               data_actu.Recordset("univta") = 1 'Para imprimir etiqueta
               data_actu.Recordset("vto") = Format(Date, "yyyy-mm-dd")
               data_actu.Recordset("margen_prd") = Val(Mid(Format(Time, "HH:mm"), 1, 2))
               data_actu.Recordset("pre_prod") = Val(Mid(Format(Time, "HH:mm"), 4, 2))
               If Xnommedi <> "" Then
                  data_actu.Recordset("zona") = Mid(Xnommedi, 1, 25)
               End If
               Xnommedi = ""
               data_actu.Recordset("numero") = Welnrou
               data_actu.Recordset.Update
               data_actu.Refresh
            End If
          End If
       End If
       pb1.Value = pb1.Value + 1
   Next Xind
    Dim Ximplaeti As String
    Ximplaeti = MsgBox("Desea imprimir la etiqueta para registros seleccionados?", vbInformation + vbYesNo, "Etiquetas")
    If Ximplaeti = vbYes Then
       Command15_Click
    Else
       lis1.ListItems.Clear
    End If
   
    If frm_menu.data_parse.Recordset("base") = 18 Or frm_menu.data_parse.Recordset("base") = 92 Then
       data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (1,2,4,18,92,3) and tot_lin >=" & 0 & " order by fecha,factura"
       data_lin.Refresh
    Else
       If frm_menu.data_parse.Recordset("base") = 16 Or frm_menu.data_parse.Recordset("base") = 91 Or WElusuario = "FLORENCIA" Or WElusuario = "SILVIAE" Or WElusuario = "MARIANAT" Or frm_menu.data_parse.Recordset("base") = 93 Then
          data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and tot_lin >=" & 0 & " order by fecha,factura"
    '      data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xdeff, "yyyy/mm/dd") & "# and nro_flia =" & 6 & " and dias in (0,4) and base in (15,16,91) and tot_lin >=" & 0 & " order by fecha,factura"
          data_lin.Refresh
       Else
          If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "LQUINTEROS" Then
             data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and tot_lin >=" & 0 & " order by fecha,factura"
             data_lin.Refresh
          Else
             If frm_menu.data_parse.Recordset("base") = 12 Then
                data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (12) and tot_lin >=" & 0 & " order by fecha,factura"
             Else
                data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (17,93) and tot_lin >=" & 0 & " order by fecha,factura"
             End If
             data_lin.Refresh
          End If
       End If
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
   DoEvents
    Xcountt = 1
End If
frm_ctrlfarm.MousePointer = 0
pb1.Visible = False

'        Sumar = Sumar + CDbl(ListView1.ListItems(i).SubItems(1))

End Sub

Private Sub Command10_Click()
If XWeltipoU = "ADM FARMACIA" Or XWeltipoU = "ADMINISTRADOR" Or XWeltipoU = "USUARIOS FARM" Then
   frm_consfarma.Show vbModal
Else
   MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub Command11_Click()
Dim Xind As Long
Dim Xmatme, Xfacme, Xlinme As Long
Dim Xmensajeme As String
Dim Xcountt As Long
Dim Xdeff As Date
Xdeff = Date - 31

Dim Xnommedi As String
Xcountt = 1
Xmensajeme = MsgBox("Desea procesar los registros marcados cómo SOLICITADO a la MUTUALISTA?", vbInformation + vbYesNo, "Control")
Xind = 0
If Xmensajeme = vbYes Then
   For Xind = 1 To lis2.ListItems.count
       lis2.ListItems(Xind).Selected = True
       If lis2.ListItems.Item(lis2.SelectedItem.index).Checked = True Then
'       MsgBox "Chequeado"
          Xmatme = lis2.SelectedItem.ListSubItems(9).Text
          Xlinme = lis2.SelectedItem.ListSubItems(8).Text
          Xfacme = lis2.SelectedItem.ListSubItems(7).Text
          Xnommedi = InputBox("Ingrese nombre de medicamento ENTREGADO", "Medicamento Entregado")
          data_actu.RecordSource = "Select * from linmmdd where cod_cli =" & Xmatme & " and factura =" & Xfacme & " and linea =" & Xlinme
          data_actu.Refresh
          If Xnommedi = "" Then
             MsgBox "Debe ingresar nombre de medicamento para procesar, NO SE REALIZA EL PROCESO DE ENTREGA", vbCritical, "SAPP"
          Else
            If data_actu.Recordset.RecordCount > 0 Then
'               data_actu.Recordset.Edit
               data_actu.Recordset("dias") = 7
               data_actu.Recordset("vto") = Date
               data_actu.Recordset("margen_prd") = Val(Mid(Format(Time, "HH:mm"), 1, 2))
               data_actu.Recordset("pre_prod") = Val(Mid(Format(Time, "HH:mm"), 4, 2))
               If Xnommedi <> "" Then
                  data_actu.Recordset("zona") = Mid(Xnommedi, 1, 25)
               End If
               data_actu.Recordset("numero") = Welnrou
               Xnommedi = ""
               data_actu.Recordset.Update
               data_actu.Refresh
            End If
          End If
       End If
   Next Xind
   Xcountt = 1
   data_lindos.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and dias =" & 2 & " and tot_lin >=" & 0 & " order by fecha,base"
   data_lindos.Refresh
   lis2.ListItems.Clear
    If data_lindos.Recordset.RecordCount > 0 Then
       data_lindos.Recordset.MoveFirst
       Do While Not data_lindos.Recordset.EOF
          If data_lindos.Recordset("cod_prod") = 60103 Or _
             data_lindos.Recordset("cod_prod") = 60105 Or _
             data_lindos.Recordset("cod_prod") = 60106 Or _
             data_lindos.Recordset("cod_prod") = 60107 Or _
             data_lindos.Recordset("cod_prod") = 60108 Or _
             data_lindos.Recordset("cod_prod") = 60109 Then
             If IsNull(data_lindos.Recordset("fecha")) = False Then
                lis2.ListItems.Add Xcountt, , data_lindos.Recordset("fecha")
             Else
                lis2.ListItems.Add Xcountt, , "01/01/2010"
             End If
             If IsNull(data_lindos.Recordset("hora")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("hora")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
             End If
             If IsNull(data_lindos.Recordset("nom_cli")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("nom_cli")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "NN"
             End If
            If IsNull(data_lindos.Recordset("convenio")) = False Then
               lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("convenio")
            Else
               lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "S/C"
            End If
             If IsNull(data_lindos.Recordset("nom_medic")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("nom_medic")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "NO INGRESADO"
             End If
             If IsNull(data_lindos.Recordset("base")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("base")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "99"
             End If
             If IsNull(data_lindos.Recordset("ced_socio")) = False Then
                If IsNull(data_lindos.Recordset("fact")) = False Then
                   lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("ced_socio") & "-" & data_lindos.Recordset("fact")
                Else
                   lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("ced_socio") & "-0"
                End If
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("factura")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("factura")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("linea")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("linea")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("cod_cli")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("cod_cli")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             Xcountt = Xcountt + 1
          End If
          data_lindos.Recordset.MoveNext
       Loop
    Else
       MsgBox "No hay registros para farmacia central", vbInformation, "Ver historial"
    End If
Else
    If XWeltipoU = "ADM FARMACIA" Or XWeltipoU = "ADMINISTRADOR" Then
       MsgBox "Se mostrará lo solicitado a mutualista que está pendiente", vbInformation
       frm_ctrlmedmut.Show vbModal
    Else
       MsgBox "Usuario no autorizado"
    End If
End If

End Sub

Private Sub Command12_Click()
Dim Xind As Long
Dim Xmatme, Xfacme, Xlinme As Long
Dim Xmensajeme As String
Dim Xcountt As Long
Dim Xdeff As Date
Xdeff = Date - 31
Xcountt = 1
Xmensajeme = MsgBox("Desea procesar los registros marcados cómo DEVOLUCION?", vbInformation + vbYesNo, "Control")
Xind = 0
frm_ctrlfarm.MousePointer = 11

If Xmensajeme = vbYes Then
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
             data_actu.Recordset("dias") = 6
             data_actu.Recordset("vto") = Date
             data_actu.Recordset("margen_prd") = Val(Mid(Format(Time, "HH:mm"), 1, 2))
             data_actu.Recordset("pre_prod") = Val(Mid(Format(Time, "HH:mm"), 4, 2))
             data_actu.Recordset("numero") = Welnrou
             data_actu.Recordset.Update
             data_actu.Refresh
          End If
          
       End If
   Next Xind
   If frm_menu.data_parse.Recordset("base") = 18 Or frm_menu.data_parse.Recordset("base") = 92 Then
      data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (1,2,4,18,92,3) and tot_lin >=" & 0 & " order by fecha,factura"
      data_lin.Refresh
   Else
      If frm_menu.data_parse.Recordset("base") = 16 Or frm_menu.data_parse.Recordset("base") = 91 Or WElusuario = "FLORENCIA" Or WElusuario = "SILVIAE" Or WElusuario = "MARIANAT" Or frm_menu.data_parse.Recordset("base") = 93 Then
         data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and tot_lin >=" & 0 & " order by fecha,factura"
   '      data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xdeff, "yyyy/mm/dd") & "# and nro_flia =" & 6 & " and dias in (0,4) and base in (15,16,91) and tot_lin >=" & 0 & " order by fecha,factura"
         data_lin.Refresh
      Else
         If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "LQUINTEROS" Then
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and tot_lin >=" & 0 & " order by fecha,factura"
            data_lin.Refresh
         Else
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (17,93) and tot_lin >=" & 0 & " order by fecha,factura"
            data_lin.Refresh
         End If
      End If
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
   DoEvents
End If
frm_ctrlfarm.MousePointer = 0

End Sub

Private Sub Command13_Click()
frm_ctrlmedet.Show vbModal

End Sub

Private Sub Command14_Click()
Dim Xind As Long
Dim Xmatme, Xfacme, Xlinme As Long
Dim Xmensajeme As String
Dim Xcountt As Long
Dim Xdeff As Date
Xdeff = Date - 31
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_inf.DatabaseName = App.path & "\informes.mdb"
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"

data_inf.RecordSource = "infvtas"
data_inf.Refresh

If WElusuario = "LQUINTEROS" Or WElusuario = "SILVIAE" Or XWeltipoU = "ADMINISTRADOR" Or WElusuario = "ANALIAG" Or WElusuario = "" Or XWeltipoU = "ADM FARMACIA" Or WElusuario = "FLORENCIA" Or WElusuario = "MARIANAT" Then

    Xcountt = 1
'    Xmensajeme = MsgBox("Desea IMPRIMIR ETIQUETA de los registros seleccionados?", vbInformation + vbYesNo, "Control")
    Xind = 0
'    If Xmensajeme = vbYes Then
       For Xind = 1 To lis2.ListItems.count
           lis2.ListItems(Xind).Selected = True
           If lis2.ListItems.Item(lis2.SelectedItem.index).Checked = True Then
    '       MsgBox "Chequeado"
              data_inf.Recordset.AddNew
              data_inf.Recordset("fecha") = Date
              If lis2.SelectedItem.ListSubItems(5).Text = 92 Then
                 data_inf.Recordset("base") = 18
              Else
                 If lis2.SelectedItem.ListSubItems(5).Text = 91 Then
                    data_inf.Recordset("base") = 16
                 Else
                    data_inf.Recordset("base") = lis2.SelectedItem.ListSubItems(5).Text
                 End If
              End If
              data_inf.Recordset("nom_cli") = lis2.SelectedItem.ListSubItems(2).Text
              data_inf.Recordset("RUC") = lis2.SelectedItem.ListSubItems(6).Text
              
              Xmatme = lis2.SelectedItem.ListSubItems(9).Text
              Xlinme = lis2.SelectedItem.ListSubItems(8).Text
              Xfacme = lis2.SelectedItem.ListSubItems(7).Text
              data_actu.RecordSource = "Select * from linmmdd where cod_cli =" & Xmatme & " and factura =" & Xfacme & " and linea =" & Xlinme
              data_actu.Refresh
              If data_actu.Recordset.RecordCount > 0 Then
                 If IsNull(data_actu.Recordset("zona")) = False Then
                    data_inf.Recordset("zona") = data_actu.Recordset("zona")
                 Else
                    If IsNull(data_actu.Recordset("nom_med_a")) = False Then
                       data_inf.Recordset("zona") = data_actu.Recordset("nom_med_a")
                    Else
                       data_inf.Recordset("zona") = lis2.SelectedItem.ListSubItems(4).Text
                    End If
                 End If
              Else
                 data_inf.Recordset("zona") = lis2.SelectedItem.ListSubItems(4).Text
              End If
              data_inf.Recordset("convenio") = lis2.SelectedItem.ListSubItems(3).Text
              data_inf.Recordset("factura") = lis2.SelectedItem.ListSubItems(7).Text
              data_inf.Recordset.Update
              data_inf.Refresh
              
           End If
       Next Xind
    ''End If
    data_inf.RecordSource = "Select * from infvtas"
    data_inf.Refresh
    If data_inf.Recordset.RecordCount > 0 Then
       data_inf.Recordset.MoveFirst
    End If
    lis2.ListItems.Clear
   cr1.ReportTitle = "ENVÍO DESDE " & frm_menu.data_parse.Recordset("localidad")
   cr1.ReportFileName = App.path & "\infetiquet.rpt"
   cr1.Action = 1
   
Else
    MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub Command15_Click()
Dim Xind As Long
Dim Xmatme, Xfacme, Xlinme As Long
Dim Xmensajeme As String
Dim Xcountt As Long
Dim Xdeff As Date
Dim Feclist As Date
Xdeff = Date - 31
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_inf.DatabaseName = App.path & "\informes.mdb"
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"
data_inf.RecordSource = "infvtas"
data_inf.Refresh

If WElusuario = "LQUINTEROS" Or WElusuario = "SILVIAE" Or XWeltipoU = "ADMINISTRADOR" Or WElusuario = "ANALIAG" Or WElusuario = "ANDREAP" Or XWeltipoU = "ADM FARMACIA" Or WElusuario = "FLORENCIA" Or WElusuario = "MARIANAT" Then

    Xcountt = 1
'    Xmensajeme = MsgBox("Desea IMPRIMIR ETIQUETA de los registros seleccionados?", vbInformation + vbYesNo, "Control")
    Xind = 0
'    If Xmensajeme = vbYes Then
       For Xind = 1 To lis1.ListItems.count
           lis1.ListItems(Xind).Selected = True
           If lis1.ListItems.Item(lis1.SelectedItem.index).Checked = True Then
    '       MsgBox "Chequeado"
              data_inf.Recordset.AddNew
              data_inf.Recordset("fecha") = Date
              If lis1.SelectedItem.ListSubItems(5).Text = 92 Then
                 data_inf.Recordset("base") = 18
              Else
                 If lis1.SelectedItem.ListSubItems(5).Text = 91 Then
                    data_inf.Recordset("base") = 16
                 Else
                    data_inf.Recordset("base") = lis1.SelectedItem.ListSubItems(5).Text
                 End If
              End If
              data_inf.Recordset("nom_cli") = lis1.SelectedItem.ListSubItems(2).Text
              data_inf.Recordset("RUC") = lis1.SelectedItem.ListSubItems(6).Text
              
              Xmatme = lis1.SelectedItem.ListSubItems(9).Text
              Xlinme = lis1.SelectedItem.ListSubItems(8).Text
              Xfacme = lis1.SelectedItem.ListSubItems(7).Text
              Feclist = CDate(lis1.SelectedItem.Text)
              data_actu.RecordSource = "Select * from linmmdd where cod_cli =" & Xmatme & " and factura =" & Xfacme & " and linea =" & Xlinme & " and fecha ='" & Format(Feclist, "yyyy-mm-dd") & "'"
              data_actu.Refresh
              If data_actu.Recordset.RecordCount > 0 Then
                 If IsNull(data_actu.Recordset("zona")) = False Then
                    data_inf.Recordset("zona") = data_actu.Recordset("zona")
                 Else
                    If IsNull(data_actu.Recordset("nom_med_a")) = False Then
                       data_inf.Recordset("zona") = data_actu.Recordset("nom_med_a")
                    Else
                       data_inf.Recordset("zona") = lis1.SelectedItem.ListSubItems(4).Text
                    End If
                 End If
              Else
                 data_inf.Recordset("zona") = lis1.SelectedItem.ListSubItems(4).Text
              End If
              data_inf.Recordset("convenio") = lis1.SelectedItem.ListSubItems(3).Text
              data_inf.Recordset("factura") = lis1.SelectedItem.ListSubItems(7).Text
              data_inf.Recordset.Update
              data_inf.Refresh
              
           End If
       Next Xind
    ''End If
    data_inf.RecordSource = "Select * from infvtas"
    data_inf.Refresh
    If data_inf.Recordset.RecordCount > 0 Then
       data_inf.Recordset.MoveFirst
    End If
    lis1.ListItems.Clear
   cr1.ReportTitle = "ENVÍO DESDE " & frm_menu.data_parse.Recordset("localidad")
   cr1.ReportFileName = App.path & "\infetiquet.rpt"
   cr1.Action = 1
   
Else
    MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub Command16_Click()
frm_medpendv.Show vbModal

End Sub

Private Sub Command2_Click()
Dim Xind As Long
Dim Xmatme, Xfacme, Xlinme As Long
Dim Xmensajeme As String
Dim Xcountt As Long
Dim Xdeff As Date
Xdeff = Date - 31

Xcountt = 1
Xmensajeme = MsgBox("Desea procesar los registros marcados a FARMACIA CENTRAL?", vbInformation + vbYesNo, "Control")
Xind = 0
If Xmensajeme = vbYes Then
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
             data_actu.Recordset("dias") = 2
             data_actu.Recordset("vto") = Date
             data_actu.Recordset("margen_prd") = Val(Mid(Format(Time, "HH:mm"), 1, 2))
             data_actu.Recordset("pre_prod") = Val(Mid(Format(Time, "HH:mm"), 4, 2))
             data_actu.Recordset("numero") = Welnrou
             data_actu.Recordset("univta") = 2 'Para imprimir etiqueta pendiente
             data_actu.Recordset.Update
             data_actu.Refresh
          End If
          
       End If
   Next Xind
   If frm_menu.data_parse.Recordset("base") = 18 Or frm_menu.data_parse.Recordset("base") = 92 Then
      data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (1,2,4,18,92,3) and tot_lin >=" & 0 & " order by fecha,factura"
      data_lin.Refresh
   Else
      If frm_menu.data_parse.Recordset("base") = 16 Or frm_menu.data_parse.Recordset("base") = 91 Or WElusuario = "FLORENCIA" Or WElusuario = "SILVIAE" Or WElusuario = "MARIANAT" Or frm_menu.data_parse.Recordset("base") = 93 Then
         data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and tot_lin >=" & 0 & " order by fecha,factura"
   '      data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xdeff, "yyyy/mm/dd") & "# and nro_flia =" & 6 & " and dias in (0,4) and base in (15,16,91) and tot_lin >=" & 0 & " order by fecha,factura"
         data_lin.Refresh
      Else
         If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "LQUINTEROS" Then
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and tot_lin >=" & 0 & " order by fecha,factura"
            data_lin.Refresh
         Else
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (17,93) and tot_lin >=" & 0 & " order by fecha,factura"
            data_lin.Refresh
         End If
      End If
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
   data_lindos.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and dias =" & 2 & " and tot_lin >=" & 0 & " order by fecha,base"
   data_lindos.Refresh
    lis2.ListItems.Clear
    If data_lindos.Recordset.RecordCount > 0 Then
       data_lindos.Recordset.MoveFirst
       Do While Not data_lindos.Recordset.EOF
          If data_lindos.Recordset("cod_prod") = 60103 Or _
             data_lindos.Recordset("cod_prod") = 60105 Or _
             data_lindos.Recordset("cod_prod") = 60106 Or _
             data_lindos.Recordset("cod_prod") = 60107 Or _
             data_lindos.Recordset("cod_prod") = 60108 Or _
             data_lindos.Recordset("cod_prod") = 60109 Then
             If IsNull(data_lindos.Recordset("fecha")) = False Then
                lis2.ListItems.Add Xcountt, , data_lindos.Recordset("fecha")
             Else
                lis2.ListItems.Add Xcountt, , "01/01/2010"
             End If
             If IsNull(data_lindos.Recordset("hora")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("hora")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
             End If
             If IsNull(data_lindos.Recordset("nom_cli")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("nom_cli")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "NN"
             End If
             If IsNull(data_lindos.Recordset("convenio")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("convenio")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "S/C"
             End If
             If IsNull(data_lindos.Recordset("nom_medic")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("nom_medic")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "NO INGRESADO"
             End If
             If IsNull(data_lindos.Recordset("base")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("base")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "99"
             End If
             If IsNull(data_lindos.Recordset("ced_socio")) = False Then
                If IsNull(data_lindos.Recordset("fact")) = False Then
                   lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("ced_socio") & "-" & data_lindos.Recordset("fact")
                Else
                   lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("ced_socio") & "-0"
                End If
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("factura")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("factura")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("linea")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("linea")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("cod_cli")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("cod_cli")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             
             Xcountt = Xcountt + 1
          End If
          data_lindos.Recordset.MoveNext
       Loop
    Else
       MsgBox "No hay registros para farmacia central", vbInformation, "Ver historial"
    End If
End If

End Sub

Private Sub Command3_Click()
Dim Xind As Long
Dim Xmatme, Xfacme, Xlinme As Long
Dim Xmensajeme As String
Dim Xcountt As Long
Dim Xdeff As Date
Xdeff = Date - 31
Dim Xnommedi As String
Xcountt = 1
Xmensajeme = MsgBox("Desea procesar los registros marcados a SEDE...?", vbInformation + vbYesNo, "Control")
Xind = 0
If Xmensajeme = vbYes Then
   For Xind = 1 To lis2.ListItems.count
       lis2.ListItems(Xind).Selected = True
       If lis2.ListItems.Item(lis2.SelectedItem.index).Checked = True Then
'       MsgBox "Chequeado"
          Xmatme = lis2.SelectedItem.ListSubItems(9).Text
          Xlinme = lis2.SelectedItem.ListSubItems(8).Text
          Xfacme = lis2.SelectedItem.ListSubItems(7).Text
          Xnommedi = InputBox("Ingrese nombre de medicamento ENTREGADO", "Medicamento entregado")
          data_actu.RecordSource = "Select * from linmmdd where cod_cli =" & Xmatme & " and factura =" & Xfacme & " and linea =" & Xlinme
          data_actu.Refresh
          If Xnommedi = "" Then
             MsgBox "Debe ingresar nombre de medicamento para procesar, NO SE REALIZA EL PROCESO DE ENTREGA", vbCritical, "SAPP"
          Else
            If data_actu.Recordset.RecordCount > 0 Then
'               data_actu.Recordset.Edit
               data_actu.Recordset("dias") = 4
               data_actu.Recordset("vto") = Date
               data_actu.Recordset("margen_prd") = Val(Mid(Format(Time, "HH:mm"), 1, 2))
               data_actu.Recordset("pre_prod") = Val(Mid(Format(Time, "HH:mm"), 4, 2))
               If Xnommedi <> "" Then
                  data_actu.Recordset("zona") = Mid(Xnommedi, 1, 25)
               End If
               data_actu.Recordset("numero") = Welnrou
               If IsNull(data_actu.Recordset("univta")) = False Then
                  If data_actu.Recordset("univta") = 1 Then
                  Else
                     data_actu.Recordset("univta") = 2 'Para imprimir etiqueta
                  End If
               Else
                  data_actu.Recordset("univta") = 2 'Para imprimir etiqueta
               End If
               Xnommedi = ""
               data_actu.Recordset.Update
               data_actu.Refresh
            End If
          End If
       End If
   Next Xind
   If frm_menu.data_parse.Recordset("base") = 18 Or frm_menu.data_parse.Recordset("base") = 92 Then
      data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (1,2,4,18,92,3) and tot_lin >=" & 0 & " order by fecha,factura"
      data_lin.Refresh
   Else
      If frm_menu.data_parse.Recordset("base") = 16 Or frm_menu.data_parse.Recordset("base") = 91 Or WElusuario = "FLORENCIA" Or WElusuario = "SILVIAE" Or WElusuario = "MARIANAT" Or frm_menu.data_parse.Recordset("base") = 93 Then
         data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and tot_lin >=" & 0 & " order by fecha,factura"
   '      data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xdeff, "yyyy/mm/dd") & "# and nro_flia =" & 6 & " and dias in (0,4) and base in (15,16,91) and tot_lin >=" & 0 & " order by fecha,factura"
         data_lin.Refresh
      Else
         If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "LQUINTEROS" Then
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and tot_lin >=" & 0 & " order by fecha,factura"
            data_lin.Refresh
         Else
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (17,93) and tot_lin >=" & 0 & " order by fecha,factura"
            data_lin.Refresh
         End If
      End If
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
    Dim Ximplaeti As String
    Ximplaeti = MsgBox("Desea imprimir la etiqueta para registros seleccionados?", vbInformation + vbYesNo, "Etiquetas")
    If Ximplaeti = vbYes Then
       Command14_Click
    Else
       lis2.ListItems.Clear
    End If
   
   data_lindos.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and dias =" & 2 & " and tot_lin >=" & 0 & " order by fecha,base"
   data_lindos.Refresh
'   lis2.ListItems.Clear
    If data_lindos.Recordset.RecordCount > 0 Then
       data_lindos.Recordset.MoveFirst
       Do While Not data_lindos.Recordset.EOF
          If data_lindos.Recordset("cod_prod") = 60103 Or _
             data_lindos.Recordset("cod_prod") = 60105 Or _
             data_lindos.Recordset("cod_prod") = 60106 Or _
             data_lindos.Recordset("cod_prod") = 60107 Or _
             data_lindos.Recordset("cod_prod") = 60108 Or _
             data_lindos.Recordset("cod_prod") = 60109 Then
             If IsNull(data_lindos.Recordset("fecha")) = False Then
                lis2.ListItems.Add Xcountt, , data_lindos.Recordset("fecha")
             Else
                lis2.ListItems.Add Xcountt, , "01/01/2010"
             End If
             If IsNull(data_lindos.Recordset("hora")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("hora")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
             End If
             If IsNull(data_lindos.Recordset("nom_cli")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("nom_cli")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "NN"
             End If
             If IsNull(data_lindos.Recordset("convenio")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("convenio")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "S/C"
             End If
             If IsNull(data_lindos.Recordset("nom_medic")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("nom_medic")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "NO INGRESADO"
             End If
             If IsNull(data_lindos.Recordset("base")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("base")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "99"
             End If
             If IsNull(data_lindos.Recordset("ced_socio")) = False Then
                If IsNull(data_lindos.Recordset("fact")) = False Then
                   lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("ced_socio") & "-" & data_lindos.Recordset("fact")
                Else
                   lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("ced_socio") & "-0"
                End If
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("factura")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("factura")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("linea")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("linea")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("cod_cli")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("cod_cli")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             
             Xcountt = Xcountt + 1
          End If
          data_lindos.Recordset.MoveNext
       Loop
    
    Else
       MsgBox "No hay registros para farmacia central", vbInformation, "Ver historial"
    End If
End If

End Sub

Private Sub Command4_Click()
Dim Xind As Long
Dim Xmatme, Xfacme, Xlinme As Long
Dim Xmensajeme As String
Dim Xcountt As Long
Dim Xdeff As Date
Xdeff = Date - 31
Dim Xnommedi As String
Xcountt = 1
Xmensajeme = MsgBox("Desea procesar los registros marcados COMO ENVIADO A BASE..?", vbInformation + vbYesNo, "Control")
Xind = 0
If Xmensajeme = vbYes Then
   For Xind = 1 To lis1.ListItems.count
       lis1.ListItems(Xind).Selected = True
       If lis1.ListItems.Item(lis1.SelectedItem.index).Checked = True Then
'       MsgBox "Chequeado"
          Xmatme = lis1.SelectedItem.ListSubItems(9).Text
          Xlinme = lis1.SelectedItem.ListSubItems(8).Text
          Xfacme = lis1.SelectedItem.ListSubItems(7).Text
          Xnommedi = InputBox("Ingrese nombre de medicamento ENTREGADO", "Medicamento entregado")
          data_actu.RecordSource = "Select * from linmmdd where cod_cli =" & Xmatme & " and factura =" & Xfacme & " and linea =" & Xlinme
          data_actu.Refresh
          If Xnommedi = "" Then
             MsgBox "Debe ingresar nombre de medicamento para procesar, NO SE REALIZA EL PROCESO DE ENTREGA", vbCritical, "SAPP"
          Else
            If data_actu.Recordset.RecordCount > 0 Then
'               data_actu.Recordset.Edit
               If data_actu.Recordset("dias") = 4 Then
                  data_actu.Recordset("dias") = 10
               Else
                  data_actu.Recordset("dias") = 3
               End If
               data_actu.Recordset("vto") = Date
               data_actu.Recordset("margen_prd") = Val(Mid(Format(Time, "HH:mm"), 1, 2))
               data_actu.Recordset("pre_prod") = Val(Mid(Format(Time, "HH:mm"), 4, 2))
               If Xnommedi <> "" Then
                  data_actu.Recordset("zona") = Mid(Xnommedi, 1, 25)
               End If
               data_actu.Recordset("numero") = Welnrou
               Xnommedi = ""
               If IsNull(data_actu.Recordset("univta")) = False Then
                  If data_actu.Recordset("univta") = 1 Then
                  Else
                     data_actu.Recordset("univta") = 2 'Para imprimir etiqueta
                  End If
               Else
                  data_actu.Recordset("univta") = 2 'Para imprimir etiqueta
               End If
               data_actu.Recordset.Update
               data_actu.Refresh
            End If
          End If
       End If
   Next Xind
    Dim Ximplaeti As String
    Ximplaeti = MsgBox("Desea imprimir la etiqueta para registros seleccionados?", vbInformation + vbYesNo, "Etiquetas")
    If Ximplaeti = vbYes Then
       Command15_Click
    Else
       lis1.ListItems.Clear
    End If
   If frm_menu.data_parse.Recordset("base") = 18 Or frm_menu.data_parse.Recordset("base") = 92 Then
      data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (1,2,4,18,92,3) and tot_lin >=" & 0 & " order by fecha,factura"
      data_lin.Refresh
   Else
      If frm_menu.data_parse.Recordset("base") = 16 Or frm_menu.data_parse.Recordset("base") = 91 Or WElusuario = "FLORENCIA" Or WElusuario = "SILVIAE" Or WElusuario = "MARIANAT" Or frm_menu.data_parse.Recordset("base") = 93 Then
         data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and tot_lin >=" & 0 & " order by fecha,factura"
   '      data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xdeff, "yyyy/mm/dd") & "# and nro_flia =" & 6 & " and dias in (0,4) and base in (15,16,91) and tot_lin >=" & 0 & " order by fecha,factura"
         data_lin.Refresh
      Else
         If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "LQUINTEROS" Then
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and tot_lin >=" & 0 & " order by fecha,factura"
            data_lin.Refresh
         Else
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (17,93) and tot_lin >=" & 0 & " order by fecha,factura"
            data_lin.Refresh
         End If
      End If
   End If
'   lis1.ListItems.Clear
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

End Sub

Private Sub Command5_Click()
Dim Xind As Long
Dim Xmatme, Xfacme, Xlinme As Long
Dim Xmensajeme As String
Dim Xcountt As Long
Dim Xdeff As Date
Xdeff = Date - 31

Xcountt = 1
   If frm_menu.data_parse.Recordset("base") = 18 Or frm_menu.data_parse.Recordset("base") = 92 Then
      data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (1,2,4,18,92,3) and tot_lin >=" & 0 & " order by fecha,factura"
      data_lin.Refresh
   Else
      If frm_menu.data_parse.Recordset("base") = 16 Or frm_menu.data_parse.Recordset("base") = 91 Or WElusuario = "FLORENCIA" Or WElusuario = "SILVIAE" Or WElusuario = "MARIANAT" Or frm_menu.data_parse.Recordset("base") = 93 Then
         data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and tot_lin >=" & 0 & " order by fecha,factura"
   '      data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xdeff, "yyyy/mm/dd") & "# and nro_flia =" & 6 & " and dias in (0,4) and base in (15,16,91) and tot_lin >=" & 0 & " order by fecha,factura"
         data_lin.Refresh
      Else
         If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "LQUINTEROS" Then
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and tot_lin >=" & 0 & " order by fecha,factura"
            data_lin.Refresh
         Else
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (17,93) and tot_lin >=" & 0 & " order by fecha,factura"
            data_lin.Refresh
         End If
      End If
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
    data_lindos.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and dias =" & 2 & " and tot_lin >=" & 0 & " order by fecha,base"
    data_lindos.Refresh
    lis2.ListItems.Clear
    If data_lindos.Recordset.RecordCount > 0 Then
       data_lindos.Recordset.MoveFirst
       Do While Not data_lindos.Recordset.EOF
          If data_lindos.Recordset("cod_prod") = 60103 Or _
             data_lindos.Recordset("cod_prod") = 60105 Or _
             data_lindos.Recordset("cod_prod") = 60106 Or _
             data_lindos.Recordset("cod_prod") = 60107 Or _
             data_lindos.Recordset("cod_prod") = 60108 Or _
             data_lindos.Recordset("cod_prod") = 60109 Then
             If IsNull(data_lindos.Recordset("fecha")) = False Then
                lis2.ListItems.Add Xcountt, , data_lindos.Recordset("fecha")
             Else
                lis2.ListItems.Add Xcountt, , "01/01/2010"
             End If
             If IsNull(data_lindos.Recordset("hora")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("hora")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
             End If
             If IsNull(data_lindos.Recordset("nom_cli")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("nom_cli")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "NN"
             End If
             If IsNull(data_lindos.Recordset("convenio")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("convenio")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "S/C"
             End If
             If IsNull(data_lindos.Recordset("nom_medic")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("nom_medic")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "NO INGRESADO"
             End If
             If IsNull(data_lindos.Recordset("base")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("base")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "99"
             End If
             If IsNull(data_lindos.Recordset("ced_socio")) = False Then
                If IsNull(data_lindos.Recordset("fact")) = False Then
                   lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("ced_socio") & "-" & data_lindos.Recordset("fact")
                Else
                   lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("ced_socio") & "-0"
                End If
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("factura")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("factura")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("linea")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("linea")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("cod_cli")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("cod_cli")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             
             Xcountt = Xcountt + 1
          End If
          data_lindos.Recordset.MoveNext
       Loop
    Else
       MsgBox "No hay registros para farmacia central", vbInformation, "Ver historial"
    End If

End Sub

Private Sub Command6_Click()
Unload Me

End Sub

Private Sub Command7_Click()
Dim Xind As Long
Dim Xmatme, Xfacme, Xlinme As Long
Dim Xmensajeme As String
Dim Xcountt As Long
Dim Xdeff As Date
Xdeff = Date - 31
Dim Xnommedi As String
Xcountt = 1
Xmensajeme = MsgBox("Desea procesar los registros marcados a BASE...?", vbInformation + vbYesNo, "Control")
Xind = 0
If Xmensajeme = vbYes Then
   For Xind = 1 To lis2.ListItems.count
       lis2.ListItems(Xind).Selected = True
       If lis2.ListItems.Item(lis2.SelectedItem.index).Checked = True Then
'       MsgBox "Chequeado"
          Xmatme = lis2.SelectedItem.ListSubItems(9).Text
          Xlinme = lis2.SelectedItem.ListSubItems(8).Text
          Xfacme = lis2.SelectedItem.ListSubItems(7).Text
          Xnommedi = InputBox("Ingrese nombre de medicamento ENTREGADO", "Medicamento Entregado")
          data_actu.RecordSource = "Select * from linmmdd where cod_cli =" & Xmatme & " and factura =" & Xfacme & " and linea =" & Xlinme
          data_actu.Refresh
          If Xnommedi = "" Then
             MsgBox "Debe ingresar nombre de medicamento para procesar, NO SE REALIZA EL PROCESO DE ENTREGA", vbCritical, "SAPP"
          Else
            If data_actu.Recordset.RecordCount > 0 Then
'               data_actu.Recordset.Edit
               data_actu.Recordset("dias") = 5
               data_actu.Recordset("vto") = Date
               data_actu.Recordset("margen_prd") = Val(Mid(Format(Time, "HH:mm"), 1, 2))
               data_actu.Recordset("pre_prod") = Val(Mid(Format(Time, "HH:mm"), 4, 2))
               If Xnommedi <> "" Then
                  data_actu.Recordset("zona") = Mid(Xnommedi, 1, 25)
               End If
               data_actu.Recordset("numero") = Welnrou
               If IsNull(data_actu.Recordset("univta")) = False Then
                  If data_actu.Recordset("univta") = 1 Then
                  Else
                     data_actu.Recordset("univta") = 2 'Para imprimir etiqueta
                  End If
               Else
                  data_actu.Recordset("univta") = 2 'Para imprimir etiqueta
               End If
               Xnommedi = ""
               data_actu.Recordset.Update
               data_actu.Refresh
            End If
          End If
       End If
   Next Xind
   Xcountt = 1
   Dim Ximplaeti As String
   Ximplaeti = MsgBox("Desea imprimir la etiqueta para registros seleccionados?", vbInformation + vbYesNo, "Etiquetas")
   If Ximplaeti = vbYes Then
      Command14_Click
   Else
      lis2.ListItems.Clear
   End If
   data_lindos.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and dias =" & 2 & " and tot_lin >=" & 0 & " order by fecha,base"
   data_lindos.Refresh

'   lis2.ListItems.Clear
    If data_lindos.Recordset.RecordCount > 0 Then
       data_lindos.Recordset.MoveFirst
       Do While Not data_lindos.Recordset.EOF
          If data_lindos.Recordset("cod_prod") = 60103 Or _
             data_lindos.Recordset("cod_prod") = 60105 Or _
             data_lindos.Recordset("cod_prod") = 60106 Or _
             data_lindos.Recordset("cod_prod") = 60107 Or _
             data_lindos.Recordset("cod_prod") = 60108 Or _
             data_lindos.Recordset("cod_prod") = 60109 Then
             If IsNull(data_lindos.Recordset("fecha")) = False Then
                lis2.ListItems.Add Xcountt, , data_lindos.Recordset("fecha")
             Else
                lis2.ListItems.Add Xcountt, , "01/01/2010"
             End If
             If IsNull(data_lindos.Recordset("hora")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("hora")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
             End If
             If IsNull(data_lindos.Recordset("nom_cli")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("nom_cli")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "NN"
             End If
             If IsNull(data_lindos.Recordset("convenio")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("convenio")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "S/C"
             End If
             If IsNull(data_lindos.Recordset("nom_medic")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("nom_medic")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "NO INGRESADO"
             End If
             If IsNull(data_lindos.Recordset("base")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("base")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "99"
             End If
             If IsNull(data_lindos.Recordset("ced_socio")) = False Then
                If IsNull(data_lindos.Recordset("fact")) = False Then
                   lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("ced_socio") & "-" & data_lindos.Recordset("fact")
                Else
                   lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("ced_socio") & "-0"
                End If
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("factura")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("factura")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("linea")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("linea")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("cod_cli")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("cod_cli")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             
             Xcountt = Xcountt + 1
          End If
          data_lindos.Recordset.MoveNext
       Loop
    
    Else
       MsgBox "No hay registros para farmacia central", vbInformation, "Ver historial"
    End If
End If

End Sub

Private Sub Command8_Click()
If XWeltipoU = "ADM FARMACIA" Or XWeltipoU = "ADMINISTRADOR" Then
   frm_infctrolfar.Show vbModal
Else
   MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub Command9_Click()
Dim Xind As Long
Dim Xmatme, Xfacme, Xlinme As Long
Dim Xmensajeme As String
Dim Xcountt As Long
Dim Xdeff As Date
Xdeff = Date - 31

Xcountt = 1
Xmensajeme = MsgBox("Desea procesar los registros marcados cómo DEVOLUCION?", vbInformation + vbYesNo, "Control")
Xind = 0
If Xmensajeme = vbYes Then
   For Xind = 1 To lis2.ListItems.count
       lis2.ListItems(Xind).Selected = True
       If lis2.ListItems.Item(lis2.SelectedItem.index).Checked = True Then
'       MsgBox "Chequeado"
          Xmatme = lis2.SelectedItem.ListSubItems(9).Text
          Xlinme = lis2.SelectedItem.ListSubItems(8).Text
          Xfacme = lis2.SelectedItem.ListSubItems(7).Text
          data_actu.RecordSource = "Select * from linmmdd where cod_cli =" & Xmatme & " and factura =" & Xfacme & " and linea =" & Xlinme
          data_actu.Refresh
          If data_actu.Recordset.RecordCount > 0 Then
'             data_actu.Recordset.Edit
             data_actu.Recordset("dias") = 6
             data_actu.Recordset("vto") = Date
             data_actu.Recordset("margen_prd") = Val(Mid(Format(Time, "HH:mm"), 1, 2))
             data_actu.Recordset("pre_prod") = Val(Mid(Format(Time, "HH:mm"), 4, 2))
             data_actu.Recordset("numero") = Welnrou
             data_actu.Recordset.Update
             data_actu.Refresh
          End If
          
       End If
   Next Xind
   Xcountt = 1
   data_lindos.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and dias =" & 2 & " and tot_lin >=" & 0 & " order by fecha,base"
   data_lindos.Refresh
   lis2.ListItems.Clear
    If data_lindos.Recordset.RecordCount > 0 Then
       data_lindos.Recordset.MoveFirst
       Do While Not data_lindos.Recordset.EOF
          If data_lindos.Recordset("cod_prod") = 60103 Or _
             data_lindos.Recordset("cod_prod") = 60105 Or _
             data_lindos.Recordset("cod_prod") = 60106 Or _
             data_lindos.Recordset("cod_prod") = 60107 Or _
             data_lindos.Recordset("cod_prod") = 60108 Or _
             data_lindos.Recordset("cod_prod") = 60109 Then
             If IsNull(data_lindos.Recordset("fecha")) = False Then
                lis2.ListItems.Add Xcountt, , data_lindos.Recordset("fecha")
             Else
                lis2.ListItems.Add Xcountt, , "01/01/2010"
             End If
             If IsNull(data_lindos.Recordset("hora")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("hora")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
             End If
             If IsNull(data_lindos.Recordset("nom_cli")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("nom_cli")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "NN"
             End If
            If IsNull(data_lindos.Recordset("convenio")) = False Then
               lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("convenio")
            Else
               lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "S/C"
            End If
             If IsNull(data_lindos.Recordset("nom_medic")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("nom_medic")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "NO INGRESADO"
             End If
             If IsNull(data_lindos.Recordset("base")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("base")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "99"
             End If
             If IsNull(data_lindos.Recordset("ced_socio")) = False Then
                If IsNull(data_lindos.Recordset("fact")) = False Then
                   lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("ced_socio") & "-" & data_lindos.Recordset("fact")
                Else
                   lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("ced_socio") & "-0"
                End If
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("factura")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("factura")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("linea")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("linea")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             If IsNull(data_lindos.Recordset("cod_cli")) = False Then
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("cod_cli")
             Else
                lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
             End If
             Xcountt = Xcountt + 1
          End If
          data_lindos.Recordset.MoveNext
       Loop
    Else
       MsgBox "No hay registros para farmacia central", vbInformation, "Ver historial"
    End If
End If

End Sub

Private Sub Form_Load()
Dim Xcountt As Long
Dim Xdeff As Date
Xdeff = Date - 31
data_lindos.ConnectionString = "dsn=" & Xconexrmt
data_actu.ConnectionString = "dsn=" & Xconexrmt
data_lin.ConnectionString = "dsn=" & Xconexrmt

Label2.Caption = WElusuario
Xcountt = 1
If frm_menu.data_parse.Recordset("base") = 16 Or WElusuario = "LQUINTEROS" Or WElusuario = "COMPUTOS" Or frm_menu.data_parse.Recordset("base") = 91 Or WElusuario = "FLORENCIA" Or WElusuario = "SILVIAE" Then
   Frame2.Enabled = True
Else
   Frame2.Enabled = False
End If

'data_actu.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_lindos.Connect = "odbc;dsn=" & Xconexrmt & ";"
If frm_menu.data_parse.Recordset("base") = 18 Or frm_menu.data_parse.Recordset("base") = 92 Then
   data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (1,2,4,18,92,3) and tot_lin >=" & 0 & " order by fecha,factura"
   data_lin.Refresh
Else
   If frm_menu.data_parse.Recordset("base") = 16 Or frm_menu.data_parse.Recordset("base") = 91 Or WElusuario = "FLORENCIA" Or WElusuario = "SILVIAE" Or WElusuario = "MARIANAT" Or frm_menu.data_parse.Recordset("base") = 93 Then
      data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and tot_lin >=" & 0 & " order by fecha,factura"
'      data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xdeff, "yyyy/mm/dd") & "# and nro_flia =" & 6 & " and dias in (0,4) and base in (15,16,91) and tot_lin >=" & 0 & " order by fecha,factura"
      data_lin.Refresh
   Else
      If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "LQUINTEROS" Then
         data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and tot_lin >=" & 0 & " order by fecha,factura"
         data_lin.Refresh
      Else
         If frm_menu.data_parse.Recordset("base") = 12 Then
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (12) and tot_lin >=" & 0 & " order by fecha,factura"
         Else
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and dias in (0,4) and base in (17,93) and tot_lin >=" & 0 & " order by fecha,factura"
         End If
         data_lin.Refresh
      End If
   End If
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
'         If data_lin.Recordset("dias") = 4 Then
'            lis1.ListItems.Item(a).Bold = True
'            lis1.ListItems.Item(a).ForeColor = &HFF&
'         End If
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
data_lindos.RecordSource = "Select * from linmmdd where fecha >='" & Format(Xdeff, "yyyy-mm-dd") & "' and dias =" & 2 & " and tot_lin >=" & 0 & " order by fecha,factura"
data_lindos.Refresh
lis2.ListItems.Clear
If data_lindos.Recordset.RecordCount > 0 Then
   data_lindos.Recordset.MoveFirst
   Do While Not data_lindos.Recordset.EOF
      If data_lindos.Recordset("cod_prod") = 60103 Or _
         data_lindos.Recordset("cod_prod") = 60105 Or _
         data_lindos.Recordset("cod_prod") = 60106 Or _
         data_lindos.Recordset("cod_prod") = 60107 Or _
         data_lindos.Recordset("cod_prod") = 60108 Or _
         data_lindos.Recordset("cod_prod") = 60109 Then
         If IsNull(data_lindos.Recordset("fecha")) = False Then
            lis2.ListItems.Add Xcountt, , data_lindos.Recordset("fecha")
         Else
            lis2.ListItems.Add Xcountt, , "01/01/2010"
         End If
         If IsNull(data_lindos.Recordset("hora")) = False Then
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("hora")
         Else
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
         End If
         If IsNull(data_lindos.Recordset("nom_cli")) = False Then
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("nom_cli")
         Else
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "NN"
         End If
         If IsNull(data_lindos.Recordset("convenio")) = False Then
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("convenio")
         Else
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "S/C"
         End If
         If IsNull(data_lindos.Recordset("nom_medic")) = False Then
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("nom_medic")
         Else
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "NO INGRESADO"
         End If
         If IsNull(data_lindos.Recordset("base")) = False Then
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("base")
         Else
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "99"
         End If
         If IsNull(data_lindos.Recordset("ced_socio")) = False Then
            If IsNull(data_lindos.Recordset("fact")) = False Then
               lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("ced_socio") & "-" & data_lindos.Recordset("fact")
            Else
               lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("ced_socio") & "-0"
            End If
         Else
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
         End If
         If IsNull(data_lindos.Recordset("factura")) = False Then
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("factura")
         Else
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
         End If
         If IsNull(data_lindos.Recordset("linea")) = False Then
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("linea")
         Else
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
         End If
         If IsNull(data_lindos.Recordset("cod_cli")) = False Then
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , data_lindos.Recordset("cod_cli")
         Else
            lis2.ListItems.Item(Xcountt).ListSubItems.Add , , "0"
         End If
         
         Xcountt = Xcountt + 1
      End If
      data_lindos.Recordset.MoveNext
   Loop
Else
   MsgBox "No hay registros para farmacia central", vbInformation, "Ver historial"
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
