VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_matvence 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vencimientos"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10275
   Icon            =   "frm_matvence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   10275
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr1 
      Left            =   4440
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_cons 
      Caption         =   "data_cons"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_item 
      Caption         =   "data_item"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_vencim 
      Caption         =   "data_vencim"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox t_busca 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5400
      TabIndex        =   24
      ToolTipText     =   "Digite la descripción del ITEM"
      Top             =   2880
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_matvence.frx":058A
      Height          =   2295
      Left            =   240
      OleObjectBlob   =   "frm_matvence.frx":05A4
      TabIndex        =   22
      ToolTipText     =   "Haga doble click para seleccionar un registro"
      Top             =   3240
      Width           =   9735
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      Picture         =   "frm_matvence.frx":1973
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Imprime los registros en un rango de fecha seleccionado y por base o todos"
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      Picture         =   "frm_matvence.frx":1EFD
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cancelar grabado de datos"
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      Picture         =   "frm_matvence.frx":2487
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Guardar datos"
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton b_edita 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      Picture         =   "frm_matvence.frx":2A11
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Modificar el registro seleccionado"
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      Picture         =   "frm_matvence.frx":2F9B
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Crear nuevo registro"
      Top             =   2640
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos para el registro de vencimientos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.TextBox t_id 
         Height          =   285
         Left            =   600
         TabIndex        =   25
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mvence 
         Height          =   375
         Left            =   8160
         TabIndex        =   16
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox t_movil 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   14
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox t_cant 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   2040
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   960
         Width           =   7455
      End
      Begin VB.TextBox t_b 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mf 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
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
      Begin VB.Label Label9 
         BackColor       =   &H00FF0000&
         Caption         =   "Vencimiento:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6720
         TabIndex        =   15
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         Caption         =   "Móvil o Base:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         Caption         =   "Item:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label labu 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7440
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6480
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "BASE:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "HORA:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "FECHA:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF0000&
      Caption         =   "Buscar por ítem:"
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
      Left            =   3600
      TabIndex        =   23
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   8880
      Picture         =   "frm_matvence.frx":3525
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "frm_matvence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_alta_Click()
b_alta.Enabled = False
b_graba.Enabled = True
b_edita.Enabled = False
b_imp.Enabled = False
b_cance.Enabled = True
XAlta = 1
Frame1.Enabled = True
mf.Text = Date
mh.Text = Format(Time, "HH:mm")
t_b.Text = frm_menu.data_parse.Recordset("base")
labu.Caption = WElusuario
Combo1.ListIndex = -1
Combo1.SetFocus
t_cant.Text = ""
t_movil.Text = ""
mvence.Text = "__/____"
   

End Sub

Private Sub b_cance_Click()
b_alta.Enabled = True
b_graba.Enabled = False
b_edita.Enabled = True
b_imp.Enabled = True
b_cance.Enabled = False
XAlta = 0
mf.Text = "__/__/____"
mh.Text = "__:__"
t_b.Text = ""
labu.Caption = ""
Combo1.ListIndex = -1
t_cant.Text = ""
t_movil.Text = ""
mvence.Text = "__/____"
Frame1.Enabled = False


End Sub

Private Sub b_edita_Click()
b_alta.Enabled = False
b_graba.Enabled = True
b_edita.Enabled = False
b_imp.Enabled = False
b_cance.Enabled = True
XAlta = 2
Frame1.Enabled = True
Combo1.SetFocus


End Sub

Private Sub b_graba_Click()
Dim Xind As Integer
Dim Xelcod As String
Dim Xbandven As Integer
Dim Xelcodn As Long

Xbandven = 0
Xelcod = ""

If t_cant.Text <> "" And t_movil.Text <> "" And Combo1.ListIndex >= 0 Then
   If XAlta = 1 Then
      data_cons.RecordSource = "Select * from vencim order by id DESC"
      data_cons.Refresh
      If data_cons.Recordset.RecordCount > 0 Then
         t_id.Text = data_cons.Recordset("id") + 1
      Else
         t_id.Text = 1
      End If
      data_vencim.Recordset.AddNew
      data_vencim.Recordset("id") = t_id.Text
      data_vencim.Recordset("fecha") = mf.Text
      data_vencim.Recordset("hora") = mh.Text
      data_vencim.Recordset("base") = t_b.Text
      data_vencim.Recordset("usuario") = labu.Caption
      data_vencim.Recordset("cant") = t_cant.Text
      data_vencim.Recordset("movil") = t_movil.Text
      data_vencim.Recordset("vence") = mvence.Text
      
      For Xind = 1 To Len(Combo1.Text)
          If Xbandven = 8 Then
             Xelcod = Xelcod & Mid(Combo1.Text, Xind, 1)
          End If
          If Mid(Combo1.Text, Xind, 1) = "|" Then
             Xbandven = 8
          End If
      Next
      If Xelcod <> "" And Xbandven = 8 Then
         Xelcodn = Val(Xelcod)
      End If
      data_item.RecordSource = "Select * from stock where id =" & Xelcodn
      data_item.Refresh
      If data_item.Recordset.RecordCount > 0 Then
         data_vencim.Recordset("itemcod") = data_item.Recordset("id")
         data_vencim.Recordset("itemdesc") = data_item.Recordset("descrip")
      Else
         MsgBox "Hay un error en el ITEM seleccionado", vbInformation
         data_vencim.Recordset.CancelUpdate
         Unload Me
      End If
      data_vencim.Recordset.Update
      data_vencim.Refresh
      b_alta.Enabled = True
      b_graba.Enabled = False
      b_edita.Enabled = True
      b_imp.Enabled = True
      b_cance.Enabled = False
      XAlta = 0
      Frame1.Enabled = False
   End If
   If XAlta = 2 Then
      data_vencim.Recordset.Edit
      data_vencim.Recordset("fecha") = mf.Text
      data_vencim.Recordset("hora") = mh.Text
      data_vencim.Recordset("base") = t_b.Text
      data_vencim.Recordset("usuario") = labu.Caption
      data_vencim.Recordset("cant") = t_cant.Text
      data_vencim.Recordset("movil") = t_movil.Text
      data_vencim.Recordset("vence") = mvence.Text
      
      For Xind = 1 To Len(Combo1.Text)
          If Mid(Combo1.Text, Xind, 1) = "|" Then
             Xbandven = 8
          End If
          If Xbandven = 8 Then
             Xelcod = Xelcod & Mid(Combo1.Text, Xind, 1)
          End If
      Next
      If Xelcod <> "" And Xbandven = 8 Then
         Xelcodn = Val(Xelcod)
      End If
      data_item.RecordSource = "Select * from stock where id =" & Xelcodn
      data_item.Refresh
      If data_item.Recordset.RecordCount > 0 Then
         data_vencim.Recordset("itemcod") = data_item.Recordset("id")
         data_vencim.Recordset("itemdesc") = data_item.Recordset("descrip")
      Else
         MsgBox "Hay un error en el ITEM seleccionado", vbInformation
         Unload Me
      End If
      data_vencim.Recordset.Update
      data_vencim.Refresh
      b_alta.Enabled = True
      b_graba.Enabled = False
      b_edita.Enabled = True
      b_imp.Enabled = True
      b_cance.Enabled = False
      XAlta = 0
      Frame1.Enabled = False
   End If
Else
   MsgBox "Verifique datos que faltan ingresar"
   
End If

End Sub

Private Sub b_imp_Click()
Dim Xdesde, Xhasta, xbase As String
xbase = InputBox("INGRESE BASE A LISTAR (99=TODAS)")
Xdesde = InputBox("INGRESE DESDE QUE FECHA")
Xhasta = InputBox("INGRESE HASTA QUE FECHA")

If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If

If xbase = "" Or Xdesde = "" Or Xhasta = "" Then
   MsgBox "Faltan datos para el informe"
Else
   If xbase = 99 Then
      data_cons.RecordSource = "Select * from vencim where fecha >=#" & Format(Xdesde, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xhasta, "yyyy/mm/dd") & "# order by fecha"
      data_cons.Refresh
   Else
      data_cons.RecordSource = "Select * from vencim where fecha >=#" & Format(Xdesde, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xhasta, "yyyy/mm/dd") & "# and base =" & Val(xbase) & " order by fecha"
      data_cons.Refresh
   End If
   If data_cons.Recordset.RecordCount > 0 Then
      data_cons.Recordset.MoveFirst
      Do While Not data_cons.Recordset.EOF
         data_inf.Recordset.AddNew
         data_inf.Recordset("cl_fnac") = data_cons.Recordset("fecha")
         data_inf.Recordset("cl_codced") = data_cons.Recordset("base")
         data_inf.Recordset("cl_nombre") = data_cons.Recordset("usuario")
         data_inf.Recordset("cl_codigo") = data_cons.Recordset("itemcod")
         data_inf.Recordset("info_debit") = data_cons.Recordset("itemdesc")
         data_inf.Recordset("cl_nrovend") = data_cons.Recordset("cant")
         data_inf.Recordset("cl_nro_sup") = data_cons.Recordset("movil")
         data_inf.Recordset("cl_celular") = data_cons.Recordset("vence")
         data_inf.Recordset.Update
         data_cons.Recordset.MoveNext
      Loop
      MsgBox "Terminado"
      cr1.ReportFileName = App.path & "\infvences.rpt"
      cr1.ReportTitle = "Informe medicación y material VENCIDO DESDE: " & Xdesde & " HASTA: " & Xhasta
      cr1.Action = 1
      
   
   End If

End If


End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_cant.SetFocus
End If

End Sub

Private Sub DBGrid1_DblClick()
If IsNull(data_vencim.Recordset("fecha")) = False Then
   mf.Text = data_vencim.Recordset("fecha")
Else
   mf.Text = "__/__/____"
End If
If IsNull(data_vencim.Recordset("hora")) = False Then
   mh.Text = data_vencim.Recordset("hora")
Else
   mh.Text = "__:__"
End If
If IsNull(data_vencim.Recordset("base")) = False Then
   t_b.Text = data_vencim.Recordset("base")
Else
   t_b.Text = 0
End If
If IsNull(data_vencim.Recordset("usuario")) = False Then
   labu.Caption = data_vencim.Recordset("usuario")
Else
   labu.Caption = ""
End If
If IsNull(data_vencim.Recordset("itemdesc")) = False Then
   Combo1.Text = data_vencim.Recordset("itemdesc")
Else
   Combo1.ListIndex = -1
End If
If IsNull(data_vencim.Recordset("cant")) = False Then
   t_cant.Text = data_vencim.Recordset("cant")
Else
   t_cant.Text = ""
End If
If IsNull(data_vencim.Recordset("movil")) = False Then
   t_movil.Text = data_vencim.Recordset("movil")
Else
   t_movil.Text = ""
End If
If IsNull(data_vencim.Recordset("vence")) = False Then
   mvence.Text = Format(data_vencim.Recordset("vence"), "mm/yyyy")
Else
   mvence.Text = "__/____"
End If



End Sub

Private Sub Form_Load()

data_item.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_item.RecordSource = "Select * from stock order by descrip"
data_item.Refresh
If data_item.Recordset.RecordCount > 0 Then
   data_item.Recordset.MoveFirst
   Do While Not data_item.Recordset.EOF
      Combo1.AddItem data_item.Recordset("descrip") & "|" & data_item.Recordset("id")
      data_item.Recordset.MoveNext
   Loop
End If

data_vencim.Connect = "ODBC;DSN=" & Xconexrmt & ";"
data_cons.Connect = "ODBC;DSN=" & Xconexrmt & ";"

If WElusuario = "JFERNAN" Or WElusuario = "MCURBELO" Then
   data_vencim.RecordSource = "Select * from vencim order by fecha"
Else
   data_vencim.RecordSource = "Select * from vencim where base =" & frm_menu.data_parse.Recordset("base") & " order by fecha"
End If
data_vencim.Refresh

data_inf.DatabaseName = App.path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub t_busca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If t_busca.Text <> "" Then
      If WElusuario = "JFERNAN" Or WElusuario = "MCURBELO" Then
         data_vencim.RecordSource = "Select * from vencim where itemdesc >='" & t_busca.Text & "' order by fecha"
      Else
         data_vencim.RecordSource = "Select * from vencim where itemdesc >='" & t_busca.Text & "' and base =" & frm_menu.data_parse.Recordset("base") & " order by fecha"
      End If
      data_vencim.Refresh
   Else
      If WElusuario = "JFERNAN" Or WElusuario = "MCURBELO" Then
         data_vencim.RecordSource = "Select * from vencim order by fecha"
      Else
         data_vencim.RecordSource = "Select * from vencim where base =" & frm_menu.data_parse.Recordset("base") & " order by fecha"
      End If
      data_vencim.Refresh
   End If
End If

End Sub

Private Sub t_cant_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_movil.SetFocus
End If

End Sub

Private Sub t_movil_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mvence.SetFocus
End If

End Sub
