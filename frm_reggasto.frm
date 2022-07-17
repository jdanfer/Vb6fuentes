VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_reggasto 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de entregas de material y medicamentos"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_reggasto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc adoalerta 
      Height          =   735
      Left            =   3480
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "adoalerta"
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
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_verent 
      Caption         =   "data_verent"
      Connect         =   "odbc;dsn=sappnew;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "gastos"
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   6960
      TabIndex        =   19
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton b_busca 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      Picture         =   "frm_reggasto.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Buscar"
      Top             =   6360
      Width           =   495
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7800
      Picture         =   "frm_reggasto.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Salir"
      Top             =   6360
      Width           =   495
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      Picture         =   "frm_reggasto.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cancelar todo el pedido"
      Top             =   6360
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      Picture         =   "frm_reggasto.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Termina con el ingreso de la entrega"
      Top             =   6360
      Width           =   495
   End
   Begin VB.CommandButton b_mod 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_reggasto.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Editar registro"
      Top             =   6360
      Width           =   495
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_reggasto.frx":1FF4
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Nuevo registro"
      Top             =   6360
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos del registro."
      Enabled         =   0   'False
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   8055
      Begin VB.TextBox t_codaut 
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   3600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   6480
         TabIndex        =   35
         ToolTipText     =   "Ingresar aquí la cantidad correcta"
         Top             =   3360
         Width           =   975
      End
      Begin MSAdodcLib.Adodc data_gasto 
         Height          =   375
         Left            =   3000
         Top             =   5400
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
         Caption         =   "data_gasto"
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
      Begin MSAdodcLib.Adodc data_sto 
         Height          =   375
         Left            =   5400
         Top             =   5280
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         Caption         =   "data_sto"
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
      Begin VB.TextBox t_cantant 
         Height          =   495
         Left            =   6960
         TabIndex        =   32
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton b_ed 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   495
         Left            =   7440
         Picture         =   "frm_reggasto.frx":257E
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Graba la modificación realizada de cantidad de una entrega"
         Top             =   3360
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   480
         Left            =   5280
         TabIndex        =   30
         Top             =   3480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_reggasto.frx":2B08
         Height          =   1335
         Left            =   240
         OleObjectBlob   =   "frm_reggasto.frx":2B22
         TabIndex        =   29
         Top             =   4440
         Width           =   7695
      End
      Begin VB.CommandButton b_el 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   495
         Left            =   4800
         Picture         =   "frm_reggasto.frx":369D
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Borrar línea de pedido"
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton b_gr 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   3960
         Picture         =   "frm_reggasto.frx":3C27
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Grabar línea de pedido"
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton b_comienza 
         Height          =   615
         Left            =   2040
         Picture         =   "frm_reggasto.frx":41B1
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Comenzar a ingresar ITEMS a la entrega"
         Top             =   1680
         Width           =   735
      End
      Begin MSMask.MaskEdBox mhor 
         Height          =   375
         Left            =   3600
         TabIndex        =   23
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfec 
         Height          =   375
         Left            =   2040
         TabIndex        =   22
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   5520
         TabIndex        =   20
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton b_cons2 
         Caption         =   "Consultar..."
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton b_cons1 
         Caption         =   "Consultar..."
         Height          =   375
         Left            =   3960
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox t_obs 
         Height          =   375
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   10
         Top             =   3960
         Width           =   5895
      End
      Begin VB.TextBox t_cant 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox t_cod 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox t_cli 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   0
         X2              =   8040
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label labus 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6120
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Usuario:"
         Height          =   375
         Left            =   4920
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Fecha/Hora:"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Observaciones:"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cantidad:"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label labdesc 
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
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2880
         Width           =   7695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Código ítem:"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label labcli 
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
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   7695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cliente:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Label labnrop 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5520
      TabIndex        =   34
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pedido Nro:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3600
      TabIndex        =   33
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4800
      Picture         =   "frm_reggasto.frx":45F3
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2415
   End
End
Attribute VB_Name = "frm_reggasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

End Sub

Private Sub b_alta_Click()
XAlta = 1
Frame1.Enabled = True
mfec.Enabled = True
mhor.Enabled = True
t_cli.Enabled = True
b_cons1.Enabled = True

mfec.SetFocus
mfec.Text = Date
mhor.Text = Format(Time, "HH:mm")

b_alta.Enabled = False
b_mod.Enabled = False
b_graba.Enabled = True
b_cance.Enabled = True
b_busca.Enabled = False
Command7.Enabled = False
If data_gasto.Recordset.RecordCount > 0 Then
'   data_gasto.Recordset.MoveLast
   Data1.Recordset.Edit
   Data1.Recordset("rub_cuotas") = Data1.Recordset("rub_cuotas") + 1
   Data1.Recordset.Update
'   Data1.Refresh
   Text1.Text = Data1.Recordset("rub_cuotas")
Else
   Data1.Recordset.Edit
   Data1.Recordset("rub_cuotas") = Data1.Recordset("rub_cuotas") + 1
   Data1.Recordset.Update
'   Data1.Refresh
   Text1.Text = Data1.Recordset("rub_cuotas")
   
'   Text1.Text = 1
End If
Text3.Text = Text1.Text
'data_verent.RecordSource = "Select * from gastos where id >=" & Text3.Text

'rub_cuotas
labnrop.Caption = Data1.Recordset("notadev") + 1
Data1.Recordset.Edit
Data1.Recordset("notadev") = Data1.Recordset("notadev") + 1
Data1.Recordset.Update
Data1.Refresh


labus.Caption = WElusuario
data_verent.RecordSource = "Select * from gastos where prec =" & labnrop.Caption
data_verent.Refresh
'data_gasto.Recordset.AddNew
t_cant.Text = ""
t_cod.Text = ""
labdesc.Caption = ""
t_obs.Text = ""
t_cli.SetFocus
Text4.Enabled = False


End Sub

Private Sub b_busca_Click()
frm_consgas.Show vbModal

End Sub

Private Sub b_cance_Click()
'If XAlta = 1 Then
'   data_gasto.Recordset.CancelUpdate
'End If
b_alta.Enabled = True
b_mod.Enabled = True
b_graba.Enabled = False
b_cance.Enabled = False
b_busca.Enabled = True
b_gr.Enabled = True
b_el.Enabled = False
b_ed.Enabled = False
t_cant.Enabled = True
t_cod.Enabled = True
Command7.Enabled = True
t_cli.Text = ""
t_cod.Text = ""
labcli.Caption = ""
labdesc.Caption = ""
labnrop.Caption = ""
Text4.Enabled = True
Text4.Text = ""

t_cant.Text = ""
t_obs.Text = ""
mfec.Text = "__/__/____"
mhor.Text = "__:__"
labus.Caption = ""
Text1.Text = ""
Text3.Text = ""
Text2.Text = 0
Frame1.Enabled = False

End Sub

Private Sub b_comienza_Click()
t_cod.SetFocus
mfec.Enabled = False
mhor.Enabled = False
t_cli.Enabled = False
b_cons1.Enabled = False
If t_cli.Text = "" Then
   t_cli.Text = 0
End If
data_verent.RecordSource = "Select * from gastos where prec =" & labnrop.Caption
'data_verent.RecordSource = "Select * from gastos where codcli =" & t_cli.Text & " and fecha =#" & Format(mfec.Text, "yyyy/mm/dd") & "#"
data_verent.Refresh

End Sub

Private Sub b_cons1_Click()
frm_conscli.Show vbModal

End Sub

Private Sub b_cons2_Click()
frm_consitem.Show vbModal

End Sub

Private Sub b_cons2_GotFocus()
If t_cod.Text = "" Then
Else
   t_cant.SetFocus
End If

End Sub

Private Sub b_ed_Click()
Dim Laclave As String
Laclave = InputBox("Igrese la clave de autorización para modificar", "Stock")
If Laclave = "12345" Then
    labus.Caption = WElusuario
    'data_gasto.Recordset.Edit
    data_gasto.RecordSource = "Select * from gastos where id =" & Val(Text1.Text)
    data_gasto.Refresh
    If Text4.Text <> data_gasto.Recordset("cant") Then
        data_gasto.Recordset("cant") = Text4.Text
        data_gasto.Recordset("obs") = t_obs.Text
        data_gasto.Recordset("usuario") = labus.Caption
        data_gasto.Recordset.Update
    End If
    
    DBGrid1.Enabled = True
    'data_verent.RecordSource = "Select * from gastos where id >=" & Text1.Text
    'data_verent.RecordSource = "Select * from gastos where codcli =" & t_cli.Text & " and fecha =#" & Format(mfec.Text, "yyyy/mm/dd") & "#"
    data_verent.RecordSource = "Select * from gastos where id <" & 0
    
    data_verent.Refresh
    DBGrid1.Enabled = False
    data_sto.RecordSource = "Select * from stock where id =" & t_cod.Text
    data_sto.Refresh
    If data_sto.Recordset.RecordCount > 0 Then
    '   data_sto.Recordset.EditMode
       data_sto.Recordset("actual") = data_sto.Recordset("actual") - Val(Text4.Text) + Val(t_cant.Text)
       data_sto.Recordset.Update
    End If
    b_el.Enabled = False
    b_ed.Enabled = False
    t_cod.Text = ""
    labdesc.Caption = ""
    t_cant.Enabled = True
    t_cant.Text = ""
    t_cantant.Visible = True
    t_cantant.Enabled = True
    t_cantant.Text = ""
    t_cantant.Visible = False
    t_obs.Text = ""
    b_gr.Enabled = True
    b_alta.Enabled = True
    b_mod.Enabled = True
    b_graba.Enabled = False
    b_cance.Enabled = False
    b_busca.Enabled = True
    Command7.Enabled = True
    t_cli.Text = ""
    t_cod.Enabled = True
    t_cod.Text = ""
    labcli.Caption = ""
    labdesc.Caption = ""
    t_cant.Text = ""
    t_obs.Text = ""
    mfec.Text = "__/__/____"
    mhor.Text = "__:__"
    labus.Caption = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    
    Frame1.Enabled = False
    XAlta = 0
Else
   b_cance_Click
End If


End Sub

Private Sub b_el_Click()
Dim Xelmqueborra As String
Xelmqueborra = MsgBox("Desea eliminar el registro de: " & data_gasto.Recordset("descrip") & "??", vbExclamation + vbYesNo, "STOCK")
If Xelmqueborra = vbYes Then
   data_sto.RecordSource = "Select * from stock where id =" & t_cod.Text
   data_sto.Refresh
   If data_sto.Recordset.RecordCount > 0 Then
'      data_sto.Recordset.Edit
      data_sto.Recordset("actual") = data_sto.Recordset("actual") + t_cant.Text
      data_sto.Recordset.Update
   End If
   data_gasto.Recordset.Delete
   data_gasto.RecordSource = "Select * from gastos where prec =" & labnrop.Caption
   data_gasto.Refresh
'   data_verent.RecordSource = "Select * from gastos where id =" & Text1.Text
   data_verent.RecordSource = "Select * from gastos where prec =" & labnrop.Caption
   
   data_verent.Refresh
   b_ed.Enabled = False
   b_el.Enabled = False
   b_gr.Enabled = True
   b_alta.Enabled = True
   b_mod.Enabled = True
   b_graba.Enabled = False
   b_cance.Enabled = False
   b_busca.Enabled = True
   Command7.Enabled = True
   t_cli.Text = ""
   t_cod.Text = ""
   labcli.Caption = ""
   labdesc.Caption = ""
   t_cant.Text = ""
   t_obs.Text = ""
   mfec.Text = "__/__/____"
   mhor.Text = "__:__"
   labus.Caption = ""
   Text1.Text = ""
   Text3.Text = ""
   Text2.Text = 0
   Frame1.Enabled = False
Else
   b_ed.Enabled = False
   b_el.Enabled = False
   b_alta.Enabled = True
   b_mod.Enabled = True
   b_graba.Enabled = False
   b_cance.Enabled = False
   b_busca.Enabled = True
   Command7.Enabled = True
   t_cli.Text = ""
   t_cod.Text = ""
   labcli.Caption = ""
   labdesc.Caption = ""
   t_cant.Text = ""
   t_obs.Text = ""
   mfec.Text = "__/__/____"
   mhor.Text = "__:__"
   labus.Caption = ""
   Text1.Text = ""
   Text3.Text = ""
   Text2.Text = 0
   Frame1.Enabled = False

End If

End Sub

Private Sub b_gr_Click()
'On Error GoTo Quepasoal

If t_cli.Text = "" Then
   MsgBox "No ingresó dato en cliente", vbCritical
Else
   If t_cod.Text = "" Then
      MsgBox "No ingresó dato en ITEM", vbCritical
   Else
      If XAlta = 1 Then
         data_gasto.Recordset.AddNew
         data_gasto.Recordset("id") = Text1.Text
         data_gasto.Recordset("fecha") = Date
         data_gasto.Recordset("hora") = Format(Time, "HH:mm:ss")
         data_gasto.Recordset("codprod") = t_cod.Text
         data_gasto.Recordset("descrip") = Mid(labdesc.Caption, 1, 90)
         data_gasto.Recordset("prec") = labnrop.Caption
         data_gasto.Recordset("codcli") = t_cli.Text
         data_gasto.Recordset("nomcli") = Mid(labcli.Caption, 1, 90)
         If t_codaut.Text <> "" Then
            data_gasto.Recordset("descrip2") = t_codaut.Text
         End If
         If t_cant.Text = "" Then
            t_cant.Text = 0
         End If
         If t_cant.Text = "" Then
            t_cant.Text = 0
         End If
         data_gasto.Recordset("cant") = t_cant.Text
         data_gasto.Recordset("obs") = t_obs.Text
         If mhor.Text = "__:__" Then
         Else
            data_gasto.Recordset("horret") = mhor.Text
         End If
'         If mfecret.Text = "__/__/____" Then
'         Else
'            data_gasto.Recordset("fecret") = mfecret.Text
'         End If
         data_gasto.Recordset("usuario") = labus.Caption
         data_gasto.Recordset.Update
         Data1.Recordset.Edit
         Text1.Text = Data1.Recordset("rub_cuotas") + 1
         Data1.Recordset("rub_cuotas") = Text1.Text
         Data1.Recordset.Update
         Data1.Refresh
'         DBGrid1.Enabled = True
'         data_verent.RecordSource = "Select * from gastos where id >=" & Text3.Text
'         data_verent.RecordSource = "Select * from gastos where codcli =" & t_cli.Text & " and fecha =#" & Format(mfec.Text, "yyyy/mm/dd") & "#"
         data_verent.RecordSource = "Select * from gastos where prec =" & labnrop.Caption
         data_verent.Refresh
'         DBGrid1.Enabled = False
         data_sto.RecordSource = "Select * from stock where id =" & t_cod.Text
         data_sto.Refresh
'         If Not data_sto.Recordset.NoMatch Then
         If data_sto.Recordset.RecordCount > 0 Then
'            data_sto.Recordset.Edit
            data_sto.Recordset("actual") = data_sto.Recordset("actual") - t_cant.Text
            data_sto.Recordset.Update
         End If
         t_cod.Text = ""
         labdesc.Caption = ""
         t_cant.Text = ""
         t_obs.Text = ""
         t_cod.SetFocus
      Else
         MsgBox "No se puede modificar, presione el botón EDITAR para realizarlo"
         t_cod.Text = ""
         labdesc.Caption = ""
         t_cant.Text = ""
         t_obs.Text = ""
      
      End If
   End If
End If
'Exit Sub

'Quepasoal:
'          If Err.Number = 3146 Then
'             MsgBox "Error al intentar grabar los datos, VERIFIQUE DATOS y vuelta a intentar", vbExclamation
'          Else
'             MsgBox "Error " & Err.Number & " AVISE A INFORMÁTICA. VERIFIQUE DATOS E INTENTE VOLVER A GRABAR", vbCritical
'          End If

End Sub

Private Sub b_graba_Click()

labnrop.Caption = ""
data_verent.RecordSource = "Select * from gastos where id <" & 0
data_verent.Refresh

b_alta.Enabled = True
b_mod.Enabled = True
b_graba.Enabled = False
b_cance.Enabled = False
b_busca.Enabled = True
Command7.Enabled = True
t_cli.Text = ""
t_cod.Text = ""
labcli.Caption = ""
labdesc.Caption = ""
t_cant.Text = ""
t_obs.Text = ""
mfec.Text = "__/__/____"
mhor.Text = "__:__"
labus.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Frame1.Enabled = False
XAlta = 0

End Sub

Private Sub b_mod_Click()
If t_cod.Text <> "" Then
   If t_cli.Text <> "" Then
'      frm_reggasto.data_verent.RecordSource = "Select * from gastos where id =" & Data1.Recordset("id")
'      frm_reggasto.Refresh
      XAlta = 0
      b_alta.Enabled = False
      b_mod.Enabled = False
      b_graba.Enabled = False
      b_cance.Enabled = True
      b_busca.Enabled = False
      Frame1.Enabled = True
      DBGrid1.Enabled = True
      b_gr.Enabled = False
      b_ed.Enabled = True
      b_el.Enabled = True
      t_cantant.Text = t_cant.Text
      t_cod.Enabled = False
      t_cant.Enabled = False
      Text4.Enabled = True
      If Text4.Enabled = True Then
         Text4.SetFocus
      End If
   End If
End If


End Sub

Private Sub Command7_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
Dim Xxelme As String
Xxelme = MsgBox("Desea editar el registro para MODIFICAR DATOS??", vbInformation + vbYesNo, "STOCK")
If Xxelme = vbYes Then
   b_el.Enabled = True
   b_ed.Enabled = True
   b_gr.Enabled = False
   b_ed.SetFocus
   data_gasto.RecordSource = "Select * from gastos where id =" & data_verent.Recordset("id")
   data_gasto.Refresh
   If data_gasto.Recordset.RecordCount > 0 Then
      If IsNull(data_gasto.Recordset("id")) = False Then
         Text1.Text = data_gasto.Recordset("id")
      Else
         Text1.Text = 0
      End If
      If IsNull(data_gasto.Recordset("codcli")) = False Then
         t_cli.Text = data_gasto.Recordset("codcli")
      Else
         t_cli.Text = 0
      End If
      If IsNull(data_gasto.Recordset("codprod")) = False Then
         t_cod.Text = data_gasto.Recordset("codprod")
      Else
         t_cod.Text = 0
      End If
      If IsNull(data_gasto.Recordset("nomcli")) = False Then
         labcli.Caption = data_gasto.Recordset("nomcli")
      Else
         labcli.Caption = ""
      End If
      If IsNull(data_gasto.Recordset("descrip")) = False Then
         labdesc.Caption = data_gasto.Recordset("descrip")
      Else
         labdesc.Caption = ""
      End If
      If IsNull(data_gasto.Recordset("cant")) = False Then
         t_cant.Text = data_gasto.Recordset("cant")
         t_cantant.Text = data_gasto.Recordset("cant")
      Else
         t_cant.Text = 0
         t_cantant.Text = 0
      End If
      If IsNull(data_gasto.Recordset("obs")) = False Then
         t_obs.Text = data_gasto.Recordset("obs")
      Else
         t_obs.Text = ""
      End If
      If IsNull(data_gasto.Recordset("fecha")) = False Then
         mfec.Text = data_gasto.Recordset("fecha")
      Else
         mfec.Text = "__/__/____"
      End If
      If IsNull(data_gasto.Recordset("hora")) = False Then
         mhor.Text = Format(data_gasto.Recordset("hora"), "HH:mm")
      Else
         mhor.Text = "__:__"
      End If
      If IsNull(data_gasto.Recordset("usuario")) = False Then
         labus.Caption = data_gasto.Recordset("usuario")
      Else
         labus.Caption = ""
      End If
   Else
      b_el.Enabled = False
      b_ed.Enabled = False
      b_gr.Enabled = True
      MsgBox "No encontrado!!"
   End If
Else
   b_el.Enabled = False
   b_ed.Enabled = False
   b_gr.Enabled = True
End If

End Sub

Private Sub Form_Load()
'data_cli.DatabaseName = App.Path & "\" & Trim(Xlabdd)
data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.RecordSource = "clieco"
data_cli.Refresh
adoalerta.ConnectionString = "dsn=" & Xconexrmt

data_gasto.ConnectionString = "dsn=" & Xconexrmt
data_gasto.RecordSource = "Select * from gastos where id =" & 18 & " order by id"
data_gasto.Refresh
data_sto.ConnectionString = "dsn=" & Xconexrmt
data_sto.RecordSource = "Select * from stock where id =" & 18
data_sto.Refresh
data_verent.RecordSource = "Select * from gastos where id <" & 0
data_verent.Refresh
Data1.DatabaseName = App.path & "\parse.mdb"
Data1.RecordSource = "parsec0"
Data1.Refresh


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhor.SetFocus
End If

End Sub

Private Sub mfecret_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_graba.SetFocus
End If

End Sub

Private Sub mhor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_cli.SetFocus
End If

End Sub

Private Sub t_cant_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If b_gr.Enabled = True Then
      b_gr.SetFocus
   Else
      If b_ed.Enabled = True Then
         b_ed.SetFocus
      End If
   End If
End If

End Sub

Private Sub t_cant_LostFocus()
Dim Xconfalerta As String
t_codaut.Text = ""
If t_cant.Text <> "" Then
   If t_cod.Text <> "" Then
      adoalerta.RecordSource = "Select * from stock where id =" & t_cod.Text
      adoalerta.Refresh
      If adoalerta.Recordset.RecordCount > 0 Then
         If IsNull(adoalerta.Recordset("alerta")) = False Then
            If adoalerta.Recordset("alerta") > 0 Then
               If Val(t_cant.Text) > adoalerta.Recordset("alerta") Then
                  MsgBox "ATENCION! Cantidad mayor a la establecida, VERIFIQUE!!", vbCritical
                  Xconfalerta = InputBox("Ingrese código de autorización para confirmar", "Stock")
                  If Xconfalerta <> "" Then
                     t_codaut.Text = Xconfalerta
                  Else
                     t_codaut.Text = "Sin Cod"
                  End If
               End If
            End If
         End If
      End If
   End If
End If

End Sub

Private Sub t_cli_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_comienza.SetFocus
End If

End Sub

Private Sub t_cli_LostFocus()
If t_cli.Text <> "" Then
   data_cli.RecordSource = "Select * from clieco where id =" & t_cli.Text
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      labcli.Caption = data_cli.Recordset("nombre")
   Else
      labcli.Caption = ""
      MsgBox "No ECONTRADO"
   End If
End If

End Sub

Private Sub t_cod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If t_cod.Text = "" Then
'      b_graba.SetFocus
      b_cons2.SetFocus
   Else
      If t_cant.Enabled = True Then
         t_cant.SetFocus
      End If
   End If
End If

End Sub

Private Sub t_cod_LostFocus()
If t_cod.Text <> "" Then
   data_sto.RecordSource = "Select * from stock where id =" & t_cod.Text
   data_sto.Refresh
   If data_sto.Recordset.RecordCount > 0 Then
      If data_sto.Recordset("actual") <= 0 Then
         MsgBox "Sin Stock, VERIFIQUE!!", vbCritical
      End If
      labdesc.Caption = data_sto.Recordset("descrip")
   Else
      labdesc.Caption = ""
      MsgBox "NO ENCONTRADO"
   End If
End If

End Sub

Private Sub t_obs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfec.SetFocus
End If

End Sub

Private Sub t_ret_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfecret.SetFocus
End If

End Sub

