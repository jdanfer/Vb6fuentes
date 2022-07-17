VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_compsto 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de compras"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_compsto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   8895
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_comp 
      Height          =   375
      Left            =   3000
      Top             =   7440
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "data_comp"
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
   Begin MSAdodcLib.Adodc data_item 
      Height          =   495
      Left            =   5880
      Top             =   7200
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "data_item"
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
   Begin MSAdodcLib.Adodc data_bus 
      Height          =   375
      Left            =   120
      Top             =   7320
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "data_bus"
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
   Begin MSFlexGridLib.MSFlexGrid DBGrid1 
      Height          =   1335
      Left            =   240
      TabIndex        =   45
      Top             =   5760
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2355
      _Version        =   393216
      BackColorBkg    =   12615680
      SelectionMode   =   1
   End
   Begin VB.TextBox t_b 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4920
      TabIndex        =   36
      ToolTipText     =   "Formato de fecha para ingresar para la búsqueda DD/MM/AAAA"
      Top             =   5400
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C00000&
      Caption         =   "Buscar por boleta"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   35
      Top             =   5400
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "Buscar por fecha"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   34
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5040
      Picture         =   "frm_compsto.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Informes"
      Top             =   7080
      Width           =   615
   End
   Begin VB.CommandButton b_elim 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4080
      Picture         =   "frm_compsto.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7080
      Width           =   615
   End
   Begin VB.CommandButton b_canc 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      Picture         =   "frm_compsto.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar acción"
      Top             =   7080
      Width           =   615
   End
   Begin VB.CommandButton b_grab 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      Picture         =   "frm_compsto.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Graba y confirma factura"
      Top             =   7080
      Width           =   615
   End
   Begin VB.CommandButton b_mod 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      Picture         =   "frm_compsto.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Editar"
      Top             =   7080
      Width           =   615
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_compsto.frx":1FF4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo registro"
      Top             =   7080
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos de la compra"
      Enabled         =   0   'False
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin MSAdodcLib.Adodc data_linfac2 
         Height          =   495
         Left            =   1920
         Top             =   4440
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         Caption         =   "data_linfac2"
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
      Begin MSAdodcLib.Adodc data_lab 
         Height          =   375
         Left            =   4440
         Top             =   1800
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
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "data_lab"
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
      Begin VB.Data data_linfac 
         Caption         =   "data_linfac"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4200
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSAdodcLib.Adodc data_actitem 
         Height          =   375
         Left            =   5760
         Top             =   4800
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
         Caption         =   "data_actitem"
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
         Height          =   375
         Left            =   7320
         TabIndex        =   42
         Top             =   4680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cbomon 
         Height          =   360
         ItemData        =   "frm_compsto.frx":257E
         Left            =   2280
         List            =   "frm_compsto.frx":2588
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton op20 
         Caption         =   "IVA 22%"
         Height          =   255
         Left            =   6600
         TabIndex        =   39
         Top             =   3720
         Width           =   1455
      End
      Begin VB.OptionButton op10 
         Caption         =   "IVA 10%"
         Height          =   255
         Left            =   4800
         TabIndex        =   38
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox t_codcom 
         Height          =   360
         Left            =   4560
         TabIndex        =   37
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   1020
         Left            =   120
         TabIndex        =   33
         Top             =   3960
         Width           =   4095
      End
      Begin VB.TextBox t_idit 
         Height          =   360
         Left            =   7080
         TabIndex        =   32
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_id 
         Height          =   375
         Left            =   6840
         TabIndex        =   31
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton b_git 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4200
         Picture         =   "frm_compsto.frx":2597
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Termina la factura"
         Top             =   4440
         Width           =   615
      End
      Begin VB.CommandButton b_eliit 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   5160
         Picture         =   "frm_compsto.frx":2B21
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Elimina ítem seleccionado"
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox t_cant 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6840
         TabIndex        =   22
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox t_pre 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2280
         TabIndex        =   20
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Buscar Item"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox t_codprod 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Top             =   2400
         Width           =   1935
      End
      Begin MSMask.MaskEdBox mfec 
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         Left            =   2280
         TabIndex        =   12
         Top             =   840
         Width           =   5895
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_compsto.frx":30AB
         Left            =   6240
         List            =   "frm_compsto.frx":30B8
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox t_bol 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label grupo 
         Height          =   375
         Left            =   6240
         TabIndex        =   46
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label labsub 
         BackColor       =   &H0080FFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   44
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sub-Total:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         Caption         =   "Moneda:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label labtotf 
         BackColor       =   &H0080FFFF&
         Height          =   255
         Left            =   6840
         TabIndex        =   28
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080FFFF&
         Caption         =   "Total factura:"
         Height          =   255
         Left            =   5280
         TabIndex        =   27
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label labiva 
         BackColor       =   &H0080FFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   26
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080FFFF&
         Caption         =   "IVA:"
         Height          =   255
         Left            =   3120
         TabIndex        =   25
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label labtot 
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF0000&
         Caption         =   "Total x item:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         Caption         =   "Cantidad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5040
         TabIndex        =   21
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         Caption         =   "Importe unitario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label labdesc 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2880
         Width           =   7935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF0000&
         Caption         =   "Item:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         X1              =   0
         X2              =   8400
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Comercio:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Tipo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4680
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Nro. de Boleta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   120
      Picture         =   "frm_compsto.frx":30D4
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1695
   End
End
Attribute VB_Name = "frm_compsto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Xtotacom As Double
Public Xtotaiv, Xsubtota As Double

Private Sub b_alta_Click()
Frame1.Enabled = True
data_comp.RecordSource = "Select * from compras where id >=" & 800 & " order by id DESC"
data_comp.Refresh
If data_comp.Recordset.RecordCount > 0 Then
'   data_comp.Recordset.MoveLast
   data_comp.Recordset.MoveFirst
   t_id.Text = data_comp.Recordset("id")
   t_id.Text = t_id.Text + 1
Else
   t_id.Text = 1
End If
'data_comp.Recordset.AddNew
XAlta = 1
t_bol.SetFocus
b_alta.Enabled = False
b_mod.Enabled = False
b_grab.Enabled = True
b_canc.Enabled = True
b_elim.Enabled = False
b_eliit.Enabled = False
b_imp.Enabled = False
borra_comp
labtotf.Caption = 0
labsub.Caption = 0
labtot.Caption = 0
labiva.Caption = 0

cbomon.ListIndex = 0


End Sub

Public Function borra_comp()
t_bol.Text = ""
Combo1.ListIndex = -1
Combo2.ListIndex = -1
mfec.Text = "__/__/____"
t_codprod.Text = ""
labdesc.Caption = ""
t_pre.Text = ""
t_cant.Text = ""
labtot.Caption = ""
labiva.Caption = ""
labtotf.Caption = ""
List1.Clear
op10.value = False
op20.value = False

End Function

Private Sub b_bus_Click()
frm_buscompit.Show vbModal

End Sub



Private Sub b_canc_Click()
Dim Msgborrar As String
Msgborrar = MsgBox("Desea CANCELAR LA FACTURA SIN GRABAR? NRO: " & t_bol.Text, vbInformation + vbYesNo, "Borrar")
If Msgborrar = vbYes Then
'   data_comp.Recordset.FindFirst "id =" & t_id.Text
   data_comp.RecordSource = "Select * from compras where id =" & t_id.Text
   data_comp.Refresh
   If data_comp.Recordset.RecordCount > 0 Then
      data_linfac.RecordSource = "Select * from lineascomp where codbol =" & data_comp.Recordset("nrobol")
      data_linfac.Refresh
      If data_linfac.Recordset.RecordCount > 0 Then
         data_linfac.Recordset.MoveFirst
         Do While Not data_linfac.Recordset.EOF
            data_linfac.Recordset.Delete
            data_linfac.Recordset.MoveNext
         Loop
      End If
      data_comp.Recordset.Delete
      data_comp.Refresh
      data_bus.Refresh
'      iguala_comp
      Frame1.Enabled = True
      borra_comp
      Frame1.Enabled = False
   Else
'      MsgBox "No se encontró la factura, VERIFIQUE!"
   End If
End If

XAlta = 0
b_alta.Enabled = True
b_mod.Enabled = True
b_grab.Enabled = False
b_canc.Enabled = False
b_imp.Enabled = True
b_elim.Enabled = True
t_bol.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
mfec.Enabled = True
borra_comp
Frame1.Enabled = False

End Sub

Private Sub b_eliit_Click()
Dim Xtext2, Xcaden2 As String
Dim Xlargo2, Xcuen2, Xidd2 As Long
Dim Xbannn2 As Integer
Xbannn2 = 0
Xtext2 = List1.List(List1.ListIndex)
Xlargo2 = Len(Trim(Xtext2))
Xcuen2 = 0
For Xcuen2 = 1 To Xlargo2
    If Xbannn2 = 1 Then
       Xcaden2 = Xcaden2 & Mid(Xtext2, Xcuen2, 1)
    End If
    If Mid(Xtext2, Xcuen2, 1) = ">" Then
       Xbannn2 = 1
    End If
Next
Xidd2 = Val(Trim(Xcaden2))
'data_linfac.RecordSource = "Select "
'data_linfac.Recordset.FindFirst "id =" & Xidd2
data_linfac.RecordSource = "Select * from lineascomp where id =" & Xidd2
data_linfac.Refresh
If data_linfac.Recordset.RecordCount > 0 Then
   data_actitem.RecordSource = "Select * from stock where id =" & data_linfac.Recordset("codprod")
   data_actitem.Refresh
   If data_actitem.Recordset.RecordCount > 0 Then
'      data_actitem.Recordset.Edit
      data_actitem.Recordset("actual") = data_actitem.Recordset("actual") - data_linfac.Recordset("cant")
      data_actitem.Recordset.Update
   End If
   data_linfac.Recordset.Delete
   data_linfac.Refresh
'   data_comp.Recordset.FindFirst "nrobol =" & t_bol.Text
   data_comp.RecordSource = "Select * from compras where nrobol =" & t_bol.Text
   data_comp.Refresh
   If data_comp.Recordset.RecordCount > 0 Then
'      data_comp.Recordset.Edit
      data_comp.Recordset("totiva") = Xtotaiv
      data_comp.Recordset("totfac") = Xtotacom
      data_comp.Recordset("totuni") = Xsubtota
      data_comp.Recordset.Update
      XAlta = 3
   Else
      MsgBox "No se encontró el cabezal de factura"
   End If
   borra_comp
   iguala_comp

Else
   MsgBox "No se encuentra el producto en la factura, VERIFIQUE!!"
   t_codprod.Text = ""
   t_idit.Text = ""
   labdesc.Caption = ""
   t_pre.Text = ""
   t_cant.Text = ""
   labtot.Caption = ""
   op10.value = False
   op20.value = False
End If

Xbannn2 = 0

End Sub

Private Sub b_elim_Click()
Dim Msgborrar As String
Msgborrar = MsgBox("Desea borrar todos los datos de la factura seleccionada? NRO: " & data_bus.Recordset("nrobol"), vbInformation + vbYesNo, "Borrar")
If Msgborrar = vbYes Then
'   data_comp.Recordset.FindFirst "id =" & data_bus.Recordset("id")
   data_comp.RecordSource = "Select * from compras where id =" & data_bus.Recordset("id")
   data_comp.Refresh
   If data_comp.Recordset.RecordCount > 0 Then
      data_linfac.RecordSource = "Select * from lineascomp where codbol =" & data_comp.Recordset("nrobol") & " and codcom =" & data_comp.Recordset("codcomer")
      data_linfac.Refresh
      If data_linfac.Recordset.RecordCount > 0 Then
         data_linfac.Recordset.MoveFirst
         Do While Not data_linfac.Recordset.EOF
            data_linfac.Recordset.Delete
            data_linfac.Recordset.MoveNext
         Loop
      End If
      data_comp.Recordset.Delete
      data_comp.Refresh
      data_bus.Refresh
'      iguala_comp
      Frame1.Enabled = True
      borra_comp
      Frame1.Enabled = False
   Else
      MsgBox "No se encontró la factura, VERIFIQUE!"
   End If
End If

End Sub

Private Sub b_git_Click()

If XAlta = 1 Then
   If t_bol.Enabled = True Then
      If t_bol.Text <> "" Then
         If mfec.Text <> "__/__/____" Then
            If t_codprod.Text <> "" Then
               data_comp.Recordset.AddNew
               data_comp.Recordset("id") = t_id.Text
               data_comp.Recordset("fecha") = mfec.Text
               data_comp.Recordset("nrobol") = t_bol.Text
               data_comp.Recordset("codcomer") = t_codcom.Text
               data_comp.Recordset("nomcomer") = Combo2.Text
               data_comp.Recordset("tipobol") = Combo1.ListIndex
               data_comp.Recordset("tipobold") = Combo1.Text
               data_comp.Recordset("moneda") = cbomon.ListIndex
               data_comp.Recordset("monedad") = cbomon.Text
               data_comp.Recordset.Update
               data_linfac.RecordSource = "Select * from lineascomp where id >=" & 4500 & " order by id DESC"
               data_linfac.Refresh
               If data_linfac.Recordset.RecordCount > 0 Then
                  t_idit.Text = data_linfac.Recordset("id") + 1
               Else
                  t_idit.Text = 1
               End If
               data_linfac.Recordset.AddNew
               data_linfac.Recordset("id") = t_idit.Text
               data_linfac.Recordset("codbol") = t_bol.Text
               data_linfac.Recordset("fecha") = mfec.Text
               data_linfac.Recordset("codprod") = t_codprod.Text
               data_linfac.Recordset("coddesc") = labdesc.Caption
               data_linfac.Recordset("cant") = t_cant.Text
               data_linfac.Recordset("precuni") = CDbl(t_pre.Text)
               data_linfac.Recordset("totprod") = CDbl(labtot.Caption)
               data_linfac.Recordset("codcom") = t_codcom.Text
               If grupo.Caption <> "" Then
                  data_linfac.Recordset("grupo") = Val(grupo.Caption)
               Else
                  data_linfac.Recordset("grupo") = 0
               End If
               If op10.value = True Then
                  data_linfac.Recordset("opiva") = 1
               Else
                  If op20.value = True Then
                     data_linfac.Recordset("opiva") = 2
                  Else
                     data_linfac.Recordset("opiva") = 0
                  End If
               End If
               data_linfac.Recordset.Update
               data_actitem.RecordSource = "Select * from stock where id =" & t_codprod.Text
               data_actitem.Refresh
               If data_actitem.Recordset.RecordCount > 0 Then
'                  data_actitem.Recordset.Edit
                  data_actitem.Recordset("actual") = data_actitem.Recordset("actual") + t_cant.Text
                  If IsNull(data_actitem.Recordset("preuni")) = False Then
                     If data_actitem.Recordset("preuni") = t_pre.Text Then
                     Else
                        data_actitem.Recordset("preuni") = CDbl(t_pre.Text)
                     End If
                  Else
                     data_actitem.Recordset("preuni") = CDbl(t_pre.Text)
                  End If
                  data_actitem.Recordset.Update
               End If
               List1.AddItem labdesc.Caption & "----" & t_codprod.Text & "--->" & t_idit.Text
               If labiva.Caption = "" Then
                  labiva.Caption = 0
               End If
               If labtotf.Caption = "" Then
                  labtotf.Caption = 0
               End If
               If labtot.Caption = "" Then
                  labtot.Caption = 0
               End If
               If labsub.Caption = "" Then
                  labsub.Caption = 0
               End If
               labsub.Caption = CDbl(labsub.Caption) + CDbl(labtot.Caption)
               If op20.value = True Then
                  labiva.Caption = CDbl(labsub.Caption) * 0.22
                  labtotf.Caption = CDbl(labsub.Caption) + CDbl(labiva.Caption)
               Else
                  If op10.value = True Then
                     labiva.Caption = CDbl(labsub.Caption) * 0.1
                     labtotf.Caption = CDbl(labsub.Caption) + CDbl(labiva.Caption)
                  Else
                     labiva.Caption = 0
                     labtotf.Caption = CDbl(labsub.Caption) + CDbl(labiva.Caption)
                  End If
               End If
               data_comp.Refresh
               labiva.Caption = Format(labiva.Caption, "Standard")
               labtotf.Caption = Format(labtotf.Caption, "Standard")
               t_bol.Enabled = False
               Combo1.Enabled = False
               Combo2.Enabled = False
               mfec.Enabled = False
               t_codprod.SetFocus
               t_codprod.Text = ""
               t_pre.Text = ""
               t_cant.Text = ""
               labtot.Caption = 0
               
            Else
               MsgBox "No ingresó código de ITEM"
               
            End If
         Else
            MsgBox "Verifique FECHA de factura"
         End If
      Else
         MsgBox "Verifique número de FACTURA"
      End If
      t_bol.Enabled = False
      Combo1.Enabled = False
      Combo2.Enabled = False
      mfec.Enabled = False
   Else
'      data_linfac.RecordSource = "Select * from lineascomp order by id"
      data_linfac.RecordSource = "Select * from lineascomp where id >=" & 4500 & " order by id DESC"
      data_linfac.Refresh
      If data_linfac.Recordset.RecordCount > 0 Then
'         data_linfac.Recordset.MoveLast
         t_idit.Text = data_linfac.Recordset("id") + 1
      Else
         t_idit.Text = 1
      End If
      data_linfac.RecordSource = "Select * from lineascomp where codbol =" & t_bol.Text & " order by id"
      data_linfac.Refresh
      If data_linfac.Recordset.RecordCount > 0 Then
'         data_linfac.Recordset.FindFirst "codprod =" & t_codprod.Text
         data_linfac2.RecordSource = "Select * from lineascomp where codprod =" & t_codprod.Text & " and codbol =" & t_bol.Text
         data_linfac2.Refresh
         If data_linfac2.Recordset.RecordCount > 0 Then
            MsgBox "Ya existe ése ITEM en la factura, puede ELIMINAR y volver a ingresarlo"
            t_codprod.SetFocus
         Else
            data_linfac.Recordset.AddNew
            data_linfac.Recordset("id") = t_idit.Text
            data_linfac.Recordset("codbol") = t_bol.Text
            data_linfac.Recordset("fecha") = mfec.Text
            data_linfac.Recordset("codprod") = t_codprod.Text
            data_linfac.Recordset("coddesc") = labdesc.Caption
            data_linfac.Recordset("cant") = t_cant.Text
            data_linfac.Recordset("precuni") = CDbl(t_pre.Text)
            data_linfac.Recordset("totprod") = CDbl(labtot.Caption)
            data_linfac.Recordset("codcom") = t_codcom.Text
            If grupo.Caption <> "" Then
               data_linfac.Recordset("grupo") = Val(grupo.Caption)
            Else
               data_linfac.Recordset("grupo") = 0
            End If
            If op10.value = True Then
               data_linfac.Recordset("opiva") = 1
            Else
               If op20.value = True Then
                  data_linfac.Recordset("opiva") = 2
               Else
                  data_linfac.Recordset("opiva") = 0
               End If
            End If
            data_linfac.Recordset.Update
            data_actitem.RecordSource = "Select * from stock where id =" & t_codprod.Text
            data_actitem.Refresh
            If data_actitem.Recordset.RecordCount > 0 Then
'               data_actitem.Recordset.Edit
               data_actitem.Recordset("actual") = data_actitem.Recordset("actual") + t_cant.Text
               If IsNull(data_actitem.Recordset("preuni")) = False Then
                  If data_actitem.Recordset("preuni") = t_pre.Text Then
                  Else
                     data_actitem.Recordset("preuni") = CDbl(t_pre.Text)
                  End If
               Else
                  data_actitem.Recordset("preuni") = CDbl(t_pre.Text)
               End If
               data_actitem.Recordset.Update
            End If
            List1.AddItem labdesc.Caption & "----" & t_codprod.Text & "--->" & t_idit.Text
            If labiva.Caption = "" Then
               labiva.Caption = 0
            End If
            If labtotf.Caption = "" Then
               labtotf.Caption = 0
            End If
            If labtot.Caption = "" Then
               labtot.Caption = 0
            End If
            If labsub.Caption = "" Then
               labsub.Caption = 0
            End If
            labsub.Caption = CDbl(labsub.Caption) + CDbl(labtot.Caption)
            If op20.value = True Then
               labiva.Caption = CDbl(labsub.Caption) * 0.22
               labtotf.Caption = CDbl(labsub.Caption) + CDbl(labiva.Caption)
            Else
               If op10.value = True Then
                  labiva.Caption = CDbl(labsub.Caption) * 0.1
                  labtotf.Caption = CDbl(labsub.Caption) + CDbl(labiva.Caption)
               Else
                  labiva.Caption = 0
                  labtotf.Caption = CDbl(labsub.Caption) + CDbl(labiva.Caption)
               End If
            End If
            labiva.Caption = Format(labiva.Caption, "Standard")
            labtotf.Caption = Format(labtotf.Caption, "Standard")
            t_bol.Enabled = False
            Combo1.Enabled = False
            Combo2.Enabled = False
            mfec.Enabled = False
            t_codprod.SetFocus
            t_codprod.Text = ""
            t_pre.Text = ""
            t_cant.Text = ""
            labtot.Caption = 0
         
         End If
      Else
         data_linfac.Recordset.AddNew
         data_linfac.Recordset("id") = t_idit.Text
         data_linfac.Recordset("codbol") = t_bol.Text
         data_linfac.Recordset("fecha") = mfec.Text
         data_linfac.Recordset("codprod") = t_codprod.Text
         data_linfac.Recordset("coddesc") = labdesc.Caption
         data_linfac.Recordset("cant") = t_cant.Text
         data_linfac.Recordset("precuni") = CDbl(t_pre.Text)
         data_linfac.Recordset("totprod") = CDbl(labtot.Caption)
         data_linfac.Recordset("codcom") = t_codcom.Text
         If op10.value = True Then
            data_linfac.Recordset("opiva") = 1
         Else
            If op20.value = True Then
               data_linfac.Recordset("opiva") = 2
            Else
               data_linfac.Recordset("opiva") = 0
            End If
         End If
         data_linfac.Recordset.Update
         data_actitem.RecordSource = "Select * from stock where id =" & t_codprod.Text
         data_actitem.Refresh
         If data_actitem.Recordset.RecordCount > 0 Then
'            data_actitem.Recordset.Edit
            data_actitem.Recordset("actual") = data_actitem.Recordset("actual") + t_cant.Text
            If IsNull(data_actitem.Recordset("preuni")) = False Then
               If data_actitem.Recordset("preuni") = t_pre.Text Then
               Else
                  data_actitem.Recordset("preuni") = CDbl(t_pre.Text)
               End If
            Else
               data_actitem.Recordset("preuni") = CDbl(t_pre.Text)
            End If
            data_actitem.Recordset.Update
         End If
         List1.AddItem labdesc.Caption & "----" & t_codprod.Text & "--->" & t_idit.Text
         If labiva.Caption = "" Then
            labiva.Caption = 0
         End If
         If labtotf.Caption = "" Then
            labtotf.Caption = 0
         End If
         If labtot.Caption = "" Then
            labtot.Caption = 0
         End If
         If labsub.Caption = "" Then
            labsub.Caption = 0
         End If
         labsub.Caption = CDbl(labsub.Caption) + CDbl(labtot.Caption)
         If op20.value = True Then
            labiva.Caption = CDbl(labsub.Caption) * 0.22
            labtotf.Caption = CDbl(labsub.Caption) + CDbl(labiva.Caption)
         Else
            If op10.value = True Then
               labiva.Caption = CDbl(labsub.Caption) * 0.1
               labtotf.Caption = CDbl(labsub.Caption) + CDbl(labiva.Caption)
            Else
               labiva.Caption = 0
               labtotf.Caption = CDbl(labsub.Caption) + CDbl(labiva.Caption)
            End If
         End If
         labiva.Caption = Format(labiva.Caption, "Standard")
         labtotf.Caption = Format(labtotf.Caption, "Standard")
         t_bol.Enabled = False
         Combo1.Enabled = False
         Combo2.Enabled = False
         mfec.Enabled = False
         t_codprod.SetFocus
         t_codprod.Text = ""
         t_pre.Text = ""
         t_cant.Text = ""
         labtot.Caption = 0
      
      End If
   End If
Else

End If

End Sub

Private Sub b_grab_Click()

If t_bol.Text <> "" Then
   If mfec.Text <> "__/__/____" Then
      If t_id.Text <> "" Then
         If XAlta = 1 Then
            data_comp.RecordSource = "Select * from compras where id =" & t_id.Text
            data_comp.Refresh
            If data_comp.Recordset.RecordCount > 0 Then
'               data_comp.Recordset.Edit
               If labiva.Caption <> "" Then
                  data_comp.Recordset("totiva") = CDbl(labiva.Caption)
               Else
                  data_comp.Recordset("totiva") = 0
               End If
               If labtotf.Caption <> "" Then
                  data_comp.Recordset("totfac") = CDbl(labtotf.Caption)
               Else
                  data_comp.Recordset("totfac") = 0
               End If
               If labsub.Caption <> "" Then
                  data_comp.Recordset("totuni") = CDbl(labsub.Caption)
               Else
                  data_comp.Recordset("totuni") = 0
               End If
               data_comp.Recordset.Update
            Else
               MsgBox "No se encontró el encabezado de la factura, cancele factura y vuelva a crearla", vbCritical, "SAPP"
            End If
            XAlta = 0
            b_alta.Enabled = True
            b_mod.Enabled = True
            b_grab.Enabled = False
            b_canc.Enabled = False
            b_imp.Enabled = True
            b_elim.Enabled = True
            t_bol.Enabled = True
            Combo1.Enabled = True
            Combo2.Enabled = True
            mfec.Enabled = True
            borra_comp
            data_comp.Refresh
            data_bus.Refresh
            data_comp.Recordset.MoveLast
            iguala_comp
            Frame1.Enabled = False
         Else
            If XAlta = 3 Then
            Else
'               data_comp.Recordset.Edit
               data_comp.Recordset("totiva") = Xtotaiv
               data_comp.Recordset("totfac") = Xtotacom
               data_comp.Recordset("totuni") = Xsubtota
               data_comp.Recordset.Update
            End If
            XAlta = 0
            b_alta.Enabled = True
            b_mod.Enabled = True
            b_grab.Enabled = False
            b_canc.Enabled = False
            b_imp.Enabled = True
            b_elim.Enabled = True
            t_bol.Enabled = True
            Combo1.Enabled = True
            Combo2.Enabled = True
            mfec.Enabled = True
            borra_comp
            data_comp.Refresh
            data_bus.Refresh
            iguala_comp
            Frame1.Enabled = False
         End If
      End If
   End If
End If
   
End Sub



Private Sub b_imp_Click()
frm_infcomp.Show vbModal

End Sub

Private Sub b_mod_Click()
XAlta = 0

Frame1.Enabled = True
t_bol.SetFocus
b_alta.Enabled = False
b_mod.Enabled = False
b_grab.Enabled = True
b_canc.Enabled = True
b_elim.Enabled = False
b_imp.Enabled = False


End Sub

Private Sub cboiva_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbolab.SetFocus
End If

End Sub


Private Sub cbomon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboiva.SetFocus
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo2.SetFocus
End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfec.SetFocus
End If

End Sub

Private Sub Combo2_LostFocus()
If Combo1.ListIndex >= 0 Then
   data_lab.RecordSource = "Select * from abmdesp where obsmot ='" & Combo2.Text & "'"
   data_lab.Refresh
   If data_lab.Recordset.RecordCount > 0 Then
      t_codcom.Text = data_lab.Recordset("nro")
   Else
      MsgBox "No se encontró LABORATORIO/COMERCIO"
      Combo2.SetFocus
   End If
End If

End Sub

Private Sub Command1_Click()
frm_buscompit.Show vbModal

End Sub

Private Sub DBGrid1_DblClick()
t_codprod.Text = ""
t_idit.Text = ""
labdesc.Caption = ""
t_pre.Text = ""
t_cant.Text = ""
labiva.Caption = ""
labtotf.Caption = ""
List1.Clear
op10.value = False
op20.value = False
t_id.Text = ""
t_bol.Text = ""
mfec.Text = "__/__/____"
Combo1.ListIndex = -1
t_codcom.Text = ""
Combo2.ListIndex = -1

data_comp.RecordSource = "Select * from compras where id =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 5)
data_comp.Refresh
If data_comp.Recordset.RecordCount > 0 Then
    data_linfac.RecordSource = "Select * from lineascomp where codbol =" & data_comp.Recordset("nrobol") & " and codcom =" & data_comp.Recordset("codcomer")
    data_linfac.Refresh
    If IsNull(data_comp.Recordset("id")) = False Then
       t_id.Text = data_comp.Recordset("id")
    Else
       t_id.Text = 1
    End If
    If IsNull(data_comp.Recordset("nrobol")) = False Then
       t_bol.Text = data_comp.Recordset("nrobol")
    Else
       t_bol.Text = 1
    End If
    If IsNull(data_comp.Recordset("fecha")) = False Then
       mfec.Text = data_comp.Recordset("fecha")
    Else
       mfec.Text = "__/__/____"
    End If
    If IsNull(data_comp.Recordset("tipobol")) = False Then
       Combo1.ListIndex = data_comp.Recordset("tipobol")
    Else
       Combo1.ListIndex = -1
    End If
    If IsNull(data_comp.Recordset("codcomer")) = False Then
       t_codcom.Text = data_comp.Recordset("codcomer")
    Else
       t_codcom.Text = ""
    End If
    If IsNull(data_comp.Recordset("nomcomer")) = False Then
       Combo2.Text = data_comp.Recordset("nomcomer")
    Else
       Combo2.Text = ""
    End If
    If IsNull(data_comp.Recordset("totiva")) = False Then
       labiva.Caption = data_comp.Recordset("totiva")
    Else
       labiva.Caption = 0
    End If
    If IsNull(data_comp.Recordset("totfac")) = False Then
       labtotf.Caption = data_comp.Recordset("totfac")
    Else
       labtotf.Caption = 0
    End If
    
    If data_linfac.Recordset.RecordCount > 0 Then
       data_linfac.Recordset.MoveFirst
       List1.Clear
       Do While Not data_linfac.Recordset.EOF
          If IsNull(data_linfac.Recordset("codprod")) = False Then
             t_codprod.Text = data_linfac.Recordset("codprod")
          Else
             t_codprod.Text = 0
          End If
          If IsNull(data_linfac.Recordset("id")) = False Then
             t_idit.Text = data_linfac.Recordset("id")
          Else
             t_idit.Text = 0
          End If
          If IsNull(data_linfac.Recordset("coddesc")) = False Then
             labdesc.Caption = data_linfac.Recordset("coddesc")
          Else
             labdesc.Caption = "S/D"
          End If
          If IsNull(data_linfac.Recordset("precuni")) = False Then
             t_pre.Text = data_linfac.Recordset("precuni")
          Else
             t_pre.Text = 0
          End If
          If IsNull(data_linfac.Recordset("opiva")) = False Then
             If data_linfac.Recordset("opiva") = 1 Then
                op10.value = True
                op20.value = False
             Else
                If data_linfac.Recordset("opiva") = 2 Then
                   op20.value = True
                   op10.value = False
                Else
                   op20.value = False
                   op10.value = False
                End If
             End If
          Else
             op20.value = False
             op10.value = False
          End If
          If IsNull(data_linfac.Recordset("cant")) = False Then
             t_cant.Text = data_linfac.Recordset("cant")
          Else
             t_cant.Text = 0
          End If
          List1.AddItem labdesc.Caption & "----" & t_codprod.Text & "--->" & t_idit.Text
          data_linfac.Recordset.MoveNext
       Loop
    End If
Else
    MsgBox "No encontrado!!", vbCritical
End If

End Sub

Private Sub Form_Load()
data_item.ConnectionString = "dsn=" & Xconexrmt
data_comp.ConnectionString = "dsn=" & Xconexrmt
data_bus.ConnectionString = "dsn=" & Xconexrmt
data_actitem.ConnectionString = "dsn=" & Xconexrmt


data_bus.RecordSource = "Select * from compras order by fecha DESC"
data_bus.Refresh

'data_item.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_item.RecordSource = "stock"
'data_item.Refresh

'data_comp.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_comp.RecordSource = "Select * from compras where id =" & 850
data_comp.Refresh

data_lab.ConnectionString = "dsn=" & Xconexrmt
data_lab.RecordSource = "Select * from abmdesp where base =" & 0 & " order by obsmot"
data_lab.Refresh

'data_actitem.RecordSource = "stock"
'data_actitem.Refresh

data_linfac.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_linfac2.ConnectionString = "dsn=" & Xconexrmt & ";"

'data_linfac.RecordSource = "lineascomp"
'data_linfac.Refresh

If data_lab.Recordset.RecordCount > 0 Then
   data_lab.Recordset.MoveFirst
   Do While Not data_lab.Recordset.EOF
      If IsNull(data_lab.Recordset("obsmot")) = False Then
         Combo2.AddItem data_lab.Recordset("obsmot")
      End If
      data_lab.Recordset.MoveNext
   Loop
Else
   Combo2.AddItem "SIN LABORATORIOS"
End If

DBGrid1.Rows = 2
DBGrid1.Cols = 6
DBGrid1.TextMatrix(0, 0) = "FECHA"
DBGrid1.ColWidth(0) = 1500
DBGrid1.TextMatrix(0, 1) = "NRO.BOLETA"
DBGrid1.ColWidth(1) = 1500
DBGrid1.TextMatrix(0, 2) = "COMERCIO"
DBGrid1.ColWidth(2) = 2900
DBGrid1.TextMatrix(0, 3) = "TIPO.DOC"
DBGrid1.ColWidth(3) = 1200
DBGrid1.TextMatrix(0, 4) = "TOTAL FACT."
DBGrid1.ColWidth(4) = 1900
DBGrid1.TextMatrix(0, 5) = "ID"
DBGrid1.ColWidth(5) = 1200

Dim Xcann As Integer
Xcann = 1
If data_bus.Recordset.RecordCount > 0 Then
    data_bus.Recordset.MoveFirst
    Do While Not data_bus.Recordset.EOF
       If IsNull(data_bus.Recordset("fecha")) = False Then
          DBGrid1.TextMatrix(Xcann, 0) = data_bus.Recordset("fecha")
       End If
       If IsNull(data_bus.Recordset("nrobol")) = False Then
          DBGrid1.TextMatrix(Xcann, 1) = data_bus.Recordset("nrobol")
       End If
       If IsNull(data_bus.Recordset("nomcomer")) = False Then
          DBGrid1.TextMatrix(Xcann, 2) = data_bus.Recordset("nomcomer")
       End If
       If IsNull(data_bus.Recordset("tipobold")) = False Then
          DBGrid1.TextMatrix(Xcann, 3) = data_bus.Recordset("tipobold")
       End If
       If IsNull(data_bus.Recordset("totfac")) = False Then
          DBGrid1.TextMatrix(Xcann, 4) = data_bus.Recordset("totfac")
       End If
       If IsNull(data_bus.Recordset("id")) = False Then
          DBGrid1.TextMatrix(Xcann, 5) = data_bus.Recordset("id")
       End If
       DBGrid1.Rows = DBGrid1.Rows + 1
       data_bus.Recordset.MoveNext
       Xcann = Xcann + 1
    Loop
End If


End Sub

Private Sub mfc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_precu.SetFocus
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

Private Sub List1_DblClick()
Dim Xtext, Xcaden As String
Dim Xlargo, Xcuen, Xidd As Long
Dim Xbannn As Integer
Xbannn = 0
Xtext = List1.List(List1.ListIndex)
Xlargo = Len(Trim(Xtext))
Xcuen = 0
For Xcuen = 1 To Xlargo
    If Xbannn = 1 Then
       Xcaden = Xcaden & Mid(Xtext, Xcuen, 1)
    End If
    If Mid(Xtext, Xcuen, 1) = ">" Then
       Xbannn = 1
    End If
Next
Xidd = Val(Trim(Xcaden))
'data_linfac.Recordset.FindFirst "id =" & Xidd
data_linfac.RecordSource = "Select * from lineascomp where id =" & Xidd
data_linfac.Refresh
If data_linfac.Recordset.RecordCount > 0 Then
   t_codprod.Text = data_linfac.Recordset("codprod")
   t_idit.Text = data_linfac.Recordset("id")
   labdesc.Caption = data_linfac.Recordset("coddesc")
   t_pre.Text = data_linfac.Recordset("precuni")
   t_cant.Text = data_linfac.Recordset("cant")
   t_cantant.Text = t_cant.Text
   labtot.Caption = data_linfac.Recordset("totprod")
   If IsNull(data_linfac.Recordset("opiva")) = False Then
      If data_linfac.Recordset("opiva") = 1 Then
         op10.value = True
         op20.value = False
      Else
         If data_linfac.Recordset("opiva") = 2 Then
            op10.value = False
            op20.value = True
         Else
            op10.value = False
            op20.value = False
         End If
      End If
   Else
      op10.value = False
      op20.value = False
   End If
Else
   MsgBox "No se encuentra el producto en la factura, VERIFIQUE!!"
   t_codprod.Text = ""
   t_idit.Text = ""
   labdesc.Caption = ""
   t_pre.Text = ""
   t_cant.Text = ""
   labtot.Caption = ""
   op10.value = False
   op20.value = False
End If

Xbannn = 0

End Sub

Private Sub mfec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_codprod.SetFocus
End If

End Sub

Private Sub t_b_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   DBGrid1.SetFocus
   If t_b.Text = "" Then
   Else
      If Option1.value = True Then
         data_bus.RecordSource = "Select * from compras where fecha >='" & Format(t_b.Text, "yyyy-mm-dd") & "' order by fecha"
         data_bus.Refresh
      Else
         If Option2.value = True Then
            data_bus.RecordSource = "Select * from compras where nrobol >=" & t_b.Text
            data_bus.Refresh
         End If
      End If
   End If
   DBGrid1.Rows = 2
   DBGrid1.Cols = 6
   DBGrid1.TextMatrix(0, 0) = "FECHA"
   DBGrid1.ColWidth(0) = 1500
   DBGrid1.TextMatrix(0, 1) = "NRO.BOLETA"
   DBGrid1.ColWidth(1) = 1500
   DBGrid1.TextMatrix(0, 2) = "COMERCIO"
   DBGrid1.ColWidth(2) = 2900
   DBGrid1.TextMatrix(0, 3) = "TIPO.DOC"
   DBGrid1.ColWidth(3) = 1200
   DBGrid1.TextMatrix(0, 4) = "TOTAL FACT."
   DBGrid1.ColWidth(4) = 1900
   DBGrid1.TextMatrix(0, 5) = "ID"
   DBGrid1.ColWidth(5) = 1200
    
   Dim Xcann As Integer
   Xcann = 1
   If data_bus.Recordset.RecordCount > 0 Then
      data_bus.Recordset.MoveFirst
      Do While Not data_bus.Recordset.EOF
         If IsNull(data_bus.Recordset("fecha")) = False Then
             DBGrid1.TextMatrix(Xcann, 0) = data_bus.Recordset("fecha")
          End If
          If IsNull(data_bus.Recordset("nrobol")) = False Then
             DBGrid1.TextMatrix(Xcann, 1) = data_bus.Recordset("nrobol")
          End If
          If IsNull(data_bus.Recordset("nomcomer")) = False Then
             DBGrid1.TextMatrix(Xcann, 2) = data_bus.Recordset("nomcomer")
          End If
          If IsNull(data_bus.Recordset("tipobold")) = False Then
             DBGrid1.TextMatrix(Xcann, 3) = data_bus.Recordset("tipobold")
          End If
          If IsNull(data_bus.Recordset("totfac")) = False Then
             DBGrid1.TextMatrix(Xcann, 4) = data_bus.Recordset("totfac")
          End If
          If IsNull(data_bus.Recordset("id")) = False Then
             DBGrid1.TextMatrix(Xcann, 5) = data_bus.Recordset("id")
          End If
          DBGrid1.Rows = DBGrid1.Rows + 1
          data_bus.Recordset.MoveNext
          Xcann = Xcann + 1
       Loop
   End If
End If

End Sub

Private Sub t_bol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub

Private Sub t_cant_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_git.SetFocus
End If

End Sub

Private Sub t_cod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfc.SetFocus
End If

End Sub

Public Function iguala_comp()
Xtotacom = 0
Xtotaiv = 0
Xsubtota = 0

If IsNull(data_comp.Recordset("id")) = False Then
   t_id.Text = data_comp.Recordset("id")
Else
   t_id.Text = 1
End If
If IsNull(data_comp.Recordset("nrobol")) = False Then
   t_bol.Text = data_comp.Recordset("nrobol")
Else
   t_bol.Text = 1
End If

data_linfac.RecordSource = "Select * from lineascomp where codbol =" & t_bol.Text & " order by id"
data_linfac.Refresh

If IsNull(data_comp.Recordset("fecha")) = False Then
   mfec.Text = data_comp.Recordset("fecha")
Else
   mfec.Text = "__/__/____"
End If
If IsNull(data_comp.Recordset("tipobol")) = False Then
   Combo1.ListIndex = data_comp.Recordset("tipobol")
Else
   Combo1.ListIndex = -1
End If
If IsNull(data_comp.Recordset("codcomer")) = False Then
   t_codcom.Text = data_comp.Recordset("codcomer")
Else
   t_codcom.Text = ""
End If
If IsNull(data_comp.Recordset("nomcomer")) = False Then
   Combo2.Text = data_comp.Recordset("nomcomer")
Else
   Combo2.Text = ""
End If
If IsNull(data_comp.Recordset("totiva")) = False Then
   labiva.Caption = Format(data_comp.Recordset("totiva"), "Standard")
Else
   labiva.Caption = 0
End If
If IsNull(data_comp.Recordset("totfac")) = False Then
   labtotf.Caption = Format(data_comp.Recordset("totfac"), "Standard")
Else
   labtotf.Caption = 0
End If
If IsNull(data_comp.Recordset("totuni")) = False Then
   labsub.Caption = Format(data_comp.Recordset("totuni"), "Standard")
Else
   labsub.Caption = 0
End If

If data_linfac.Recordset.RecordCount > 0 Then
   data_linfac.Recordset.MoveFirst
   List1.Clear
   Do While Not data_linfac.Recordset.EOF
      If IsNull(data_linfac.Recordset("codprod")) = False Then
         t_codprod.Text = data_linfac.Recordset("codprod")
      Else
         t_codprod.Text = 0
      End If
      If IsNull(data_linfac.Recordset("id")) = False Then
         t_idit.Text = data_linfac.Recordset("id")
      Else
         t_idit.Text = 0
      End If
      If IsNull(data_linfac.Recordset("coddesc")) = False Then
         labdesc.Caption = data_linfac.Recordset("coddesc")
      Else
         labdesc.Caption = "S/D"
      End If
      If IsNull(data_linfac.Recordset("precuni")) = False Then
         t_pre.Text = data_linfac.Recordset("precuni")
      Else
         t_pre.Text = 0
      End If
      If IsNull(data_linfac.Recordset("cant")) = False Then
         t_cant.Text = data_linfac.Recordset("cant")
      Else
         t_cant.Text = 0
      End If
      List1.AddItem labdesc.Caption & "----" & t_codprod.Text & "--->" & t_idit.Text
      Xsubtota = Xsubtota + data_linfac.Recordset("totprod")
      If IsNull(data_linfac.Recordset("opiva")) = False Then
         If data_linfac.Recordset("opiva") = 1 Then
            Xtotaiv = Xsubtota * 0.1
         Else
            If data_linfac.Recordset("opiva") = 2 Then
               Xtotaiv = Xsubtota * 0.22
            Else
               Xtotaiv = 0
            End If
         End If
      Else
         Xtotaiv = 0
      End If
      data_linfac.Recordset.MoveNext
   Loop
   Xtotacom = Xsubtota + Xtotaiv
   t_codprod.Text = ""
   t_idit.Text = ""
   labdesc.Caption = ""
   t_pre.Text = ""
   t_cant.Text = ""
Else
   t_codprod.Text = ""
   t_idit.Text = ""
   labdesc.Caption = ""
   t_pre.Text = ""
   t_cant.Text = ""
   labiva.Caption = ""
   labtotf.Caption = ""
   List1.Clear
   MsgBox "No tiene ingresado ningun ITEM"
End If

End Function

Private Sub t_precu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_cant.SetFocus
   If t_precu.Text <> "" Then
      t_precu.Text = Format(t_precu.Text, "Standard")
   Else
      t_precu.Text = 0
   End If
End If

End Sub

Private Sub t_cant_LostFocus()
If t_pre.Text <> "" Then
   If t_cant.Text <> "" Then
      labtot.Caption = CDbl(t_cant.Text) * CDbl(t_pre.Text)
   Else
   End If
End If

End Sub

Private Sub t_codprod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If t_codprod.Text = "" Then
      b_grab.SetFocus
   Else
      t_pre.SetFocus
   End If
End If

End Sub

Private Sub t_codprod_LostFocus()
If t_codprod.Text <> "" Then
   data_item.RecordSource = "Select * from stock where id =" & t_codprod.Text
   data_item.Refresh
   If data_item.Recordset.RecordCount > 0 Then
      labdesc.Caption = data_item.Recordset("descrip")
      t_pre.Text = data_item.Recordset("preuni")
      grupo.Caption = data_item.Recordset("grupo")
   Else
      labdesc.Caption = ""
      t_pre.Text = 0
      grupo.Caption = 0
      MsgBox "Código de ITEM no encontrado", vbCritical, "STOCK"
   End If
End If

End Sub

Private Sub t_pre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_cant.SetFocus
End If

End Sub

Private Sub t_pre_LostFocus()
If t_pre.Text <> "" Then
   If t_cant.Text <> "" Then
      t_pre.Text = Format(t_pre.Text, "Standard")
      labtot.Caption = t_cant.Text * t_pre.Text
      labtot.Caption = Format(labtot.Caption, "Standard")
   Else
   End If
End If

End Sub
