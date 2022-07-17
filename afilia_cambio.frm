VERSION 5.00
Begin VB.Form afilia_cambio 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Cambio de categoría"
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione la Categoría a la cual se va a cambiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.Data valores_selec 
         Caption         =   "valores_selec"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1200
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1560
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ok"
         Height          =   615
         Left            =   2280
         Picture         =   "afilia_cambio.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   360
         ItemData        =   "afilia_cambio.frx":058A
         Left            =   240
         List            =   "afilia_cambio.frx":059A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   5055
      End
   End
End
Attribute VB_Name = "afilia_cambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Valor_af As Long
Valor_af = 0
If Combo1.ListIndex >= 0 Then
   Data1.RecordSource = "select clientes.cl_cedula,clientes.estado,clientes.cl_codconv," & _
   "convenio.cnv_codigo,convenio.cnv_precio,convenio.cnv_cant_r from clientes inner join " & _
   "convenio on clientes.cl_codconv=convenio.cnv_codigo where clientes.cl_cedula =" & frm_afilia.t_ced.Text & " and estado =" & 1
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      valores_selec.RecordSource = "select * from afiliaciones_categ where descrip ='" & Combo1.Text & "'"
      valores_selec.Refresh
      If valores_selec.Recordset.RecordCount > 0 Then
         If Data1.Recordset("cnv_precio") >= valores_selec.Recordset("valor") Then
            Valor_af = 0
         Else
            Valor_af = valores_selec.Recordset("valor") - Data1.Recordset("cnv_precio")
         End If
      Else
         Valor_af = 0
      End If
      MsgBox "Valor de la afiliación: $." & Format(Valor_af, "Standard")
      frm_afilia.t_valor.Text = Val(Valor_af)
      frm_afilia.labcatnomsol.Caption = valores_selec.Recordset("catnom")
      frm_afilia.labcatcodsol.Caption = valores_selec.Recordset("catsapp")
      If IsNull(valores_selec.Recordset("catrealcod")) = False Then
         frm_afilia.labcatreal.Caption = valores_selec.Recordset("catrealcod")
         frm_afilia.labcatrealdes.Caption = valores_selec.Recordset("catrealdes")
      Else
         frm_afilia.labcatreal.Caption = ""
         frm_afilia.labcatrealdes.Caption = ""
      End If
      Unload Me
   Else
      Combo1.ListIndex = -1
      MsgBox "No se encuentra el socio, VERIFIQUE!", vbCritical
      Unload Me
   End If
Else
   MsgBox "Seleccione una categoría.", vbCritical
End If

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=sappnew;"
valores_selec.Connect = "odbc;dsn=sappnew;"

End Sub
