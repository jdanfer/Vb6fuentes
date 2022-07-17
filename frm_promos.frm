VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_promos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Promociones"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8190
   Icon            =   "frm_promos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8190
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_promos.frx":058A
      Height          =   1695
      Left            =   240
      OleObjectBlob   =   "frm_promos.frx":059E
      TabIndex        =   18
      Top             =   4200
      Width           =   7575
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      Picture         =   "frm_promos.frx":0F81
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Cancelar"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      Picture         =   "frm_promos.frx":150B
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Grabar datos"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton b_edit 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_promos.frx":1A95
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Editar registro seleccionado"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_promos.frx":201F
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Nuevo registro"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Tabla de promociones"
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
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2760
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.TextBox t_pesos 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5520
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox t_porc 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t_descrip 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   4
         Top             =   960
         Width           =   4695
      End
      Begin VB.TextBox t_cod 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.Label labfec 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4800
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vigencia Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vigencia Desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Descuento en $."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3720
         TabIndex        =   7
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Descuento en %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_promos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_alta_Click()
Frame1.Enabled = True
t_cod.Text = ""
t_descrip.Text = ""
labfec.Caption = Format(Date, "dd/mm/yyyy")
t_porc.Text = ""
t_pesos.Text = ""
md.Text = "__/__/____"
mh.Text = "__/__/____"
t_descrip.SetFocus
b_alta.Enabled = False
b_edit.Enabled = False
b_graba.Enabled = True
b_cance.Enabled = True
XAlta = 1

End Sub

Private Sub b_cance_Click()
t_cod.Text = ""
t_descrip.Text = ""
labfec.Caption = Format(Date, "dd/mm/yyyy")
t_porc.Text = ""
t_pesos.Text = ""
md.Text = "__/__/____"
mh.Text = "__/__/____"
   
   Frame1.Enabled = False
   b_alta.Enabled = True
   b_edit.Enabled = True
   b_graba.Enabled = False
   b_cance.Enabled = False
   t_descrip.SetFocus

End Sub

Private Sub b_edit_Click()

If t_cod.Text <> "" Then
   Frame1.Enabled = True
   b_alta.Enabled = False
   b_edit.Enabled = False
   b_graba.Enabled = True
   b_cance.Enabled = True
   t_descrip.SetFocus
Else
   MsgBox "No seleccionó registro"

End If

End Sub

Private Sub b_graba_Click()

If t_descrip.Text <> "" Then
   If XAlta = 1 Then
      Data2.RecordSource = "select * from promocion_gpo where descrip ='" & t_descrip.Text & "'"
      Data2.Refresh
      If Data2.Recordset.RecordCount > 0 Then
         MsgBox "Ya existe una promoción con esta descripción", vbInformation
      Else
         Data1.Recordset.AddNew
         Data1.Recordset("descrip") = t_descrip.Text
         Data1.Recordset("fecha") = Date
         If md.Text <> "__/__/____" Then
            Data1.Recordset("desde") = Format(md.Text, "dd/mm/yyyy")
         End If
         If mh.Text <> "__/__/____" Then
            Data1.Recordset("hasta") = Format(mh.Text, "dd/mm/yyyy")
         End If
         If t_porc.Text = "" Then
            t_porc.Text = 0
         End If
         If t_pesos.Text = "" Then
            t_pesos.Text = 0
         End If
         Data1.Recordset("descu_por") = t_porc.Text
         Data1.Recordset("descu_imp") = t_pesos.Text
         Data1.Recordset.Update
         Data1.Refresh
         t_cod.Text = ""
         t_descrip.Text = ""
         labfec.Caption = ""
         t_porc.Text = ""
         t_pesos.Text = ""
         md.Text = "__/__/____"
         mh.Text = "__/__/____"
         b_alta.Enabled = True
         b_edit.Enabled = True
         b_graba.Enabled = False
         b_cance.Enabled = False
         Frame1.Enabled = False
         XAlta = 0
      End If
   Else
      Data1.Recordset.Edit
      Data1.Recordset("descrip") = t_descrip.Text
      Data1.Recordset("fecha") = Date
      If md.Text <> "__/__/____" Then
         Data1.Recordset("desde") = Format(md.Text, "dd/mm/yyyy")
      End If
      If mh.Text <> "__/__/____" Then
         Data1.Recordset("hasta") = Format(mh.Text, "dd/mm/yyyy")
      End If
      If t_porc.Text = "" Then
         t_porc.Text = 0
      End If
      If t_pesos.Text = "" Then
         t_pesos.Text = 0
      End If
      Data1.Recordset("descu_por") = t_porc.Text
      Data1.Recordset("descu_imp") = t_pesos.Text
      Data1.Recordset.Update
      Data1.Refresh
      b_alta.Enabled = True
      b_edit.Enabled = True
      b_graba.Enabled = False
      b_cance.Enabled = False
      t_cod.Text = ""
      t_descrip.Text = ""
      labfec.Caption = ""
      t_porc.Text = ""
      t_pesos.Text = ""
      md.Text = "__/__/____"
      mh.Text = "__/__/____"
      Frame1.Enabled = False
      XAlta = 0
   End If
Else
   MsgBox "No se ingresó descripción"
   
End If

End Sub

Private Sub DBGrid1_Click()
t_cod.Text = Data1.Recordset("id")

t_descrip.Text = Data1.Recordset("descrip")
If IsNull(Data1.Recordset("desde")) = False Then
   md.Text = Data1.Recordset("desde")
Else
   md.Text = "__/__/____"
End If
If IsNull(Data1.Recordset("hasta")) = False Then
   md.Text = Data1.Recordset("desde")
Else
   md.Text = "__/__/____"
End If
labfec.Caption = Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
If IsNull(Data1.Recordset("descu_por")) = False Then
   t_porc.Text = Data1.Recordset("descu_por")
Else
   t_porc.Text = ""
End If
If IsNull(Data1.Recordset("descu_imp")) = False Then
   t_pesos.Text = Data1.Recordset("descu_imp")
Else
   t_pesos.Text = ""
End If

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=sappnew;"
Data1.RecordSource = "promocion_gpo"
Data1.Refresh
Data2.Connect = "odbc;dsn=sappnew;"

End Sub

