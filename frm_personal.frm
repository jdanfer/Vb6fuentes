VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_personal 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos para tarjetas BROU y Caja Profesional"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7245
   Icon            =   "frm_personal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7245
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Procesar a CJPPU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Solicitar Tarjetas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6480
      Picture         =   "frm_personal.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Salir"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton b_bus 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      Picture         =   "frm_personal.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Buscar"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton b_can 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      Picture         =   "frm_personal.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cancelar acción"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton b_gra 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      Picture         =   "frm_personal.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Grabar datos"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton b_mod 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_personal.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Editar registro"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton b_nue 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_personal.frx":1FF4
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Nuevo registro"
      Top             =   4920
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos del personal"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6735
      Begin VB.TextBox txt_fpag 
         Alignment       =   1  'Right Justify
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
         Left            =   5640
         TabIndex        =   32
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox txt_rlab 
         Alignment       =   1  'Right Justify
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
         Left            =   1920
         TabIndex        =   30
         Top             =   3720
         Width           =   975
      End
      Begin VB.Data data_per 
         Caption         =   "data_per"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "tarjbrou"
         Top             =   3120
         Visible         =   0   'False
         Width           =   3135
      End
      Begin MSMask.MaskEdBox mfing 
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   3360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_dpto 
         Height          =   285
         Left            =   5400
         MaxLength       =   20
         TabIndex        =   19
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txt_loc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   17
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox txt_codp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         MaxLength       =   5
         TabIndex        =   15
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txt_tel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txt_dir 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   60
         TabIndex        =   11
         Top             =   1920
         Width           =   4575
      End
      Begin VB.TextBox txt_ape2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txt_ape1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txt_nom2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         MaxLength       =   15
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txt_nom1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   5
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txt_cod 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txt_ced 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   7
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Código Forma de Pago= 1,2,3,4,5,6 o 7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   34
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Códigos de relación laboral= 1,2,3,4 o 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   4080
         Width           =   3135
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cod.F.Pago:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   31
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cod.Relac.Lab."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Fecha Ingreso:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Dpto."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Localidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cod.P."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   14
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Apellidos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nombres:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "CEDULA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   960
      Picture         =   "frm_personal.frx":257E
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1575
   End
End
Attribute VB_Name = "frm_personal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text4_Change()

End Sub

Private Sub b_bus_Click()
frm_busper.Show vbModal

End Sub

Private Sub b_can_Click()
If XAlta = 1 Then
   data_per.Recordset.CancelUpdate
   Frame1.Enabled = False
   b_gra.Enabled = False
   b_can.Enabled = False
   b_nue.Enabled = True
   b_mod.Enabled = True
   b_bus.Enabled = True
   XAlta = 0
   borrar_tar
   igualaper
Else
   Frame1.Enabled = False
   b_gra.Enabled = False
   b_can.Enabled = False
   b_nue.Enabled = True
   b_mod.Enabled = True
   b_bus.Enabled = True
   XAlta = 0
   borrar_tar
   igualaper
End If

End Sub

Private Sub b_gra_Click()
If txt_ced.Text <> "" Then
   If XAlta = 1 Then
      data_per.Recordset("cedula") = txt_ced.Text
      data_per.Recordset("codver") = txt_cod.Text
      data_per.Recordset("nom1") = UCase(txt_nom1.Text)
      data_per.Recordset("nom2") = UCase(txt_nom2.Text)
      data_per.Recordset("ape1") = UCase(txt_ape1.Text)
      data_per.Recordset("ape2") = UCase(txt_ape2.Text)
      data_per.Recordset("calle") = txt_dir.Text
      If txt_codp.Text = "" Then
         txt_codp.Text = 0
      End If
      data_per.Recordset("codpos") = txt_codp.Text
      data_per.Recordset("telpart") = txt_tel.Text
      data_per.Recordset("localid") = txt_loc.Text
      data_per.Recordset("depto") = txt_dpto.Text
      If mfing.Text <> "__/__/____" Then
         data_per.Recordset("fecing") = Format(mfing.Text, "dd/mm/yyyy")
      End If
      If txt_rlab.Text <> "" Then
         data_per.Recordset("rellab") = txt_rlab.Text
      End If
      If txt_fpag.Text <> "" Then
         data_per.Recordset("forpag") = txt_fpag.Text
      End If
      data_per.Recordset.Update
      Frame1.Enabled = False
      b_gra.Enabled = False
      b_can.Enabled = False
      b_nue.Enabled = True
      b_mod.Enabled = True
      b_bus.Enabled = True
      XAlta = 0
      borrar_tar
      igualaper
   Else
      data_per.Recordset.Edit
      data_per.Recordset("cedula") = txt_ced.Text
      data_per.Recordset("codver") = txt_cod.Text
      data_per.Recordset("nom1") = UCase(txt_nom1.Text)
      data_per.Recordset("nom2") = UCase(txt_nom2.Text)
      data_per.Recordset("ape1") = UCase(txt_ape1.Text)
      data_per.Recordset("ape2") = UCase(txt_ape2.Text)
      data_per.Recordset("calle") = txt_dir.Text
      data_per.Recordset("codpos") = txt_codp.Text
      data_per.Recordset("telpart") = txt_tel.Text
      data_per.Recordset("localid") = txt_loc.Text
      data_per.Recordset("depto") = txt_dpto.Text
      If mfing.Text <> "__/__/____" Then
         data_per.Recordset("fecing") = Format(mfing.Text, "dd/mm/yyyy")
      End If
      If txt_rlab.Text <> "" Then
         data_per.Recordset("rellab") = txt_rlab.Text
      End If
      If txt_fpag.Text <> "" Then
         data_per.Recordset("forpag") = txt_fpag.Text
      End If
      data_per.Recordset.Update
      Frame1.Enabled = False
      b_gra.Enabled = False
      b_can.Enabled = False
      b_nue.Enabled = True
      b_mod.Enabled = True
      b_bus.Enabled = True
      XAlta = 0
      borrar_tar
      igualaper
   End If
End If

End Sub

Private Sub b_mod_Click()
Frame1.Enabled = True
b_gra.Enabled = True
b_can.Enabled = True
b_nue.Enabled = False
b_mod.Enabled = False
b_bus.Enabled = False
XAlta = 0

End Sub

Private Sub b_nue_Click()
Frame1.Enabled = True
b_gra.Enabled = True
b_can.Enabled = True
b_nue.Enabled = False
b_mod.Enabled = False
b_bus.Enabled = False
XAlta = 1
data_per.Recordset.AddNew

borrar_tar
txt_ced.SetFocus

End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
Dim Micadtar As String
Dim xcuenta As Integer
Dim Xdesdef As String

Xdesdef = InputBox("Ingrese desde que fecha de ingreso? (Ej.01/10/2008)...:", "Selección de datos")
If Xdesdef <> "" Then
   frm_personal.MousePointer = 11
   data_per.RecordSource = "select * from tarjbrou where fecing >=#" & Format(Xdesdef, "yyyy/mm/dd") & "#"
   data_per.Refresh
   If data_per.Recordset.RecordCount > 0 Then
      data_per.Recordset.MoveFirst
      Open App.Path & "\nomina.txt" For Output As #1
      Do While Not data_per.Recordset.EOF
         Micadtar = "SJ073000385"
         If Len(Trim(Str(data_per.Recordset("cedula")))) < 7 Then
            Micadtar = Micadtar + "0" + Trim(Str(data_per.Recordset("cedula"))) + Trim(Str(data_per.Recordset("codver")))
            Micadtar = Micadtar + "0000000"
         Else
            Micadtar = Micadtar + Trim(Str(data_per.Recordset("cedula"))) + Trim(Str(data_per.Recordset("codver")))
            Micadtar = Micadtar + "0000000"
         End If
         Micadtar = Micadtar + "9800002845"
         If Len(Trim(Str(data_per.Recordset("cedula")))) < 7 Then
            Micadtar = Micadtar + "00000" + Trim(Str(data_per.Recordset("cedula"))) + Trim(Str(data_per.Recordset("codver")))
         Else
            Micadtar = Micadtar + "0000" + Trim(Str(data_per.Recordset("cedula"))) + Trim(Str(data_per.Recordset("codver")))
         End If
         Micadtar = Micadtar + Mid(data_per.Recordset("nom1"), 1, 15)
         xcuenta = Len(data_per.Recordset("nom1"))
         xcuenta = xcuenta + 1
         For xcuenta = xcuenta To 15
             Micadtar = Micadtar + " "
         Next
         If IsNull(data_per.Recordset("nom2")) = False Then
            Micadtar = Micadtar + Mid(data_per.Recordset("nom2"), 1, 15)
            xcuenta = Len(data_per.Recordset("nom2"))
            xcuenta = xcuenta + 1
            For xcuenta = xcuenta To 15
                Micadtar = Micadtar + " "
            Next
         Else
            Micadtar = Micadtar + "               "
         End If
         Micadtar = Micadtar + Mid(data_per.Recordset("ape1"), 1, 15)
         xcuenta = Len(data_per.Recordset("ape1"))
         xcuenta = xcuenta + 1
         For xcuenta = xcuenta To 15
             Micadtar = Micadtar + " "
         Next
         If IsNull(data_per.Recordset("ape2")) = False Then
            Micadtar = Micadtar + Mid(data_per.Recordset("ape2"), 1, 15)
            xcuenta = Len(data_per.Recordset("ape2"))
            xcuenta = xcuenta + 1
            For xcuenta = xcuenta To 15
                Micadtar = Micadtar + " "
            Next
         Else
            Micadtar = Micadtar + "               "
         End If
         Micadtar = Micadtar + "000000000000000"
         Micadtar = Micadtar + Mid(data_per.Recordset("calle"), 1, 60)
         xcuenta = Len(data_per.Recordset("calle"))
         xcuenta = xcuenta + 1
         For xcuenta = xcuenta To 60
             Micadtar = Micadtar + " "
         Next
         If Len(Trim(Str(data_per.Recordset("codpos")))) = 5 Then
            Micadtar = Micadtar + Trim(data_per.Recordset("codpos"))
         Else
            If Len(Trim(Str(data_per.Recordset("codpos")))) = 4 Then
               Micadtar = Micadtar + "0" + Trim(data_per.Recordset("codpos"))
            Else
               Micadtar = Micadtar + "00000"
            End If
         End If
         If IsNull(data_per.Recordset("telpart")) = False Then
            Micadtar = Micadtar + Mid(data_per.Recordset("telpart"), 1, 20)
            xcuenta = Len(data_per.Recordset("telpart"))
            xcuenta = xcuenta + 1
            For xcuenta = xcuenta To 20
                Micadtar = Micadtar + " "
            Next
         Else
            Micadtar = Micadtar + "                    "
         End If
         If IsNull(data_per.Recordset("localid")) = False Then
            Micadtar = Micadtar + Mid(data_per.Recordset("localid"), 1, 20)
            xcuenta = Len(data_per.Recordset("localid"))
            xcuenta = xcuenta + 1
            For xcuenta = xcuenta To 20
                Micadtar = Micadtar + " "
            Next
         Else
            Micadtar = Micadtar + "                    "
         End If
         Micadtar = Micadtar + Mid(data_per.Recordset("depto"), 1, 20)
         xcuenta = Len(data_per.Recordset("depto"))
         xcuenta = xcuenta + 1
         For xcuenta = xcuenta To 20
             Micadtar = Micadtar + " "
         Next
         Micadtar = Micadtar + "845"
         Print #1, Trim(Micadtar)
         data_per.Recordset.MoveNext
      Loop
      MsgBox "Proceso terminado, se guardó el archivo NOMINA.txt", vbInformation, "Mensaje"
   Else
      MsgBox "No existen registros con ésta fecha", vbCritical, "Mensaje"
      
   End If
   Close #1
End If
frm_personal.MousePointer = 0

End Sub

Private Sub Command3_Click()
Dim Xdesdeff As String
Dim Micadtarr As String
Dim XCuentaa As Long
Dim Xcadapel As String
Xdesdeff = InputBox("Ingrese desde que fecha de ingreso? (Ej.01/10/2008)...:", "Selección de datos")
If Xdesdeff <> "" Then
   frm_personal.MousePointer = 11
   data_per.RecordSource = "select * from tarjbrou where fecing >=#" & Format(Xdesdeff, "yyyy/mm/dd") & "# And rellab <>" & 99 & " order by fecing"
   data_per.Refresh
   If data_per.Recordset.RecordCount > 0 Then
      data_per.Recordset.MoveFirst
      Open App.Path & "\CP_SAPP.txt" For Output As #1
      Do While Not data_per.Recordset.EOF
         If Len(Trim(Str(data_per.Recordset("cedula")))) < 7 Then
            Micadtarr = "0" + Trim(Str(data_per.Recordset("cedula"))) + Trim(Str(data_per.Recordset("codver")))
         Else
            Micadtarr = Trim(Str(data_per.Recordset("cedula"))) + Trim(Str(data_per.Recordset("codver")))
         End If
         Xcadapel = Trim(data_per.Recordset("ape1")) + " "
         If IsNull(data_per.Recordset("ape2")) = True Then
         Else
            If data_per.Recordset("ape2") <> "" Then
               Xcadapel = Xcadapel + Trim(data_per.Recordset("ape2")) + " "
            End If
         End If
         Xcadapel = Xcadapel + Trim(data_per.Recordset("nom1")) + " "
         If IsNull(data_per.Recordset("nom2")) = True Then
         Else
            Xcadapel = Xcadapel + Trim(data_per.Recordset("nom2")) + " "
         End If
         If Len(Xcadapel) > 30 Then
            Micadtarr = Micadtarr + Mid(Xcadapel, 1, 30)
            XCuentaa = 99
         Else
            Micadtarr = Micadtarr + Xcadapel
            XCuentaa = Len(Xcadapel)
         End If
         If XCuentaa = 99 Then
         Else
            For XCuentaa = XCuentaa To 29
'                Xcadapel = Xcadapel + " "
                 Micadtarr = Micadtarr + " "
            Next
'            Micadtarr = Micadtarr + Xcadapel
         End If
         If IsNull(data_per.Recordset("fecing")) = False Then
            Micadtarr = Micadtarr + Mid(data_per.Recordset("fecing"), 1, 2)
            Micadtarr = Micadtarr + Mid(data_per.Recordset("fecing"), 4, 2)
            Micadtarr = Micadtarr + Mid(data_per.Recordset("fecing"), 7, 4)
         Else
            Micadtarr = Micadtarr + "01012010"
         End If
         If IsNull(data_per.Recordset("rellab")) = False Then
            Micadtarr = Micadtarr + Trim(Str(data_per.Recordset("rellab")))
         Else
            Micadtarr = Micadtarr + "3"
         End If
         If IsNull(data_per.Recordset("forpag")) = False Then
            Micadtarr = Micadtarr + Trim(Str(data_per.Recordset("forpag")))
         Else
            Micadtarr = Micadtarr + "1"
         End If
         Print #1, Trim(Micadtarr)
         data_per.Recordset.MoveNext
         XCuentaa = 0
      Loop
      Close #1
      MsgBox "Proceso terminado, se guardó el archivo CP_SAPP.txt", vbInformation, "Mensaje"
   Else
      MsgBox "No existen registros con ésta fecha", vbCritical, "Mensaje"
      
   End If
End If
frm_personal.MousePointer = 0

End Sub

Private Sub Form_Load()
data_per.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_per.Refresh
igualaper

End Sub

Public Function borrar_tar()
txt_ced.Text = ""
txt_cod.Text = ""
txt_nom1.Text = ""
txt_nom2.Text = ""
txt_ape1.Text = ""
txt_ape2.Text = ""
txt_tel.Text = ""
txt_dir.Text = ""
txt_codp.Text = ""
txt_loc.Text = ""
txt_dpto.Text = ""
mfing.Text = "__/__/____"
txt_rlab.Text = ""
txt_fpag.Text = ""

End Function

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfing_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_rlab.SetFocus
End If

End Sub

Private Sub txt_ape1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_ape2.SetFocus
End If

End Sub

Private Sub txt_ape2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_dir.SetFocus
End If

End Sub

Private Sub txt_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_cod.SetFocus
End If

End Sub

Private Sub txt_cod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nom1.SetFocus
End If

End Sub

Private Sub txt_codp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_loc.SetFocus
End If

End Sub

Private Sub txt_dir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_tel.SetFocus
End If

End Sub

Private Sub txt_dpto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfing.SetFocus
End If

End Sub

Private Sub txt_fpag_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_gra.SetFocus
End If

End Sub

Private Sub txt_loc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_dpto.SetFocus
End If

End Sub

Private Sub txt_nom1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nom2.SetFocus
End If

End Sub

Private Sub txt_nom2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_ape1.SetFocus
End If

End Sub

Private Sub txt_rlab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_fpag.SetFocus
End If

End Sub

Private Sub txt_tel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_codp.SetFocus
End If

End Sub

Public Function igualaper()
txt_ced.Text = data_per.Recordset("cedula")
txt_cod.Text = data_per.Recordset("codver")
txt_nom1.Text = data_per.Recordset("nom1")
If IsNull(data_per.Recordset("nom2")) = False Then
   txt_nom2.Text = data_per.Recordset("nom2")
Else
   txt_nom2.Text = ""
End If
txt_ape1.Text = data_per.Recordset("ape1")
If IsNull(data_per.Recordset("ape2")) = False Then
   txt_ape2.Text = data_per.Recordset("ape2")
Else
   txt_ape2.Text = ""
End If
If IsNull(data_per.Recordset("calle")) = False Then
   txt_dir.Text = data_per.Recordset("calle")
Else
   txt_dir.Text = ""
End If
If IsNull(data_per.Recordset("codpos")) = False Then
   txt_codp.Text = data_per.Recordset("codpos")
Else
   txt_codp.Text = 0
End If
If IsNull(data_per.Recordset("telpart")) = False Then
   txt_tel.Text = data_per.Recordset("telpart")
Else
   txt_tel.Text = ""
End If
If IsNull(data_per.Recordset("localid")) = False Then
   txt_loc.Text = data_per.Recordset("localid")
Else
   txt_loc.Text = ""
End If
If IsNull(data_per.Recordset("depto")) = False Then
   txt_dpto.Text = data_per.Recordset("depto")
Else
   txt_dpto.Text = ""
End If
If IsNull(data_per.Recordset("fecing")) = False Then
   mfing.Text = Format(data_per.Recordset("fecing"), "dd/mm/yyyy")
Else
   mfing.Text = "__/__/____"
End If
If IsNull(data_per.Recordset("rellab")) = False Then
   txt_rlab.Text = data_per.Recordset("rellab")
Else
   txt_rlab.Text = 3
End If
If IsNull(data_per.Recordset("forpag")) = False Then
   txt_fpag.Text = data_per.Recordset("forpag")
Else
   txt_fpag.Text = 1
End If

End Function
