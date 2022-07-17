VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_busperso 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos del personal"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   Icon            =   "frm_busperso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9030
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_bus 
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
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8160
      Picture         =   "frm_busperso.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   4320
      Width           =   615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_busperso.frx":09CC
      Height          =   3855
      Left            =   240
      OleObjectBlob   =   "frm_busperso.frx":09E0
      TabIndex        =   1
      Top             =   480
      Width           =   8535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Con doble click edita los datos en el formulario."
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
      TabIndex        =   3
      Top             =   4320
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Apellido a buscar:"
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
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   6120
      Picture         =   "frm_busperso.frx":39BF
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "frm_busperso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()

frm_prestamo.tced.Text = Data1.Recordset("cedula")
frm_prestamo.tcod.Text = Data1.Recordset("codver")
frm_prestamo.tsex.Text = Data1.Recordset("sexo")
frm_prestamo.tn1.Text = Data1.Recordset("nom1")
If IsNull(Data1.Recordset("nom2")) = False Then
   frm_prestamo.tn2.Text = Data1.Recordset("nom2")
Else
   frm_prestamo.tn2.Text = ""
End If
frm_prestamo.tape1.Text = Data1.Recordset("ape1")
If IsNull(Data1.Recordset("ape2")) = False Then
   frm_prestamo.tape2.Text = Data1.Recordset("ape2")
Else
   frm_prestamo.tape2.Text = ""
End If
frm_prestamo.tdir.Text = Data1.Recordset("calle")
If IsNull(Data1.Recordset("nropuerta")) = False Then
   frm_prestamo.tnp.Text = Data1.Recordset("nropuerta")
Else
   frm_prestamo.tnp.Text = ""
End If
frm_prestamo.tcoddep.Text = Data1.Recordset("coddep")
frm_prestamo.tloc.Text = Data1.Recordset("localid")
frm_prestamo.tcodpos.Text = Data1.Recordset("codpos")
If IsNull(Data1.Recordset("telef")) = False Then
   frm_prestamo.txt_tel.Text = Data1.Recordset("telef")
Else
   frm_prestamo.txt_tel.Text = ""
End If
If IsNull(Data1.Recordset("estciv")) = False Then
   frm_prestamo.txt_estciv.Text = Data1.Recordset("estciv")
Else
   frm_prestamo.txt_estciv.Text = 1
End If
If IsNull(Data1.Recordset("cedc")) = False Then
   frm_prestamo.txt_cedc.Text = Data1.Recordset("cedc")
Else
   frm_prestamo.txt_cedc.Text = 0
End If
If IsNull(Data1.Recordset("codcedc")) = False Then
   frm_prestamo.txt_codcedc.Text = Data1.Recordset("codcedc")
Else
   frm_prestamo.txt_codcedc.Text = 0
End If
If IsNull(Data1.Recordset("nomc")) = False Then
   frm_prestamo.txt_nomc.Text = Data1.Recordset("nomc")
Else
   frm_prestamo.txt_nomc.Text = ""
End If
frm_prestamo.mnac.Text = Format(Data1.Recordset("fecnac"), "dd/mm/yyyy")
frm_prestamo.ming.Text = Format(Data1.Recordset("fecing"), "dd/mm/yyyy")
frm_prestamo.tden.Text = Data1.Recordset("desccar")
frm_prestamo.tcarg.Text = Data1.Recordset("caraccar")
frm_prestamo.tmj.Text = Data1.Recordset("mj")
frm_prestamo.thd.Text = Data1.Recordset("hd")
frm_prestamo.tcanhd.Text = Data1.Recordset("canths")
frm_prestamo.tcuo.Text = Format(Data1.Recordset("cuosug"), "Standard")
frm_prestamo.tret.Text = Format(Data1.Recordset("retjud"), "Standard")
frm_prestamo.tgtia.Text = Format(Data1.Recordset("alqui"), "Standard")
frm_prestamo.tretleg.Text = Format(Data1.Recordset("retleg"), "Standard")
frm_prestamo.timps.Text = Format(Data1.Recordset("impsue"), "Standard")
frm_prestamo.ttoth.Text = Format(Data1.Recordset("tothab"), "Standard")
Unload Me

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   DBGrid1_DblClick
End If

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\prestamo.mdb"
Data1.RecordSource = "prestamo"
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

Private Sub txt_bus_KeyPress(KeyAscii As Integer)
Data1.RecordSource = "Select * from prestamo where ape1 >='" & txt_bus.Text & "' order by ape1"
Data1.Refresh
If KeyAscii = 13 Then
   DBGrid1.SetFocus
End If

End Sub
