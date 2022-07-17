VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_llamadotot 
   BackColor       =   &H00404000&
   Caption         =   "Llamados"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13065
   Icon            =   "frm_llamadotot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   13065
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_mov 
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
      Height          =   375
      Left            =   9360
      TabIndex        =   18
      Top             =   600
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FF80&
      Caption         =   "Desde respaldos"
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
      Left            =   9480
      TabIndex        =   16
      Top             =   4680
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   12120
      Picture         =   "frm_llamadotot.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox txt_apel 
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
      Left            =   2640
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "TELÉFONO"
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
      Left            =   7920
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "NOMBRES"
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
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "FECHAS"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   3180
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_llamadotot.frx":09CC
      Height          =   3735
      Left            =   240
      OleObjectBlob   =   "frm_llamadotot.frx":09E0
      TabIndex        =   3
      Top             =   960
      Width           =   12375
   End
   Begin MSMask.MaskEdBox mfech 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
   Begin MSMask.MaskEdBox mfecd 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nro. de Móvil:"
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
      Left            =   7920
      TabIndex        =   17
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
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
      Left            =   5760
      TabIndex        =   14
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Boleta No."
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
      Left            =   4680
      TabIndex        =   13
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2760
      TabIndex        =   12
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Llamado con costo....."
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
      TabIndex        =   11
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Haga click sobre el registro para ver referencias."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5760
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   240
      TabIndex        =   9
      Top             =   5040
      Width           =   12375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Buscar por:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   7080
      Picture         =   "frm_llamadotot.frx":4DAB
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2415
   End
End
Attribute VB_Name = "frm_llamadotot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub DBGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data1.Recordset.RecordCount > 0 Then
   If IsNull(Data1.Recordset("referen")) = False Then
      Label2.Caption = Data1.Recordset("referen")
   Else
      Label2.Caption = ""
   End If
   If IsNull(Data1.Recordset("realiza")) = False Then
      If Data1.Recordset("realiza") = 1 Then
         If IsNull(Data1.Recordset("mes")) = False Then
            Label5.Caption = Format(Data1.Recordset("mes"), "Standard")
            If IsNull(Data1.Recordset("totend")) = False Then
               Label8.Caption = Data1.Recordset("totend")
            Else
               Label8.Caption = ""
            End If
         Else
            Label5.Caption = ""
         End If
         If IsNull(Data1.Recordset("ano")) = False Then
            Label7.Caption = Data1.Recordset("ano")
         Else
            Label7.Caption = ""
         End If
      Else
         Label5.Caption = ""
         Label7.Caption = ""
         Label8.Caption = ""
      End If
   Else
      Label5.Caption = ""
      Label7.Caption = ""
      Label8.Caption = ""
   End If
End If

End Sub

Private Sub Form_Load()
Dim Xfec, Xfec2 As Date
Xfec = Date - 5
Xfec2 = Date - 1
If Check1.value = 1 Then
   Data1.DatabaseName = App.Path & "\llamado.mdb"
Else
   Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If

If XWeltipoU = "USUARIOS" Then
   Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(Xfec, "yyyy/mm/dd") & "# and fecha <#" & Format(Date, "yyyy/mm/dd") & "# and hor_rea is not null order by fecha,hora"
   Option2.Visible = True
   Option3.Visible = False
Else
   Option2.Visible = True
   Option3.Visible = True
   Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(Xfec, "yyyy/mm/dd") & "# and fecha <#" & Format(Date, "yyyy/mm/dd") & "# order by fecha,hora"
End If
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveLast
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

Private Sub Form_Terminate()
Unload Me

End Sub

Private Sub mfecd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfech.SetFocus
End If

End Sub

Private Sub mfecd_LostFocus()
If mfecd.Text <> "__/__/____" Then
   If Format(mfecd.Text, "dd/mm/yyyy") = Format(Date, "dd/mm/yyyy") Then
      MsgBox "El día no puede ser igual a la fecha actual"
      mfecd.SetFocus
   End If
End If

End Sub

Private Sub mfech_KeyPress(KeyAscii As Integer)

If Format(mfech.Text, "dd/mm/yyyy") = Format(Date, "dd/mm/yyyy") Then
   mfech.Text = Format(CDate(mfech.Text) - 1, "dd/mm/yyyy")
End If

If Check1.value = 1 Then
   Data1.DatabaseName = App.Path & "\llamado.mdb"
Else
   Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If

If KeyAscii = 13 Then
   If Option1.value = True Then
      If XWeltipoU = "USUARIOS" Then
         If mfecd.Text = mfech.Text Then
            t_mov.SetFocus
         Else
            MsgBox "Las fechas deben ser iguales"
         End If
      Else
         Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(mfecd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfech.Text, "yyyy/mm/dd") & "# order by fecha,hora"
         Data1.Refresh
         DBGrid1.SetFocus
      End If
   End If
End If
      
End Sub

Private Sub mhd_KeyPress(KeyAscii As Integer)

End Sub

Private Sub mfech_LostFocus()
If mfech.Text <> "__/__/____" Then
   If Format(mfech.Text, "dd/mm/yyyy") = Format(Date, "dd/mm/yyyy") Then
      MsgBox "El día no puede ser igual a la fecha actual"
      mfech.SetFocus
   End If
End If

End Sub

Private Sub Option1_Click()
txt_apel.Visible = False
mfecd.Visible = True
mfech.Visible = True


End Sub

Private Sub Option2_Click()
mfecd.Visible = False
mfech.Visible = False
txt_apel.Visible = True

End Sub

Private Sub Option3_Click()
mfecd.Visible = False
mfech.Visible = False
txt_apel.Visible = True

End Sub

Private Sub t_mov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   DBGrid1.SetFocus
End If

End Sub

Private Sub t_mov_LostFocus()
frm_llamadotot.MousePointer = 11
If t_mov.Text = "" Then
   Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(mfecd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfech.Text, "yyyy/mm/dd") & "# order by fecha,hora"
   Data1.Refresh
Else
   Data1.RecordSource = "Select * from llamado where fecha >=#" & Format(mfecd.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mfech.Text, "yyyy/mm/dd") & "# and movilpas =" & t_mov.Text & " order by fecha,hora"
   Data1.Refresh
End If

frm_llamadotot.MousePointer = 0


End Sub

Private Sub txt_apel_KeyPress(KeyAscii As Integer)
Dim Xfeccc, Xfecdesde As Date
Xfecdesde = Date - 395
Xfeccc = Date - 1
If Check1.value = 1 Then
   Data1.DatabaseName = App.Path & "\llamado.mdb"
Else
   Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If

If KeyAscii = 13 Then
   If Option2.value = True Then
      Data1.RecordSource = "Select * from llamado where nombre >='" & txt_apel.Text & "' and fecha >=#" & Format(Xfecdesde, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xfeccc, "yyyy/mm/dd") & "# order by nombre,fecha"
      Data1.Refresh
   End If
   If Option3.value = True Then
      Data1.RecordSource = "Select * from llamado where telef >='" & txt_apel.Text & "' and fecha >=#" & Format(Xfecdesde, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xfeccc, "yyyy/mm/dd") & "# order by telef,fecha"
      Data1.Refresh
   End If
   DBGrid1.SetFocus
End If

End Sub
