VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_impcarne 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Informes Carné de Salud"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_impcarne.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   6405
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr1 
      Left            =   2400
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_carne 
      Caption         =   "data_carne"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_hc 
      Caption         =   "data_hc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   615
      Left            =   4080
      TabIndex        =   10
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   600
      TabIndex        =   9
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Informes"
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin VB.TextBox t_codced 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox t_ced 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_impcarne.frx":058A
         Left            =   1560
         List            =   "frm_impcarne.frx":0594
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   3375
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cédula (opcional)"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Opción:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fechas:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   120
      Picture         =   "frm_impcarne.frx":05BA
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   975
   End
End
Attribute VB_Name = "frm_impcarne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Combo1.ListIndex = 0 Then
   data_inf.DatabaseName = App.Path & "\carne.mdb"
   data_inf.RecordSource = "carne"
   data_inf.Refresh
   If data_inf.Recordset.RecordCount > 0 Then
      data_inf.Recordset.MoveFirst
      Do While Not data_inf.Recordset.EOF
         data_inf.Recordset.Delete
         data_inf.Recordset.MoveNext
      Loop
   End If
   If t_ced.Text <> "" Then
      data_hc.RecordSource = "Select * from cabezal_hcdig where tipo_consd ='" & "Carné de Salud" & "' and cednum =" & t_ced.Text
      data_hc.Refresh
      If data_hc.Recordset.RecordCount > 0 Then
         data_inf.Recordset.AddNew
         data_carne.RecordSource = "Select * from agenda where mat =" & data_hc.Recordset("mat")
         data_carne.Refresh
         If data_carne.Recordset.RecordCount > 0 Then
            data_inf.Recordset("valido") = data_carne.Recordset("fec_nac")
            data_inf.Recordset("cedula") = Trim(Str(data_hc.Recordset("cednum"))) & "-" & Trim(Str(data_hc.Recordset("codced")))
            data_inf.Recordset("expedido") = data_hc.Recordset("fecha")
            data_inf.Recordset("medico") = data_carne.Recordset("nommed")
            data_inf.Recordset("cp") = data_carne.Recordset("tipo_cons")
            data_carne.RecordSource = "Select * from cabezal_hc where cb_mat =" & data_hc.Recordset("mat")
            data_carne.Refresh
            If data_carne.Recordset.RecordCount > 0 Then
               data_inf.Recordset("nombre") = data_carne.Recordset("cb_nom1")
               data_inf.Recordset("apellido") = data_carne.Recordset("cb_ape1")
               data_inf.Recordset("fnac") = data_carne.Recordset("cb_fnac")
               If IsNull(data_carne.Recordset("cb_sexo")) = False Then
                  If data_carne.Recordset("cb_sexo") = "F" Then
                     data_inf.Recordset("sexo") = "FEMENINO"
                  Else
                     If data_carne.Recordset("cb_sexo") = "M" Then
                        data_inf.Recordset("sexo") = "MASCULINO"
                     Else
                        data_inf.Recordset("sexo") = "NO REG."
                     End If
                  End If
               Else
                   data_inf.Recordset("sexo") = "NO REG."
               End If
            Else
               MsgBox "No se encuentra cabezal de HC"
            End If
            data_inf.Recordset.Update
            data_inf.Refresh
            cr1.ReportFileName = App.Path & "\carne.rpt"
            cr1.Action = 1
            
         Else
            MsgBox "No se encuentra HC"
         End If
      Else
         MsgBox "No se encuentra HC"
      End If
   Else
      MsgBox "Debe seleccionar una cédula para imprimir el carné"
      
   End If
End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_inf.DatabaseName = App.Path & "\informes.mdb"
'data_inf.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_hc.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_carne.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

Private Sub mfh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
   
End If

End Sub
