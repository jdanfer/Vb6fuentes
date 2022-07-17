VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_cancella2 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelar pasaje de llamado"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6060
   Icon            =   "frm_cancella2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6060
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_nro 
      Height          =   405
      Left            =   4200
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox t_mov 
      Alignment       =   1  'Right Justify
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
      Left            =   1920
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4320
      Picture         =   "frm_cancella2.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   480
      Picture         =   "frm_cancella2.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frm_cancella2.frx":109E
      Left            =   1920
      List            =   "frm_cancella2.frx":10A8
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1320
      Width           =   3135
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H008080FF&
      Caption         =   "Cancelación solo traslado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000080FF&
      Caption         =   "Cancelación de llamado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin MSMask.MaskEdBox mh 
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MOVIL:"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Motivo:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha y Hora:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   3000
      Picture         =   "frm_cancella2.frx":10C6
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "frm_cancella2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_mov.SetFocus
End If

End Sub

Private Sub Command1_Click()
Dim Xdesea As String
Dim Xcolo As String
Xcolo = ""

If t_nro.Text <> "" Then
   If mf.Text = "__/__/____" And mh.Text = "__:__" Then
      Xdesea = MsgBox("Desea CANCELAR el registro del llamado NO ACEPTADO?", vbInformation + vbYesNo)
      If Xdesea = vbYes Then
         Data1.RecordSource = "Select * from resplla where nro =" & t_nro.Text
         Data1.Refresh
         If Data1.Recordset.RecordCount > 0 Then
            If IsNull(Data1.Recordset("fecpas")) = False Then
               Data1.Recordset.Edit
               Data1.Recordset("fecpas") = Null
               Data1.Recordset.Update
            End If
            If IsNull(Data1.Recordset("horpas")) = False Then
               Data1.Recordset.Edit
               Data1.Recordset("horpas") = Null
               Data1.Recordset.Update
            End If
         End If
      End If
   Else
        Data1.RecordSource = "Select * from resplla where nro =" & t_nro.Text
        Data1.Refresh
        If Data1.Recordset.RecordCount > 0 Then
           If IsNull(Data1.Recordset("fecpas")) = False Then
              If Format(Data1.Recordset("fecpas"), "dd/mm/yyyy") <> Format(mf.Text, "dd/mm/yyyy") Then
                 Data1.Recordset.Edit
                 Data1.Recordset("fecpas") = Format(mf.Text, "dd/mm/yyyy")
                 Data1.Recordset.Update
              End If
           Else
              Data1.Recordset.Edit
              Data1.Recordset("fecpas") = Format(mf.Text, "dd/mm/yyyy")
              Data1.Recordset.Update
           End If
           If IsNull(Data1.Recordset("horpas")) = False Then
              If Format(Data1.Recordset("horpas"), "HH:mm") <> Format(mh.Text, "HH:mm") Then
                 Data1.Recordset.Edit
                 Data1.Recordset("horpas") = Format(mh.Text, "HH:mm")
                 Data1.Recordset.Update
              End If
           Else
              Data1.Recordset.Edit
              Data1.Recordset("horpas") = Format(mh.Text, "HH:mm")
              Data1.Recordset.Update
           End If
           If IsNull(Data1.Recordset("realiza")) = False Then
              If Data1.Recordset("realiza") <> Combo1.ListIndex Then
                 Data1.Recordset.Edit
                 Data1.Recordset("realiza") = Combo1.ListIndex
                 Data1.Recordset.Update
              End If
           Else
              Data1.Recordset.Edit
              Data1.Recordset("realiza") = Combo1.ListIndex
              Data1.Recordset.Update
           End If
           If frm_largador.cbocolor.Text = "ROJO" Then
              Xcolo = "R"
           Else
              If frm_largador.cbocolor.Text = "CELESTE" Then
                 Xcolo = "C"
              Else
                 If frm_largador.cbocolor.Text = "AMARILLO" Then
                    Xcolo = "A"
                 Else
                    If frm_largador.cbocolor.Text = "AZUL" Then
                       Xcolo = "Z"
                    Else
                       Xcolo = "V"
                    End If
                 End If
              End If
           End If
           If frm_largador.cbocolor.ListIndex >= 0 Then
              If IsNull(Data1.Recordset("colormot")) = False Then
                 If Data1.Recordset("colormot") <> Trim(Xcolo) Then
                    Data1.Recordset.Edit
                    Data1.Recordset("colormot") = Trim(Xcolo)
                    Data1.Recordset.Update
                 End If
              Else
                 Data1.Recordset.Edit
                 Data1.Recordset("colormot") = Trim(Xcolo)
                 Data1.Recordset.Update
              End If
           End If
           If IsNull(Data1.Recordset("codmot")) = False Then
              If Option1.Value = True Then
                 If Data1.Recordset("codmot") <> "L" Then
                    Data1.Recordset.Edit
                    Data1.Recordset("codmot") = "L"
                    Data1.Recordset.Update
                 End If
              End If
              If Option2.Value = True Then
                 If Data1.Recordset("codmot") <> "T" Then
                    Data1.Recordset.Edit
                    Data1.Recordset("codmot") = "T"
                    Data1.Recordset.Update
                 End If
              End If
           Else
              If Option1.Value = True Then
                 Data1.Recordset.Edit
                 Data1.Recordset("codmot") = "L"
                 Data1.Recordset.Update
              End If
              If Option2.Value = True Then
                 Data1.Recordset.Edit
                 Data1.Recordset("codmot") = "T"
                 Data1.Recordset.Update
              End If
           End If
           If t_mov.Text <> "" Then
              If IsNull(Data1.Recordset("trasla")) = False Then
                 If Data1.Recordset("trasla") <> t_mov.Text Then
                    Data1.Recordset.Edit
                    Data1.Recordset("trasla") = t_mov.Text
                    Data1.Recordset.Update
                 End If
              Else
                 Data1.Recordset.Edit
                 Data1.Recordset("trasla") = t_mov.Text
                 Data1.Recordset.Update
              End If
           End If
        Else
           Data1.Recordset.AddNew
           If mf.Text <> "__/__/____" Then
              Data1.Recordset("fecpas") = Format(mf.Text, "dd/mm/yyyy")
           End If
           If mh.Text <> "__:__" Then
              Data1.Recordset("horpas") = Format(mh.Text, "HH:mm")
           End If
           Data1.Recordset("realiza") = Combo1.ListIndex
           If Option1.Value = True Then
              Data1.Recordset("codmot") = "L"
           End If
           If Option2.Value = True Then
              Data1.Recordset("codmot") = "T"
           End If
           If t_mov.Text <> "" Then
               Data1.Recordset("trasla") = t_mov.Text
           End If
           If frm_largador.cbocolor.Text = "ROJO" Then
              Xcolo = "R"
           Else
              If frm_largador.cbocolor.Text = "CELESTE" Then
                 Xcolo = "C"
              Else
                 If frm_largador.cbocolor.Text = "AMARILLO" Then
                    Xcolo = "A"
                 Else
                    If frm_largador.cbocolor.Text = "AZUL" Then
                       Xcolo = "Z"
                    Else
                       Xcolo = "V"
                    End If
                 End If
              End If
           End If
           If Xcolo <> "" Then
              Data1.Recordset("colormot") = Trim(Xcolo)
           End If
           Data1.Recordset.Update
           Unload Me
        End If
   End If
Else
   MsgBox "No hay llamado seleccionado"
   
End If




End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()

mf.Text = Format(Date, "dd/mm/yyyy")
mh.Text = Format(Time, "HH:mm")

Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

t_nro.Text = frm_largador.txt_nro.Text

Data1.RecordSource = "Select * from resplla where nro =" & frm_largador.txt_nro.Text
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   If IsNull(Data1.Recordset("fecpas")) = False And IsNull(Data1.Recordset("horpas")) = False Then
      mf.Text = Format(Data1.Recordset("fecpas"), "dd/mm/yyyy")
      mh.Text = Format(Data1.Recordset("horpas"), "HH:mm")
      If IsNull(Data1.Recordset("realiza")) = False Then
         Combo1.ListIndex = Data1.Recordset("realiza")
      Else
         Combo1.ListIndex = -1
      End If
      If IsNull(Data1.Recordset("codmot")) = False Then
         If Data1.Recordset("codmot") = "L" Then
            Option1.Value = True
         Else
            If Data1.Recordset("codmot") = "T" Then
               Option2.Value = True
            End If
         End If
      End If
      If IsNull(Data1.Recordset("trasla")) = False Then
         t_mov.Text = Data1.Recordset("trasla")
      Else
         If frm_largador.txt_movil.Text <> "" Then
            t_mov.Text = frm_largador.txt_movil.Text
         End If
      End If
   Else
   
   End If
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

Private Sub t_mov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub
