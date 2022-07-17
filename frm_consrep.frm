VERSION 5.00
Begin VB.Form frm_consrep 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas repetidas"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6480
   Icon            =   "frm_consrep.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_lla 
      Caption         =   "data_lla"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de informe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5895
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
         Caption         =   "Informe sin detalle"
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
         TabIndex        =   8
         Top             =   2040
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808080&
         Caption         =   "Desde facturación"
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
         Left            =   3000
         TabIndex        =   7
         Top             =   1200
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808080&
         Caption         =   "Desde llamado"
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
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txt_h 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txt_d 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "FECHAS:"
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
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frm_consrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text2_Change()

End Sub

Private Sub Command1_Click()
Dim Xquematr, Xquecanti, Xcantrgs, Xquecantgral As Long
frm_consrep.MousePointer = 11
Command1.Enabled = False
Command2.Enabled = False
If txt_d.Text <> "" Then
   If Option1.Value = True Then
      data_lla.RecordSource = "select * from llamado where fecha >=#" & Format(CDate(txt_d.Text), "yyyy/mm/dd") & "# and fecha <=#" & Format(CDate(txt_h.Text), "yyyy/mm/dd") & "# order by matric"
      data_lla.Refresh
   Else
      data_lla.RecordSource = "select * from linmmdd where fecha >=#" & Format(CDate(txt_d.Text), "yyyy/mm/dd") & "# and fecha <=#" & Format(CDate(txt_h.Text), "yyyy/mm/dd") & "# And nro_flia =" & 1 & " order by cod_cli"
      data_lla.Refresh
   End If
   If data_inf.Recordset.RecordCount > 0 Then
      data_inf.Recordset.MoveFirst
      Do While Not data_inf.Recordset.EOF
         data_inf.Recordset.Delete
         data_inf.Recordset.MoveNext
      Loop
   End If
   If data_lla.Recordset.RecordCount > 0 Then
      data_lla.Recordset.MoveLast
      Xcantrgs = data_lla.Recordset.RecordCount
      Xquecanti = 0
      data_lla.Recordset.MoveFirst
      If Option1.Value = True Then
         Xquematr = data_lla.Recordset("matric")
      Else
         Xquematr = data_lla.Recordset("cod_cli")
      End If
      Do While Not data_lla.Recordset.EOF
         If Option1.Value = True Then
            If data_lla.Recordset("matric") = 0 Then
               data_lla.Recordset.MoveNext
               Xquecanti = 0
            Else
               If data_lla.Recordset("matric") = 99999999 Then
                  data_lla.Recordset.MoveNext
                  Xquecanti = 0
               Else
                  If data_lla.Recordset("matric") = 999999999 Then
                     data_lla.Recordset.MoveNext
                     Xquecanti = 0
                  Else
                    If Xquematr = data_lla.Recordset("matric") Then
                       Xquecanti = Xquecanti + 1
                       data_lla.Recordset.MoveNext
                    Else
                       If Xquecanti > 1 Then
                          Xquecantgral = Xquecantgral + Xquecanti
                          data_lla.Recordset.MovePrevious
                          data_inf.Recordset.AddNew
                          data_inf.Recordset("cl_codigo") = data_lla.Recordset("matric")
                          data_inf.Recordset("cl_apellid") = data_lla.Recordset("nombre")
                          data_inf.Recordset("cl_codconv") = data_lla.Recordset("categ")
                          data_inf.Recordset("saldo_cc") = Xquecanti - 1
                          data_inf.Recordset("saldo_cc2") = Xcantrgs
                          data_inf.Recordset("cl_base") = data_lla.Recordset("movilpas")
                          data_inf.Recordset.Update
                          data_lla.Recordset.MoveNext
                          Xquematr = data_lla.Recordset("matric")
                          Xquecanti = 0
                       Else
                          Xquematr = data_lla.Recordset("matric")
                          Xquecanti = 0
                       End If
                    End If
                  End If
               End If
            End If
         End If
         If Option2.Value = True Then
            If Xquematr = data_lla.Recordset("cod_cli") Then
               Xquecanti = Xquecanti + 1
               data_lla.Recordset.MoveNext
            Else
               If Xquecanti > 1 Then
                  Xquecantgral = Xquecantgral + Xquecanti
                  data_lla.Recordset.MovePrevious
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_codigo") = data_lla.Recordset("cod_cli")
                  data_inf.Recordset("cl_apellid") = data_lla.Recordset("nom_cli")
                  data_inf.Recordset("cl_codconv") = data_lla.Recordset("convenio")
                  data_inf.Recordset("cl_base") = data_lla.Recordset("base")
                  data_inf.Recordset("saldo_cc") = Xquecanti - 1
                  data_inf.Recordset("saldo_cc2") = Xcantrgs
                  data_inf.Recordset.Update
                  data_lla.Recordset.MoveNext
                  Xquematr = data_lla.Recordset("cod_cli")
                  Xquecanti = 0
               Else
                  Xquematr = data_lla.Recordset("cod_cli")
                  Xquecanti = 0
               End If
            End If
         End If
      Loop
   End If
End If
frm_consrep.MousePointer = 0
Command1.Enabled = False
Command2.Enabled = False
MsgBox "Proceso terminado"


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_lla.DatabaseName = App.Path & "\sapp.mdb"
data_lla.RecordSource = "llamado"
data_lla.Refresh
data_inf.DatabaseName = App.Path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh


End Sub
