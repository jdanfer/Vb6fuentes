VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_cmt 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Registro de datos CMT"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_cmt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   1920
      Picture         =   "frm_cmt.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Ver agenda"
      Top             =   7320
      Width           =   495
   End
   Begin VB.Data data_agenda 
      Caption         =   "data_agenda"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data data_medsapp 
      Caption         =   "data_medsapp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Elegir otro paciente y liberar este CMT"
      Height          =   615
      Left            =   6360
      Picture         =   "frm_cmt.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Deja el registro pendiente en CMT para que pueda tomarlo otro médico"
      Top             =   7920
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   1080
      Picture         =   "frm_cmt.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Imprimir pantalla"
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   9840
      Picture         =   "frm_cmt.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Salir y NO GRABAR los datos"
      Top             =   7320
      Width           =   495
   End
   Begin VB.Data data_lla2 
      Caption         =   "data_lla2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_llam 
      Caption         =   "data_llam"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      Picture         =   "frm_cmt.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Grabar los datos ingresados y CERRAR ventana"
      Top             =   7320
      Width           =   495
   End
   Begin MSMask.MaskEdBox mhcom 
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   16711680
      ForeColor       =   16777215
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "HH:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos para el cierre de CMT"
      Height          =   6855
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   10095
      Begin VB.Data data_conshc 
         Caption         =   "data_conshc"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5400
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data data_hce 
         Caption         =   "data_hce"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4920
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Data data_par 
         Caption         =   "data_par"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4440
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C00000&
         Caption         =   "Derivado a BASE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   1800
         Width           =   3495
      End
      Begin VB.CheckBox chutiliza 
         BackColor       =   &H00FF0000&
         Caption         =   "Confirmación de utilización"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   31
         Top             =   5280
         Width           =   3255
      End
      Begin VB.CheckBox chcompre 
         BackColor       =   &H00FF0000&
         Caption         =   "Comprensión de la medicación"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   30
         Top             =   4920
         Width           =   3255
      End
      Begin VB.CheckBox chiat 
         BackColor       =   &H00FF0000&
         Caption         =   "Confirmación de no iatrogenia"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   29
         Top             =   4560
         Width           =   3255
      End
      Begin VB.CheckBox chaler 
         BackColor       =   &H00FF0000&
         Caption         =   "Confirmación de no alergias"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   28
         Top             =   4200
         Width           =   3255
      End
      Begin VB.CheckBox chansie 
         BackColor       =   &H00FF0000&
         Caption         =   "Ansiedad de los cuidadores"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   27
         Top             =   3360
         Width           =   3135
      End
      Begin VB.CheckBox chempeo 
         BackColor       =   &H00FF0000&
         Caption         =   "Empeoramiento del estado general del paciente."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6600
         TabIndex        =   26
         Top             =   2760
         Width           =   3255
      End
      Begin VB.CheckBox chclicamb 
         BackColor       =   &H00FF0000&
         Caption         =   "Clínica cambiante"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   25
         Top             =   2400
         Width           =   3255
      End
      Begin VB.CheckBox chclisim 
         BackColor       =   &H00FF0000&
         Caption         =   "Clínica similar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   24
         Top             =   2040
         Width           =   3255
      End
      Begin MSMask.MaskEdBox mhfinno 
         Height          =   375
         Left            =   1800
         TabIndex        =   20
         Top             =   3720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mhfin 
         Height          =   375
         Left            =   6600
         TabIndex        =   18
         Top             =   6360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox t_obs2 
         Height          =   735
         Left            =   6600
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   5640
         Width           =   3255
      End
      Begin VB.TextBox t_obsmedic 
         Height          =   495
         Left            =   6600
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   3720
         Width           =   3255
      End
      Begin VB.CheckBox chmedic 
         BackColor       =   &H00C00000&
         Caption         =   "Prescripción de medicamentos"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4440
         TabIndex        =   13
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox t_obsrella 
         Height          =   495
         Left            =   6600
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1560
         Width           =   3255
      End
      Begin VB.CheckBox chrella 
         BackColor       =   &H00C00000&
         Caption         =   "Circunstancias de re-llamada"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4440
         TabIndex        =   9
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox t_obsenf 
         Height          =   495
         Left            =   6600
         MaxLength       =   70
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   3255
      End
      Begin VB.CheckBox chenf 
         BackColor       =   &H00C00000&
         Caption         =   "Curso probable de la enfermedad"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4440
         TabIndex        =   7
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox t_obs1 
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   2760
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_cmt.frx":213C
         Left            =   1920
         List            =   "frm_cmt.frx":2149
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "No envío de recurso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   480
         Width           =   3495
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "Envío de recurso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         Caption         =   "Hora derivado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Hora de Cierre:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   17
         Top             =   6480
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "EN SUMA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Observación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Reclasificación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   4200
         X2              =   4200
         Y1              =   120
         Y2              =   6840
      End
   End
   Begin VB.Label labcedpac 
      Height          =   255
      Left            =   7560
      TabIndex        =   35
      Top             =   8160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   33
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "HORA DE COMIENZO DE CMT:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   120
      Picture         =   "frm_cmt.frx":2164
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   1335
   End
End
Attribute VB_Name = "frm_cmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_graba_Click()
Dim Xlacreclas, XEltextoobs As String
Dim Xseguro, Xdeseacerrar As String
Dim Xhcesi As String
Dim Xcrear As Integer
Dim XnroMedSapp As Integer
Dim XnomMedSapp As String

Xcrear = 0

Xseguro = ""
Xdeseacerrar = ""
'10113480
On Error GoTo Verqueescmt

If frm_largador.txt_ced.Text <> "" Then
   If frm_largador.txt_ced.Text > 0 Then
      If frm_largador.txt_mat.Text <> "" Then
         If frm_largador.txt_mat.Text > 0 Then
            Xcrear = 8
         Else
            Xcrear = 0
         End If
      Else
         Xcrear = 0
      End If
   Else
      Xcrear = 0
   End If
Else
   Xcrear = 0
End If
If Option1.Value = True Or Option3.Value = True Then ''1 no resuelto, 3 deriva a la base
   If t_obs2.Text = "" And chenf.Value = 0 And chrella.Value = 0 And chmedic.Value = 0 And mhfin.Text = "__:__" Then
      If Option1.Value = True Then
         If IsNull(data_llam.Recordset("mm")) = False Then
            If data_llam.Recordset("mm") <> 1 Then
               data_llam.Recordset.Edit
               data_llam.Recordset("mm") = 1
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("mm") = 1
            data_llam.Recordset.Update
         End If
      End If
      If Option3.Value = True Then
         If IsNull(data_llam.Recordset("mm")) = False Then
            If data_llam.Recordset("mm") <> 3 Then
               data_llam.Recordset.Edit
               data_llam.Recordset("mm") = 3
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("mm") = 3
            data_llam.Recordset.Update
         End If
      End If
      If Combo1.ListIndex >= 0 Then
         If Combo1.ListIndex = 0 Then
            Xlacreclas = "R"
         Else
            If Combo1.ListIndex = 1 Then
               Xlacreclas = "A"
            Else
               If Combo1.ListIndex = 2 Then
                  Xlacreclas = "V"
               Else
                  Xlacreclas = ""
               End If
            End If
         End If
         If IsNull(data_llam.Recordset("totend")) = False Then
            If Trim(data_llam.Recordset("totend")) <> Trim(Xlacreclas) Then
               If Xlacreclas = "" Then
                  data_llam.Recordset.Edit
                  data_llam.Recordset("totend") = Null
                  data_llam.Recordset.Update
               Else
                  data_llam.Recordset.Edit
                  data_llam.Recordset("totend") = Xlacreclas
                  data_llam.Recordset.Update
               End If
            End If
         Else
            If Xlacreclas = "" Then
            Else
               data_llam.Recordset.Edit
               data_llam.Recordset("totend") = Xlacreclas
               data_llam.Recordset.Update
            End If
         End If
      Else
         If IsNull(data_llam.Recordset("totend")) = False Then
            data_llam.Recordset.Edit
            data_llam.Recordset("totend") = Null
            data_llam.Recordset.Update
         End If
      End If
      If t_obs1.Text <> "" Then
         If IsNull(data_llam.Recordset("obsmot")) = False Then
            If data_llam.Recordset("obsmot") <> t_obs1.Text Then
               data_llam.Recordset.Edit
               data_llam.Recordset("obsmot") = t_obs1.Text
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("obsmot") = t_obs1.Text
            data_llam.Recordset.Update
         End If
      Else
         If IsNull(data_llam.Recordset("obsmot")) = False Then
            data_llam.Recordset.Edit
            data_llam.Recordset("obsmot") = Null
            data_llam.Recordset.Update
         End If
      End If
      If mhfinno.Text <> "__:__" Then
         If IsNull(data_llam.Recordset("hsald")) = False Then
            If data_llam.Recordset("hsald") <> mhfinno.Text Then
               data_llam.Recordset.Edit
               data_llam.Recordset("hsald") = mhfinno.Text
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("hsald") = mhfinno.Text
            data_llam.Recordset.Update
         End If
      Else
         If IsNull(data_llam.Recordset("hsald")) = False Then
            data_llam.Recordset.Edit
            data_llam.Recordset("hsald") = Null
            data_llam.Recordset.Update
         End If
      End If
      If mhcom.Text <> "__:__" Then
         If IsNull(data_llam.Recordset("horsali")) = False Then
            If data_llam.Recordset("horsali") <> mhcom.Text Then
               data_llam.Recordset.Edit
               data_llam.Recordset("horsali") = mhcom.Text
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("horsali") = mhcom.Text
            data_llam.Recordset.Update
         End If
      Else
         If IsNull(data_llam.Recordset("horsali")) = False Then
            data_llam.Recordset.Edit
            data_llam.Recordset("horsali") = Null
            data_llam.Recordset.Update
         End If
      End If
      If IsNull(data_llam.Recordset("timdes")) = False Then
         If data_llam.Recordset("timdes") <> WElusuario Then
            data_llam.Recordset.Edit
            data_llam.Recordset("timdes") = Mid(WElusuario, 1, 15)
            data_llam.Recordset.Update
         End If
      Else
         data_llam.Recordset.Edit
         data_llam.Recordset("timdes") = Mid(WElusuario, 1, 15)
         data_llam.Recordset.Update
      End If
      If Option1.Value = True Then
         data_lla2.RecordSource = "Select * from llamado where nrolla =" & frm_largador.txt_nro.Text
         data_lla2.Refresh
         If data_lla2.Recordset.RecordCount > 0 Then
            If data_lla2.Recordset("pend") = 4 Then
               data_lla2.Recordset.Edit
               data_lla2.Recordset("pend") = 0
               data_lla2.Recordset("obsmot") = data_lla2.Recordset("obsmot") & " Clasifica como:" & Combo1.Text & " Mot:" & Mid(t_obs1.Text, 1, 50) & " Hora anterior:" & data_lla2.Recordset("hora")
               data_lla2.Recordset("hora_anterior") = data_lla2.Recordset("hora")
               data_lla2.Recordset("hora") = Format(Time, "HH:mm")
               data_lla2.Recordset("activo") = Format(Time, "HH:mm:ss")
               data_lla2.Recordset("cmt_enproceso") = 2
               data_lla2.Recordset.Update
            End If
         End If
         MsgBox "El llamado pasó a PENDIENTES! COMUNIQUE AL LARGADOR!", vbInformation
         frm_largador.Frame2.Visible = True
         frm_largador.Command5.Visible = False
         Unload Me
      Else
         Xdeseacerrar = MsgBox("Desea CERRAR TODO EL LLAMADO COMO REALIZADO?", vbInformation + vbYesNo)
         If Xdeseacerrar = vbYes Then
            data_lla2.RecordSource = "Select * from llamado where nrolla =" & frm_largador.txt_nro.Text
            data_lla2.Refresh
            If data_lla2.Recordset.RecordCount > 0 Then
               If data_lla2.Recordset("pend") = 4 Then
                  data_lla2.Recordset.Edit
                  data_lla2.Recordset("pend") = 2
                  data_lla2.Recordset("movilpas") = 2015
                  data_lla2.Recordset("fecpas") = Format(Date, "dd/mm/yyyy")
                  data_lla2.Recordset("horpas") = Format(Time, "HH:mm")
                  data_lla2.Recordset("fecsali") = Format(Date, "dd/mm/yyyy")
                  data_lla2.Recordset("horsali") = Format(Time, "HH:mm")
                  data_lla2.Recordset("fec_llega") = Format(Date, "dd/mm/yyyy")
                  data_lla2.Recordset("hor_llega") = Format(mhfinno.Text, "HH:mm")
                  data_lla2.Recordset("fec_rea") = Format(Date, "dd/mm/yyyy")
                  data_lla2.Recordset("hor_rea") = Format(mhfinno.Text, "HH:mm")
                  data_lla2.Recordset("diag") = "CMT DERIVADO A BASE"
                  data_lla2.Recordset("colormot") = data_lla2.Recordset("codmot")
                  data_lla2.Recordset("cmt_enproceso") = 2
                  If IsNull(data_lla2.Recordset("mes")) = False Then
                     If data_lla2.Recordset("mes") > 30 Then
                        data_lla2.Recordset("mes") = 0
                     End If
                  End If
                   'data_lla.Recordset("codmed") = txt_codmed.Text
                   'data_lla.Recordset("nommed") = dbcbomed.Text
                  data_lla2.Recordset.Update
               Else
                  If data_lla2.Recordset("pend") <> 2 Then
                    data_lla2.Recordset.Edit
                    data_lla2.Recordset("pend") = 2
                    data_lla2.Recordset("movilpas") = 2015
                    data_lla2.Recordset("fecpas") = Format(Date, "dd/mm/yyyy")
                    data_lla2.Recordset("horpas") = Format(Time, "HH:mm")
                    data_lla2.Recordset("fecsali") = Format(Date, "dd/mm/yyyy")
                    data_lla2.Recordset("horsali") = Format(Time, "HH:mm")
                    data_lla2.Recordset("fec_llega") = Format(Date, "dd/mm/yyyy")
                    data_lla2.Recordset("hor_llega") = Format(mhfinno.Text, "HH:mm")
                    data_lla2.Recordset("fec_rea") = Format(Date, "dd/mm/yyyy")
                    data_lla2.Recordset("hor_rea") = Format(mhfinno.Text, "HH:mm")
                    data_lla2.Recordset("diag") = "CMT DERIVADO A BASE"
                    data_lla2.Recordset("colormot") = data_lla2.Recordset("codmot")
                    data_lla2.Recordset("cmt_enproceso") = 2
                    If IsNull(data_lla2.Recordset("mes")) = False Then
                       If data_lla2.Recordset("mes") > 30 Then
                          data_lla2.Recordset("mes") = 0
                       End If
                    End If
                     'data_lla.Recordset("codmed") = txt_codmed.Text
                     'data_lla.Recordset("nommed") = dbcbomed.Text
                    data_lla2.Recordset.Update
                  End If
               End If
            End If
            frm_largador.Frame2.Visible = True
            frm_largador.Command5.Visible = False
            Unload Me
         Else
            If data_lla2.Recordset("pend") = 4 Then
               data_lla2.Recordset.Edit
               data_lla2.Recordset("pend") = 0
               data_lla2.Recordset("cmt_enproceso") = 2
               data_lla2.Recordset.Update
            End If
            MsgBox "El llamado pasó a PENDIENTES! COMUNIQUE AL LARGADOR!", vbInformation
            frm_largador.Frame2.Visible = True
            frm_largador.Command5.Visible = False
            Unload Me
         End If
      End If
   Else
      MsgBox "Verifique la opción que tiene marcada y los datos igresados"
   End If
Else
   If Option2.Value = True Then
      XcedDoc = ""
      If t_obs1.Text = "" And Combo1.ListIndex < 0 And mhfinno.Text = "__:__" Then
         Xhcesi = MsgBox("Desea crear Historia Clínica para este paciente con estos datos?", vbInformation + vbYesNo)
         XcedDoc = InputBox("Ingrese su número de cédula completo (Ejemplo: para CI:1234567-8, ingresar: 12345678)")
         If Option2.Value = True Then
            If IsNull(data_llam.Recordset("mm")) = False Then
               If data_llam.Recordset("mm") <> 2 Then
                  data_llam.Recordset.Edit
                  data_llam.Recordset("mm") = 2
                  data_llam.Recordset.Update
               End If
            Else
               data_llam.Recordset.Edit
               data_llam.Recordset("mm") = 2
               data_llam.Recordset.Update
            End If
         End If
         If IsNull(data_llam.Recordset("codmed")) = False Then
            If data_llam.Recordset("codmed") <> chclisim.Value Then
               data_llam.Recordset.Edit
               data_llam.Recordset("codmed") = chclisim.Value
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("codmed") = chclisim.Value
            data_llam.Recordset.Update
         End If
         If IsNull(data_llam.Recordset("thh")) = False Then
            If data_llam.Recordset("thh") <> chclicamb.Value Then
               data_llam.Recordset.Edit
               data_llam.Recordset("thh") = chclicamb.Value
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("thh") = chclicamb.Value
            data_llam.Recordset.Update
         End If
         If IsNull(data_llam.Recordset("ano")) = False Then
            If data_llam.Recordset("ano") <> chempeo.Value Then
               data_llam.Recordset.Edit
               data_llam.Recordset("ano") = chempeo.Value
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("ano") = chempeo.Value
            data_llam.Recordset.Update
         End If
         If IsNull(data_llam.Recordset("movtras")) = False Then
            If data_llam.Recordset("movtras") <> chansie.Value Then
               data_llam.Recordset.Edit
               data_llam.Recordset("movtras") = chansie.Value
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("movtras") = chansie.Value
            data_llam.Recordset.Update
         End If
         If IsNull(data_llam.Recordset("trasla")) = False Then
            If data_llam.Recordset("trasla") <> chaler.Value Then
               data_llam.Recordset.Edit
               data_llam.Recordset("trasla") = chaler.Value
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("trasla") = chaler.Value
            data_llam.Recordset.Update
         End If
         If IsNull(data_llam.Recordset("realiza")) = False Then
            If data_llam.Recordset("realiza") <> chiat.Value Then
               data_llam.Recordset.Edit
               data_llam.Recordset("realiza") = chiat.Value
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("realiza") = chiat.Value
            data_llam.Recordset.Update
         End If
         If IsNull(data_llam.Recordset("unied")) = False Then
            If data_llam.Recordset("unied") <> chcompre.Value Then
               data_llam.Recordset.Edit
               data_llam.Recordset("unied") = chcompre.Value
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("unied") = chcompre.Value
            data_llam.Recordset.Update
         End If
         If IsNull(data_llam.Recordset("tmm")) = False Then
            If data_llam.Recordset("tmm") <> chutiliza.Value Then
               data_llam.Recordset.Edit
               data_llam.Recordset("tmm") = chutiliza.Value
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("tmm") = chutiliza.Value
            data_llam.Recordset.Update
         End If
         If IsNull(data_llam.Recordset("ncobr")) = False Then
            If data_llam.Recordset("ncobr") <> chenf.Value Then
               data_llam.Recordset.Edit
               data_llam.Recordset("ncobr") = chenf.Value
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("ncobr") = chenf.Value
            data_llam.Recordset.Update
         End If
         If IsNull(data_llam.Recordset("enfer")) = False Then
            If data_llam.Recordset("enfer") <> chrella.Value Then
               data_llam.Recordset.Edit
               data_llam.Recordset("enfer") = chrella.Value
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("enfer") = chrella.Value
            data_llam.Recordset.Update
         End If
         If IsNull(data_llam.Recordset("base")) = False Then
            If data_llam.Recordset("base") <> chmedic.Value Then
               data_llam.Recordset.Edit
               data_llam.Recordset("base") = chmedic.Value
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("base") = chmedic.Value
            data_llam.Recordset.Update
         End If
         If t_obsenf.Text <> "" Then
            If IsNull(data_llam.Recordset("nombre")) = False Then
               If data_llam.Recordset("nombre") <> t_obsenf.Text Then
                  data_llam.Recordset.Edit
                  data_llam.Recordset("nombre") = t_obsenf.Text
                  data_llam.Recordset.Update
               End If
            Else
               data_llam.Recordset.Edit
               data_llam.Recordset("nombre") = t_obsenf.Text
               data_llam.Recordset.Update
            End If
         Else
            If IsNull(data_llam.Recordset("nombre")) = False Then
               data_llam.Recordset.Edit
               data_llam.Recordset("nombre") = Null
               data_llam.Recordset.Update
            End If
         End If
         If mhcom.Text <> "__:__" Then
            If IsNull(data_llam.Recordset("horsali")) = False Then
               If data_llam.Recordset("horsali") <> mhcom.Text Then
                  data_llam.Recordset.Edit
                  data_llam.Recordset("horsali") = mhcom.Text
                  data_llam.Recordset.Update
               End If
            Else
               data_llam.Recordset.Edit
               data_llam.Recordset("horsali") = mhcom.Text
               data_llam.Recordset.Update
            End If
            
         Else
            If IsNull(data_llam.Recordset("horsali")) = False Then
               data_llam.Recordset.Edit
               data_llam.Recordset("horsali") = Null
               data_llam.Recordset.Update
            End If
         End If
         If t_obsrella.Text <> "" Then
            If IsNull(data_llam.Recordset("motcon")) = False Then
               If data_llam.Recordset("motcon") <> t_obsrella.Text Then
                  data_llam.Recordset.Edit
                  data_llam.Recordset("motcon") = t_obsrella.Text
                  data_llam.Recordset.Update
               End If
            Else
               data_llam.Recordset.Edit
               data_llam.Recordset("motcon") = t_obsrella.Text
               data_llam.Recordset.Update
            End If
         Else
            If IsNull(data_llam.Recordset("motcon")) = False Then
               data_llam.Recordset.Edit
               data_llam.Recordset("motcon") = Null
               data_llam.Recordset.Update
            End If
         End If
         If t_obsmedic.Text <> "" Then
            If IsNull(data_llam.Recordset("motcance")) = False Then
               If data_llam.Recordset("motcance") <> t_obsmedic.Text Then
                  data_llam.Recordset.Edit
                  data_llam.Recordset("motcance") = t_obsmedic.Text
                  data_llam.Recordset.Update
               End If
            Else
               data_llam.Recordset.Edit
               data_llam.Recordset("motcance") = t_obsmedic.Text
               data_llam.Recordset.Update
            End If
         Else
            If IsNull(data_llam.Recordset("motcance")) = False Then
               data_llam.Recordset.Edit
               data_llam.Recordset("motcance") = Null
               data_llam.Recordset.Update
            End If
         End If
         If t_obs2.Text <> "" Then
            If IsNull(data_llam.Recordset("obsmot")) = False Then
               If data_llam.Recordset("obsmot") <> t_obs2.Text Then
                  data_llam.Recordset.Edit
                  data_llam.Recordset("obsmot") = t_obs2.Text
                  data_llam.Recordset.Update
               End If
            Else
               data_llam.Recordset.Edit
               data_llam.Recordset("obsmot") = t_obs2.Text
               data_llam.Recordset.Update
            End If
         Else
            If IsNull(data_llam.Recordset("obsmot")) = False Then
               data_llam.Recordset.Edit
               data_llam.Recordset("obsmot") = Null
               data_llam.Recordset.Update
            End If
         End If
         If IsNull(data_llam.Recordset("timdes")) = False Then
            If data_llam.Recordset("timdes") <> WElusuario Then
               data_llam.Recordset.Edit
               data_llam.Recordset("timdes") = Mid(WElusuario, 1, 15)
               data_llam.Recordset.Update
            End If
         Else
            data_llam.Recordset.Edit
            data_llam.Recordset("timdes") = Mid(WElusuario, 1, 15)
            data_llam.Recordset.Update
         End If
         If mhfin.Text <> "__:__" Then
            If IsNull(data_llam.Recordset("hor_rea")) = False Then
               If data_llam.Recordset("hor_rea") <> Format(Time, "HH:mm") Then
                  data_llam.Recordset.Edit
                  data_llam.Recordset("hor_rea") = Format(Time, "HH:mm")
                  data_llam.Recordset.Update
               End If
            Else
               data_llam.Recordset.Edit
               data_llam.Recordset("hor_rea") = Format(Time, "HH:mm")
               data_llam.Recordset.Update
            End If
            Xdeseacerrar = MsgBox("Desea CERRAR TODO EL LLAMADO COMO REALIZADO?", vbInformation + vbYesNo)
            If Xdeseacerrar = vbYes Then
               data_lla2.RecordSource = "Select * from llamado where nrolla =" & frm_largador.txt_nro.Text
               data_lla2.Refresh
               If data_lla2.Recordset.RecordCount > 0 Then
                  If data_lla2.Recordset("pend") = 4 Then
                     data_lla2.Recordset.Edit
                     data_lla2.Recordset("pend") = 2
                     data_lla2.Recordset("cmt_enproceso") = 2
                     data_lla2.Recordset("movilpas") = 2015
                     data_lla2.Recordset("fecpas") = Format(Date, "dd/mm/yyyy")
                     data_lla2.Recordset("horpas") = Format(Time, "HH:mm")
                     data_lla2.Recordset("fecsali") = Format(Date, "dd/mm/yyyy")
                     data_lla2.Recordset("horsali") = Format(Time, "HH:mm")
                     data_lla2.Recordset("fec_llega") = Format(Date, "dd/mm/yyyy")
                     data_lla2.Recordset("hor_llega") = Format(mhfin.Text, "HH:mm")
                     data_lla2.Recordset("fec_rea") = Format(Date, "dd/mm/yyyy")
                     data_lla2.Recordset("hor_rea") = Format(mhfin.Text, "HH:mm")
                     data_lla2.Recordset("diag") = "CMT " & Mid(t_obs2.Text, 1, 55)
                     data_lla2.Recordset("colormot") = data_lla2.Recordset("codmot")
                     If IsNull(data_lla2.Recordset("mes")) = False Then
                        If data_lla2.Recordset("mes") > 30 Then
                           data_lla2.Recordset("mes") = 0
                        End If
                     End If
                     'data_lla.Recordset("codmed") = txt_codmed.Text
                     'data_lla.Recordset("nommed") = dbcbomed.Text
                     data_lla2.Recordset.Update
                  Else
                    If data_lla2.Recordset("pend") <> 2 Then
                        data_lla2.Recordset.Edit
                        data_lla2.Recordset("pend") = 2
                        data_lla2.Recordset("cmt_enproceso") = 2
                        data_lla2.Recordset("movilpas") = 2015
                        data_lla2.Recordset("fecpas") = Format(Date, "dd/mm/yyyy")
                        data_lla2.Recordset("horpas") = Format(Time, "HH:mm")
                        data_lla2.Recordset("fecsali") = Format(Date, "dd/mm/yyyy")
                        data_lla2.Recordset("horsali") = Format(Time, "HH:mm")
                        data_lla2.Recordset("fec_llega") = Format(Date, "dd/mm/yyyy")
                        data_lla2.Recordset("hor_llega") = Format(mhfinno.Text, "HH:mm")
                        data_lla2.Recordset("fec_rea") = Format(Date, "dd/mm/yyyy")
                        data_lla2.Recordset("hor_rea") = Format(mhfinno.Text, "HH:mm")
                        data_lla2.Recordset("diag") = "CMT DERIVADO A BASE"
                        data_lla2.Recordset("colormot") = data_lla2.Recordset("codmot")
                        If IsNull(data_lla2.Recordset("mes")) = False Then
                           If data_lla2.Recordset("mes") > 30 Then
                              data_lla2.Recordset("mes") = 0
                           End If
                        End If
                         'data_lla.Recordset("codmed") = txt_codmed.Text
                         'data_lla.Recordset("nommed") = dbcbomed.Text
                        data_lla2.Recordset.Update
                         
                         MsgBox "ATENCION!! El llamado ya NO figura cómo pasado a CMT! Se pasará como terminado.", vbInformation
                     End If
                  End If
               End If
            Else
               MsgBox "El llamado continuará cómo PASADO A CMT"
            End If
         Else
            If IsNull(data_llam.Recordset("hor_rea")) = False Then
               data_llam.Recordset.Edit
               data_llam.Recordset("hor_rea") = Null
               data_llam.Recordset.Update
            End If
         End If
         data_agenda.RecordSource = "select * from t_fechas where fecha ='" & Format(Date, "dd/mm/yyyy") & "' and ced_pac ='" & Trim(labcedpac.Caption) & "' and cod_med =" & data_lla2.Recordset("codmedcmt")
         data_agenda.Refresh
         If data_agenda.Recordset.RecordCount > 0 Then
            If IsNull(data_agenda.Recordset("hora_realizacmt")) = False Then
               If Format(data_agenda.Recordset("hora_realizacmt"), "HH:mm") <> Format(mhfin.Text, "HH:mm") Then
                  data_agenda.Recordset.Edit
                  data_agenda.Recordset("hora_realizacmt") = mhfin.Text
                  data_agenda.Recordset.Update
               End If
            Else
               data_agenda.Recordset.Edit
               data_agenda.Recordset("hora_realizacmt") = mhfin.Text
               data_agenda.Recordset.Update
            End If
         Else
            MsgBox "No se pudo actualizar la agenda.", vbCritical
         End If
         If Xhcesi = vbYes And XcedDoc <> "" Then
            data_conshc.RecordSource = "select * from us where documento ='" & Trim(XcedDoc) & "'"
            data_conshc.Refresh
                        
            If Xcrear = 8 And data_conshc.Recordset.RecordCount > 0 Then
               If IsNull(data_conshc.Recordset("us_desc")) = False Then
                  XnroMedSapp = Val(data_conshc.Recordset("us_desc"))
                  data_medsapp.RecordSource = "select * from medicos where med_cod =" & XnroMedSapp
                  data_medsapp.Refresh
                  If data_medsapp.Recordset.RecordCount > 0 Then
                     If IsNull(data_medsapp.Recordset("med_nombre")) = False Then
                        XnomMedSapp = data_medsapp.Recordset("med_nombre")
                     Else
                        XnroMedSapp = 440
                        XnomMedSapp = "OTROS MEDICOS"
                     End If
                  Else
                     XnroMedSapp = 440
                     XnomMedSapp = "OTROS MEDICOS"
                  End If
               Else
                  XnroMedSapp = 440
                  XnomMedSapp = "OTROS MEDICOS"
               End If
               If IsNull(data_lla2.Recordset("codmed")) = False Then
                  If Val(data_lla2.Recordset("codmed")) <> Val(XnroMedSapp) Then
                     data_lla2.Recordset.Edit
                     data_lla2.Recordset("codmed") = XnroMedSapp
                     data_lla2.Recordset.Update
                   End If
               Else
                   data_lla2.Recordset.Edit
                   data_lla2.Recordset("codmed") = XnroMedSapp
                   data_lla2.Recordset.Update
               End If
               If IsNull(data_lla2.Recordset("nommed")) = False Then
                  If Trim(data_lla2.Recordset("nommed")) <> Trim(XnomMedSapp) Then
                     data_lla2.Recordset.Edit
                     data_lla2.Recordset("nommed") = Mid(XnomMedSapp, 1, 45)
                     data_lla2.Recordset.Update
                  End If
               Else
                  data_lla2.Recordset.Edit
                  data_lla2.Recordset("nommed") = Mid(XnomMedSapp, 1, 45)
                  data_lla2.Recordset.Update
               End If
               XEltextoobs = ""
               data_hce.RecordSource = "select * from cabezal_hc where cb_mat =" & frm_largador.txt_mat.Text
               data_hce.Refresh
               If data_hce.Recordset.RecordCount > 0 Then
                  data_hce.RecordSource = "select * from cabezal_hcdig where mat =" & frm_largador.txt_mat.Text
                  data_hce.Refresh
                  data_par.Recordset.Edit
                  data_par.Recordset("p_hc") = data_par.Recordset("p_hc") + 1
                  data_par.Recordset.Update
                  data_hce.Recordset.AddNew
                  data_hce.Recordset("id") = data_par.Recordset("p_hc")
                  data_hce.Recordset("hc_nro") = data_par.Recordset("p_hc")
                  data_hce.Recordset("mat") = frm_largador.txt_mat.Text
                  data_hce.Recordset("cednum") = frm_largador.txt_ced.Text
                  If frm_largador.t_codced.Text <> "" Then
                     data_hce.Recordset("cedtext") = frm_largador.txt_ced.Text & frm_largador.t_codced.Text
                     data_hce.Recordset("codced") = frm_largador.t_codced.Text
                  Else
                     data_hce.Recordset("cedtext") = frm_largador.txt_ced.Text
                     data_hce.Recordset("codced") = 0
                  End If
                  data_hce.Recordset("fecha") = Format(Date, "dd-mm-yyyy")
                  data_hce.Recordset("hora") = Format(Time, "HH:mm:ss")
                  data_hce.Recordset("codigo") = 3
                  data_hce.Recordset("tipo_cons") = 9
                  data_hce.Recordset("tipo_consd") = "Orientación Telefónica"
                  data_hce.Recordset("hc_base") = 19
                  data_hce.Recordset("hc_codmed") = data_conshc.Recordset("id")
                  data_hce.Recordset("hc_nommed") = data_conshc.Recordset("nombre") & " " & data_conshc.Recordset("apellidos")
                  data_hce.Recordset("hc_cpmed") = data_conshc.Recordset("cp")
                  If frm_largador.txt_edad.Text <> "" Then
                     data_hce.Recordset("hc_naca") = frm_largador.txt_edad.Text
                  End If
'                  adohc1.Recordset("hc_nacm") = Xwedm
'                  adohc1.Recordset("hc_nacd") = Xwedd
                  data_hce.Recordset.Update
      
                  data_hce.RecordSource = "Select * from hc_mcyotro where id =" & 529
                  data_hce.Refresh
                  data_hce.Recordset.AddNew
                  data_hce.Recordset("id") = data_par.Recordset("p_hc")
                  data_hce.Recordset("hc_nro") = data_par.Recordset("p_hc")
                  data_hce.Recordset("hc_mat") = frm_largador.txt_mat.Text
                  data_hce.Recordset("fecha") = Format(Date, "dd-mm-yyyy")
                  data_hce.Recordset("hora") = Format(Time, "HH:mm")
                  data_hce.Recordset("hc_mc") = "Orientación telefónica"
                  If t_obsenf.Text <> "" Then
                     XEltextoobs = "Curso probable de la enfermedad:" & t_obsenf.Text
                  End If
                  If t_obsrella.Text <> "" Then
                     If Trim(XEltextoobs) = "" Then
                        XEltextoobs = "Circunstancia de re-llamada: " & t_obsrella.Text
                     Else
                        XEltextoobs = XEltextoobs & vbCrLf & "Circunstancia de re-llamada: " & t_obsrella.Text
                     End If
                  End If
                  If t_obsmedic.Text <> "" Then
                     If Trim(XEltextoobs) = "" Then
                        XEltextoobs = "Prescripción de medicamentos:" & t_obsmedic.Text
                     Else
                        XEltextoobs = XEltextoobs & vbCrLf & "Prescripción de medicamentos:" & t_obsmedic.Text
                     End If
                  End If
                  If t_obs2.Text <> "" Then
                     If Trim(XEltextoobs) = "" Then
                        XEltextoobs = "EN SUMA: " & t_obs2.Text
                     Else
                        XEltextoobs = XEltextoobs & vbCrLf & "EN SUMA: " & t_obs2.Text
                     End If
                  End If
                  If Trim(XEltextoobs) <> "" Then
                     data_hce.Recordset("hc_otros") = XEltextoobs
                  Else
                     data_hce.Recordset("hc_otros") = "Sin Datos"
                  End If
                  data_hce.Recordset.Update

                  data_hce.RecordSource = "Select * from cli_crmdeudas where nrofact =" & data_par.Recordset("p_hc")
                  data_hce.Refresh
                  data_hce.Recordset.AddNew
                  data_hce.Recordset("id") = data_par.Recordset("p_hc")
                  data_hce.Recordset("base") = frm_largador.txt_mat.Text
                  data_hce.Recordset("nrofact") = data_par.Recordset("p_hc")
                  data_hce.Recordset("obs") = "registro de orientación clínica por vía telefónica"
                  data_hce.Recordset("usuario") = "Z719"
                  data_hce.Recordset("forma_pago") = 1
                  data_hce.Recordset("var1n") = 3
                  data_hce.Recordset.Update
   
                  data_hce.RecordSource = "Select * from cabezal_hcdig where id =" & data_par.Recordset("p_hc") & " and mat =" & frm_largador.txt_mat.Text
                  data_hce.Refresh
                  If data_hce.Recordset.RecordCount > 0 Then
                     If IsNull(data_hce.Recordset("hc_fin")) = True Then
                        data_hce.Recordset.Edit
                        data_hce.Recordset("hc_fin") = 5
                        data_hce.Recordset.Update
                     End If
                  End If
          
                  MsgBox "HC creada correctamente", vbInformation
               Else
                  MsgBox "No se encuentra registro de ficha en HCE, deberá crearla manualmente.", vbInformation
               End If
            Else
               MsgBox "No se encuentra número de cédula del médico o faltan datos del paciente, no se puede crear HCE", vbInformation
            End If
         End If
         
         Unload Me
      Else
         MsgBox "Tiene datos en la Opción de envío de recurso, VERIFIQUE!!"
      End If
   Else
      Xseguro = MsgBox("Desea SALIR sin GRABAR y SIN SELECCIONAR ninguna OPCION?", vbInformation + vbYesNo)
      If Xseguro = vbYes Then
         Unload Me
      Else
         MsgBox "Seleccione una opción de RECURSO y luego grabe los datos."
      End If
   End If
End If

Exit Sub

Verqueescmt:
            If Err.Number = 3155 Then
               MsgBox "3155 Verifique si hay datos para modificar o presione el botón de Cancelar"
            Else
               MsgBox "ERROR: " & Err.Description & " " & Err.Number

            End If
            

End Sub

Private Sub Command1_Click()
Dim Xseguro2 As String
Xseguro2 = MsgBox("Desea SALIR sin GRABAR?", vbInformation + vbYesNo)
If Xseguro2 = vbYes Then
   Unload Me
Else
   MsgBox "Seleccione una opción de RECURSO y luego grabe los datos."
End If

End Sub

Private Sub Command2_Click()
frm_cmt.PrintForm

End Sub

Private Sub Command3_Click()
data_lla2.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lla2.RecordSource = "Select * from llamado where nrolla =" & frm_largador.txt_nro.Text
data_lla2.Refresh
If data_lla2.Recordset.RecordCount > 0 Then
   If IsNull(data_lla2.Recordset("cmt_enproceso")) = False Then
      data_lla2.Recordset.Edit
      data_lla2.Recordset("cmt_enproceso") = Null
      data_lla2.Recordset("cmt_usproc") = Null
      data_lla2.Recordset.Update
      MsgBox "El llamado ha quedado habilitado para otro médico", vbInformation
   End If
End If
Unload Me


End Sub

Private Sub Command4_Click()
frm_seleccmt.Show vbModal

End Sub

Private Sub Form_Load()
Dim XcedDoc As String

data_hce.Connect = "odbc;dsn=sappnew;"

data_conshc.Connect = "odbc;dsn=sappnew;"

data_agenda.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_par.Connect = "odbc;dsn=sappnew;"
data_par.RecordSource = "param_gral"
data_par.Refresh

data_lla2.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_medsapp.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_llam.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_llam.RecordSource = "Select * from resplla where nro =" & frm_largador.txt_nro.Text
data_llam.Refresh
If data_llam.Recordset.RecordCount > 0 Then
   If frm_largador.txt_nomb.Text <> "" Then
      Label7.Caption = frm_largador.txt_nomb.Text
   Else
      Label7.Caption = ""
   End If
   If IsNull(data_llam.Recordset("hzona")) = False Then
      If IsNull(data_llam.Recordset("horsali")) = False Then
         mhcom.Text = data_llam.Recordset("horsali")
         mhcom.Enabled = False
      Else
         mhcom.Enabled = True
      End If
      If IsNull(data_llam.Recordset("mm")) = False Then
         If data_llam.Recordset("mm") = 1 Then
            Option1.Value = True
         Else
            If data_llam.Recordset("mm") = 2 Then
               Option2.Value = True
            Else
               If data_llam.Recordset("mm") = 3 Then
                  Option3.Value = True
               End If
            End If
         End If
      End If
      If IsNull(data_llam.Recordset("obsmot")) = False Then
         If Option1.Value = True Then
            t_obs1.Text = data_llam.Recordset("obsmot")
         Else
            If Option2.Value = True Then
               t_obs2.Text = data_llam.Recordset("obsmot")
            Else
               t_obs1.Text = data_llam.Recordset("obsmot")
            End If
         End If
      Else
         t_obs1.Text = ""
         t_obs2.Text = ""
      End If
      
      If IsNull(data_llam.Recordset("totend")) = False Then
         If data_llam.Recordset("totend") = "R" Then
            Combo1.ListIndex = 0
         Else
            If data_llam.Recordset("totend") = "A" Then
               Combo1.ListIndex = 1
            Else
               If data_llam.Recordset("totend") = "V" Then
                  Combo1.ListIndex = 2
               Else
                  Combo1.ListIndex = -1
               End If
            End If
         End If
      Else
         Combo1.ListIndex = -1
      End If
      If IsNull(data_llam.Recordset("hsald")) = False Then
         mhfinno.Text = Format(data_llam.Recordset("hsald"), "HH:mm")
      Else
         mhfinno.Text = "__:__"
      End If
      If IsNull(data_llam.Recordset("ncobr")) = False Then
         chenf.Value = data_llam.Recordset("ncobr")
      Else
         chenf.Value = 0
      End If
      If IsNull(data_llam.Recordset("enfer")) = False Then
         chrella.Value = data_llam.Recordset("enfer")
      Else
         chrella.Value = 0
      End If
      If IsNull(data_llam.Recordset("base")) = False Then
         chmedic.Value = data_llam.Recordset("base")
      Else
         chmedic.Value = 0
      End If
      If IsNull(data_llam.Recordset("nombre")) = False Then
         t_obsenf.Text = data_llam.Recordset("nombre")
      Else
         t_obsenf.Text = ""
      End If
      If IsNull(data_llam.Recordset("motcon")) = False Then
         t_obsrella.Text = data_llam.Recordset("motcon")
      Else
         t_obsrella.Text = ""
      End If
      If IsNull(data_llam.Recordset("motcance")) = False Then
         t_obsmedic.Text = data_llam.Recordset("motcance")
      Else
         t_obsmedic.Text = ""
      End If
      If IsNull(data_llam.Recordset("hor_rea")) = False Then
         mhfin.Text = Format(data_llam.Recordset("hor_rea"), "HH:mm")
      Else
         mhfin.Text = "__:__"
      End If
      If IsNull(data_llam.Recordset("codmed")) = False Then
         chclisim.Value = data_llam.Recordset("codmed")
      Else
         chclisim.Value = 0
      End If
      If IsNull(data_llam.Recordset("thh")) = False Then
         chclicamb.Value = data_llam.Recordset("thh")
      Else
         chclicamb.Value = 0
      End If
      If IsNull(data_llam.Recordset("ano")) = False Then
         chempeo.Value = data_llam.Recordset("ano")
      Else
         chempeo.Value = 0
      End If
      If IsNull(data_llam.Recordset("movtras")) = False Then
         chansie.Value = data_llam.Recordset("movtras")
      Else
         chansie.Value = 0
      End If
      If IsNull(data_llam.Recordset("trasla")) = False Then
         chaler.Value = data_llam.Recordset("trasla")
      Else
         chaler.Value = 0
      End If
      If IsNull(data_llam.Recordset("realiza")) = False Then
         chiat.Value = data_llam.Recordset("realiza")
      Else
         chiat.Value = 0
      End If
      If IsNull(data_llam.Recordset("unied")) = False Then
         chcompre.Value = data_llam.Recordset("unied")
      Else
         chcompre.Value = 0
      End If
      If IsNull(data_llam.Recordset("tmm")) = False Then
         chutiliza.Value = data_llam.Recordset("tmm")
      Else
         chutiliza.Value = 0
      End If
   Else
      MsgBox "No figura HORA de pasado a CMT, Avise a Telefonista que realice el pase nuevamente"
   End If
Else
   MsgBox "No figura llamado como pasado a CMT, Avise a Telefonista que realice el pase nuevamente"
End If

If mhcom.Text = "__:__" Then
   mhcom.Text = Format(Time, "HH:mm")
End If

If frm_largador.txt_ced.Text <> "" Then
   labcedpac.Caption = frm_largador.txt_ced.Text & frm_largador.t_codced.Text
Else
   labcedpac.Caption = ""
End If

If XWeltipoU = "ADMINISTRADOR" Or WElusuario = "HVENTRE" Or WElusuario = "VICTORIAM" Or WElusuario = "TMUÑOZ" Or XWeltipoU = "USUARIOS DESP" Then
   
Else
   b_graba.Enabled = False
   Frame1.Enabled = False
   mhcom.Enabled = False
   Command2.Enabled = False
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

Private Sub mhfin_GotFocus()
If mhfin.Text = "__:__" Then
   mhfin.Text = Format(Time, "HH:mm")
End If

End Sub

Private Sub mhfinno_GotFocus()
If mhfinno.Text = "__:__" Then
   mhfinno.Text = Format(Time, "HH:mm")
End If

End Sub

