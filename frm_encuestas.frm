VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_encuestas 
   BackColor       =   &H00400040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encuestas"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_encuestas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   11160
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_cli 
      Height          =   375
      Left            =   7200
      Top             =   8400
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "data_cli"
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
   Begin MSAdodcLib.Adodc data_encu 
      Height          =   375
      Left            =   7440
      Top             =   8040
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
      Caption         =   "data_encu"
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
   Begin VB.CommandButton b_inf 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4440
      Picture         =   "frm_encuestas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Informes"
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton b_busca 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      Picture         =   "frm_encuestas.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Consultar datos ingresados"
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      Picture         =   "frm_encuestas.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Cancelar acción de GRABAR"
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      Picture         =   "frm_encuestas.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_encuestas.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Modificar registro seleccionado"
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_encuestas.frx":1FF4
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Nuevo registro"
      Top             =   7920
      Width           =   495
   End
   Begin VB.Frame frame_pol 
      BackColor       =   &H0080FF80&
      Caption         =   "Atención en policlínica y Especialista"
      Height          =   6735
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   10695
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Se agendó vía web"
         Height          =   375
         Left            =   8280
         TabIndex        =   40
         Top             =   1320
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Se agendó vía telefónica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Left            =   6360
         TabIndex        =   38
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cbop7 
         Height          =   360
         ItemData        =   "frm_encuestas.frx":257E
         Left            =   7920
         List            =   "frm_encuestas.frx":2591
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   5400
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ComboBox cbop6 
         Height          =   360
         ItemData        =   "frm_encuestas.frx":25D0
         Left            =   7920
         List            =   "frm_encuestas.frx":25E3
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   4800
         Width           =   2535
      End
      Begin VB.ComboBox cbop5 
         Height          =   360
         ItemData        =   "frm_encuestas.frx":2622
         Left            =   7920
         List            =   "frm_encuestas.frx":2635
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   4200
         Width           =   2535
      End
      Begin VB.ComboBox cbop4 
         Height          =   360
         ItemData        =   "frm_encuestas.frx":2674
         Left            =   7920
         List            =   "frm_encuestas.frx":2687
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3600
         Width           =   2535
      End
      Begin VB.ComboBox cbop3 
         Height          =   360
         ItemData        =   "frm_encuestas.frx":26C6
         Left            =   7920
         List            =   "frm_encuestas.frx":26D9
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3000
         Width           =   2535
      End
      Begin VB.ComboBox cbop2 
         Height          =   360
         ItemData        =   "frm_encuestas.frx":2718
         Left            =   7920
         List            =   "frm_encuestas.frx":272B
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox t_obs 
         Height          =   735
         Left            =   2280
         TabIndex        =   24
         Top             =   5880
         Width           =   8175
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   5280
         TabIndex        =   17
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cbop1 
         Height          =   360
         ItemData        =   "frm_encuestas.frx":276A
         Left            =   7920
         List            =   "frm_encuestas.frx":277D
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox t_nom 
         Height          =   375
         Left            =   2040
         MaxLength       =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   6135
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Caption         =   "Consultar..."
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox t_mat 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin MSMask.MaskEdBox mf 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label labop7 
         BackColor       =   &H0080FFFF&
         Caption         =   "7)- SOLAMENTE: En caso de haber sido trasladado, cómo califica dicho traslado?"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   5400
         Visible         =   0   'False
         Width           =   7815
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080FFFF&
         Caption         =   "Recomendación o Sugerencias:"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Label labop6 
         BackColor       =   &H0080FFFF&
         Caption         =   "6)- Cómo califica la solución ofrecida a su estado de salud?"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   4800
         Width           =   7815
      End
      Begin VB.Label labop5 
         BackColor       =   &H0080FFFF&
         Caption         =   "5)- ¿Cómo califica la atención recibida por parte del médico?"
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   4200
         Width           =   7815
      End
      Begin VB.Label labop4 
         BackColor       =   &H0080FFFF&
         Caption         =   "4)- ¿Cómo califica el aspecto general de la policlínica donde fue atendido?"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   3600
         Width           =   7815
      End
      Begin VB.Label labop3 
         BackColor       =   &H0080FFFF&
         Caption         =   "3)- ¿Cómo califica el aspecto general del equipo asistencial (modales, ropa, aseo, etc)?"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   7815
      End
      Begin VB.Label labop2 
         BackColor       =   &H0080FFFF&
         Caption         =   "2)- ¿Cómo calificaría el tiempo de espera entre que llegó a la policlínica y el momento de su atención?"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   7815
      End
      Begin VB.Label labop1 
         BackColor       =   &H0080FFFF&
         Caption         =   "1)- ¿Cómo calificaría la atención recibida por el telefonista cuando Ud. solicitó la reserva para atención en policlínicas?"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   7815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Nombre:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Matrícula:"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label labus 
         Height          =   375
         Left            =   8280
         TabIndex        =   9
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Usuario:"
         Height          =   375
         Left            =   6840
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "HORA:"
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "FECHA:"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Datos de la encuesta"
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      Begin VB.CommandButton b_acep 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   8880
         Picture         =   "frm_encuestas.frx":27BC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frm_encuestas.frx":2D46
         Left            =   3360
         List            =   "frm_encuestas.frx":2D56
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Encuesta de atención en:"
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
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   3000
      Picture         =   "frm_encuestas.frx":2DAC
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   5280
      Picture         =   "frm_encuestas.frx":13F5A
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1695
   End
End
Attribute VB_Name = "frm_encuestas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_acep_Click()
If Combo1.ListIndex = 0 Then
   frame_pol.Visible = True
   labop7.Visible = False
   cbop7.Visible = False
   mf.Text = Date
   mh.Text = Format(Time, "HH:mm")
   labus.Caption = WElusuario
   t_mat.SetFocus
   Frame1.Enabled = False
   labop1.Caption = "1)- ¿Cómo calificaría la atención recibida por el telefonista cuando Ud. solicitó la reserva para atención en policlínicas?"
   labop2.Caption = "2)- ¿Cómo calificaría el tiempo de espera entre que llegó a la policlínica y el momento de su atención?"
   labop3.Caption = "3)- ¿Cómo califica el aspecto general del equipo asistencial (modales, ropa, aseo, etc)?"
   labop4.Caption = "4)- ¿Cómo califica el aspecto general de la policlínica donde fue atendido?"
   labop5.Caption = "5)- ¿Cómo califica la atención recibida por parte del médico?"
   labop6.Caption = "6)- Cómo califica la solución ofrecida a su estado de salud?"
   Option1.Visible = False
   Option2.Visible = False
Else
   If Combo1.ListIndex = 1 Then
      frame_pol.Visible = True
      labop1.Caption = "1)- ¿Cómo calificaría la atención recibida por el telefonista cuando Ud. solicitó la atención médica?"
'(1) Como calificaría la atención recibida por el telefonista cuando UD. solicito la atención médica?
      labop2.Caption = "2)- ¿Cómo calificaría el tiempo de llegada a su domicilio?"
'(2) Como calificaría el tiempo de llegada a su domicilio?
      labop3.Caption = "3)- ¿Cómo califica el aspecto general del equipo asistencial (modales, ropa, aseo, etc.)?"
'(3) Como califica el aspecto general del equipo asistencial (modales ,ropa ,aseo ,   etc )?
      labop4.Caption = "4)- ¿Cómo califica el aspecto general del móvil que concurrió a su domicilio?"
'(4) Como califica el aspecto general del móvil que concurrió a su domicilio ?
      labop5.Caption = "5)- ¿Cómo califica la atención recibida por parte del médico?"
      labop6.Caption = "6)- ¿Cómo califica la solución ofrecida a su estado de salud?"
      labop7.Visible = True
      cbop7.Visible = True
      mf.Text = Date
      mh.Text = Format(Time, "HH:mm")
      labus.Caption = WElusuario
      t_mat.SetFocus
      Frame1.Enabled = False
      Option1.Visible = False
      Option2.Visible = False
   Else
      If Combo1.ListIndex = 2 Then
         frame_pol.Visible = True
         labop7.Visible = False
         cbop7.Visible = False
         labop1.Caption = "1)- ¿Cómo calificaría la atención recibida por el telefonista cuando Ud. solicitó la atención médica?"
         labop2.Caption = "2)- ¿Cómo calificaría el tiempo de llegada al área protegida?"
         labop3.Caption = "3)- ¿Cómo califica el aspecto general del equipo asistencial (modales, ropa, aseo, etc.)?"
         labop4.Caption = "4)- ¿Cómo califica el aspecto general del móvil que concurrió al área protegida?"
         labop5.Caption = "5)- ¿Cómo califica la atención recibida por parte del médico?"
         labop6.Caption = "6)- ¿Cómo califica la solución ofrecida al problema de salud que motivó el llamado?"
         mf.Text = Date
         mh.Text = Format(Time, "HH:mm")
         labus.Caption = WElusuario
         t_mat.SetFocus
         Frame1.Enabled = False
         Option1.Visible = False
         Option2.Visible = False
      Else
         If Combo1.ListIndex = 3 Then
            Option1.Visible = True
            Option2.Visible = True
            frame_pol.Visible = True
            Option1.Value = True
            If Option1.Value = True Then
               labop1.Caption = "1)- ¿cómo calificaría la atención recibida por el telefonista cuando solicitó reserva para atención con especialista?"
            Else
               If Option2.Value = True Then
                  labop1.Caption = "1)- ¿cómo calificaría el sistema (ágil, dinámico, amigable)?"
               Else
                  labop1.Caption = "1)- ¿No seleccionó opción 1 o 2. Vuelva a reingresar encuesta?"
               End If
            End If
            labop2.Caption = "2)- ¿La hora con el especialista fue dada en un plazo considerado por usted razonable?"
            labop3.Caption = "3)- ¿cómo califica el aspecto general de la policlínica donde fue atendido?"
            labop4.Caption = "4)- ¿cómo califica la atención del personal de enfermería/administrativo de la policlínica?"
            labop5.Caption = "5)- ¿cómo califica la atención recibida por parte del médico?"
            labop6.Caption = "6)- ¿cómo califica la solución ofrecida a su estado de salud?"
            labop7.Visible = False
            cbop7.Visible = False
            mf.Text = Date
            mh.Text = Format(Time, "HH:mm")
            labus.Caption = WElusuario
            t_mat.SetFocus
            Frame1.Enabled = False
         Else
            labop7.Visible = False
            cbop7.Visible = False
            frame_pol.Visible = False
            Frame1.Enabled = True
            Option1.Visible = False
            Option2.Visible = False
            
         End If
      End If
   End If
End If

End Sub

Private Sub b_alta_Click()
XAlta = 1
data_encu.RecordSource = "select * from rrhh_sol where cl_codigo >=" & 8100 & " order by cl_codigo DESC"
data_encu.Refresh
If data_encu.Recordset.RecordCount > 0 Then
   Text1.Text = data_encu.Recordset("cl_codigo") + 1
Else
   Text1.Text = 1
End If
Frame1.Enabled = True
b_alta.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_cance.Enabled = True
b_busca.Enabled = False
b_inf.Enabled = False
Combo1.SetFocus

   
End Sub

Private Sub b_busca_Click()
frm_buscaencu.Show vbModal

End Sub

Private Sub b_cance_Click()
    frame_pol.Visible = True
    t_mat.Text = ""
    t_nom.Text = ""
    Text1.Text = ""
    cbop1.ListIndex = -1
    cbop2.ListIndex = -1
    cbop3.ListIndex = -1
    cbop4.ListIndex = -1
    cbop5.ListIndex = -1
    cbop6.ListIndex = -1
    If cbop7.Visible = True Then
       cbop7.ListIndex = -1
    End If
    t_obs.Text = ""
    frame_pol.Visible = False
    Combo1.ListIndex = -1
    b_acep.Enabled = True
    Frame1.Enabled = False
    b_alta.Enabled = True
    b_modif.Enabled = True
    b_graba.Enabled = False
    b_cance.Enabled = False
    b_busca.Enabled = True
    b_inf.Enabled = True

End Sub

Private Sub b_graba_Click()
Dim Xressigue As String
If mf.Text <> "__/__/____" Then
   If XAlta = 1 Then
      If t_mat.Text = "" Then
         t_mat.Text = 0
      End If
      data_encu.Recordset.AddNew
      data_encu.Recordset("cl_codigo") = Text1.Text
      data_encu.Recordset("cl_atrasoa") = Combo1.ListIndex
      data_encu.Recordset("cl_desc2") = Combo1.Text
      data_encu.Recordset("cl_fnac") = mf.Text
      data_encu.Recordset("cl_ruc") = mh.Text
      data_encu.Recordset("cl_descpag") = labus.Caption
      data_encu.Recordset("cl_nrovend") = t_mat.Text
      data_encu.Recordset("cl_desc1") = t_nom.Text
      data_encu.Recordset("cl_numero") = cbop1.ListIndex
      data_encu.Recordset("cl_zona") = cbop2.ListIndex
      data_encu.Recordset("cl_nomcobr") = cbop3.ListIndex
      data_encu.Recordset("cl_val1") = cbop4.ListIndex
      data_encu.Recordset("cl_val2") = cbop5.ListIndex
      data_encu.Recordset("cl_val3") = cbop6.ListIndex
      If cbop7.Visible = True Then
         data_encu.Recordset("cl_etiquet") = cbop7.ListIndex
      End If
      data_encu.Recordset("info_debit") = t_obs.Text
      data_encu.Recordset.Update
      t_mat.Text = ""
      t_nom.Text = ""
      Text1.Text = ""
      cbop1.ListIndex = -1
      cbop2.ListIndex = -1
      cbop3.ListIndex = -1
      cbop4.ListIndex = -1
      cbop5.ListIndex = -1
      cbop6.ListIndex = -1
      If cbop7.Visible = True Then
         cbop7.ListIndex = -1
      End If
      t_obs.Text = ""
      Xressigue = MsgBox("Desea continuar con la encuesta?", vbInformation + vbYesNo, "Encuestas")
      If Xressigue = vbYes Then
         XAlta = 1
         data_encu.RecordSource = "Select * from rrhh_sol order by cl_codigo"
         data_encu.Refresh
         If data_encu.Recordset.RecordCount > 0 Then
            data_encu.Recordset.MoveLast
            Text1.Text = data_encu.Recordset("cl_codigo") + 1
         Else
            Text1.Text = 1
         End If
         t_mat.SetFocus
      Else
         frame_pol.Visible = False
         Combo1.ListIndex = -1
         b_acep.Enabled = True
         Frame1.Enabled = False
         b_alta.Enabled = True
         b_modif.Enabled = True
         b_graba.Enabled = False
         b_cance.Enabled = False
         b_busca.Enabled = True
         b_inf.Enabled = True
         data_encu.Refresh
      End If
   Else
   
   End If
Else
   MsgBox "Ingrese Fecha"
End If

      
End Sub

Private Sub b_inf_Click()
frm_infencu.Show vbModal

End Sub

Private Sub b_modif_Click()
If frame_pol.Visible = True Then

Else
   MsgBox "Primero debe buscar un registro para modificar"
End If

End Sub

Private Sub cbop1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbop2.SetFocus
End If

End Sub

Private Sub cbop2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbop3.SetFocus
End If

End Sub

Private Sub cbop3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbop4.SetFocus
End If

End Sub

Private Sub cbop4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbop5.SetFocus
End If

End Sub

Private Sub cbop5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbop6.SetFocus
End If

End Sub

Private Sub cbop6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_obs.SetFocus
End If

End Sub

Private Sub Command1_Click()
frm_conscliencu.Show vbModal

End Sub

Private Sub Form_Load()
'data_encu.DatabaseName = App.Path & "\sapp.mdb"

data_encu.ConnectionString = "DSN=" & Xconexrmt

data_cli.ConnectionString = "DSN=" & Xconexrmt


End Sub

Private Sub Form_Resize()
With Image2
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
   labop1.Caption = "1)- ¿cómo calificaria la atención recibida por el telefonista cuando solicitó reserva para atención con especilista?"
Else
   If Option2.Value = True Then
      labop1.Caption = "1)- ¿cómo calificaría el sistema (ágil, dinámico, amigable)?"
   End If
End If

End Sub

Private Sub Option2_Click()
If Option1.Value = True Then
   labop1.Caption = "1)- ¿cómo calificaría la atención recibida por el telefonista cuando solicitó reserva para atención con especilista?"
Else
   If Option2.Value = True Then
      labop1.Caption = "1)- ¿cómo calificaría el sistema (ágil, dinámico, amigable)?"
   End If
End If

End Sub

Private Sub t_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_nom.SetFocus
End If

End Sub

Private Sub t_mat_LostFocus()
If t_mat.Text <> "" Then
   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & t_mat.Text
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      t_nom.Text = data_cli.Recordset("cl_apellid")
   Else
      t_nom.Text = "NO ENCONTRADO"
   End If
End If

End Sub

Private Sub t_nom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbop1.SetFocus
End If

End Sub

