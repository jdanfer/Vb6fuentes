VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_buscaencu 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar encuentas"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_buscaencu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9720
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data1 
      Height          =   495
      Left            =   5640
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
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
      Caption         =   "data1"
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
   Begin MSFlexGridLib.MSFlexGrid DBgrid1 
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4895
      _Version        =   393216
      BackColorBkg    =   12615680
      SelectionMode   =   1
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8880
      Picture         =   "frm_buscaencu.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Ingrese la fecha con el siguiente formato : dd/mm/aaaa (Ej. 01/01/2012)"
      Top             =   600
      Width           =   5175
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "frm_buscaencu.frx":09CC
      Left            =   2640
      List            =   "frm_buscaencu.frx":09D6
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Doble click para editar"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Buscar por..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   5400
      Picture         =   "frm_buscaencu.frx":09F4
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1695
   End
End
Attribute VB_Name = "frm_buscaencu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
If XAlta = 85 Then
Else
    data1.RecordSource = "Select * from rrhh_sol where cl_nrovend =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 1) & " and cl_codigo =" & DBGrid1.TextMatrix(DBGrid1.RowSel, 4)
    data1.Refresh
    If data1.Recordset.RecordCount > 0 Then
       data1.Recordset.MoveFirst
        frm_encuestas.frame_pol.Visible = True
        If IsNull(data1.Recordset("cl_fnac")) = False Then
           frm_encuestas.mf.Text = data1.Recordset("cl_fnac")
        Else
           frm_encuestas.mf.Text = "__/__/____"
        End If
        If IsNull(data1.Recordset("cl_ruc")) = False Then
           frm_encuestas.mh.Text = data1.Recordset("cl_ruc")
        Else
           frm_encuestas.mh.Text = "__:__"
        End If
        If IsNull(data1.Recordset("cl_descpag")) = False Then
           frm_encuestas.labus.Caption = data1.Recordset("cl_descpag")
        Else
           frm_encuestas.labus.Caption = "S/U"
        End If
        If IsNull(data1.Recordset("cl_nrovend")) = False Then
           frm_encuestas.t_mat.Text = data1.Recordset("cl_nrovend")
        Else
           frm_encuestas.t_mat.Text = 0
        End If
        If IsNull(data1.Recordset("cl_desc2")) = False Then
           frm_encuestas.t_nom.Text = data1.Recordset("cl_desc2")
        Else
           frm_encuestas.t_nom.Text = ""
        End If
        
        If IsNull(data1.Recordset("cl_atrasoa")) = False Then
           frm_encuestas.Combo1.ListIndex = data1.Recordset("cl_atrasoa")
           If data1.Recordset("cl_atrasoa") = 1 Then
              frm_encuestas.cbop7.Visible = True
              frm_encuestas.labop7.Visible = True
           Else
              frm_encuestas.cbop7.Visible = False
              frm_encuestas.labop7.Visible = False
           End If
        Else
           frm_encuestas.Combo1.ListIndex = -1
        End If
        If IsNull(data1.Recordset("cl_numero")) = False Then
           frm_encuestas.cbop1.ListIndex = data1.Recordset("cl_numero")
        Else
           frm_encuestas.cbop1.ListIndex = -1
        End If
        If IsNull(data1.Recordset("cl_zona")) = False Then
           frm_encuestas.cbop2.ListIndex = data1.Recordset("cl_zona")
        Else
           frm_encuestas.cbop2.ListIndex = -1
        End If
        If IsNull(data1.Recordset("cl_nomcobr")) = False Then
           frm_encuestas.cbop3.ListIndex = data1.Recordset("cl_nomcobr")
        Else
           frm_encuestas.cbop3.ListIndex = -1
        End If
        If IsNull(data1.Recordset("cl_val1")) = False Then
           frm_encuestas.cbop4.ListIndex = data1.Recordset("cl_val1")
        Else
           frm_encuestas.cbop4.ListIndex = -1
        End If
        If IsNull(data1.Recordset("cl_val2")) = False Then
           frm_encuestas.cbop5.ListIndex = data1.Recordset("cl_val2")
        Else
           frm_encuestas.cbop5.ListIndex = -1
        End If
        If IsNull(data1.Recordset("cl_val3")) = False Then
           frm_encuestas.cbop6.ListIndex = data1.Recordset("cl_val3")
        Else
           frm_encuestas.cbop6.ListIndex = -1
        End If
        If frm_encuestas.cbop7.Visible = True Then
            If IsNull(data1.Recordset("cl_etiquet")) = False Then
               frm_encuestas.cbop7.ListIndex = data1.Recordset("cl_etiquet")
            Else
               frm_encuestas.cbop7.ListIndex = -1
            End If
        End If
        If IsNull(data1.Recordset("info_debit")) = False Then
           frm_encuestas.t_obs.Text = data1.Recordset("info_debit")
        Else
           frm_encuestas.t_obs.Text = ""
        End If
        
        frm_encuestas.frame_pol.Enabled = False
        
        If frm_encuestas.Combo1.ListIndex = 0 Then
           frm_encuestas.labop1.Caption = "1)- ¿Cómo calificaría la atención recibida por el telefonista cuando Ud. solicitó la reserva para atención en policlínicas?"
           frm_encuestas.labop2.Caption = "2)- ¿Cómo calificaría el tiempo de espera entre que llegó a la policlínica y el momento de su atención?"
           frm_encuestas.labop3.Caption = "3)- ¿Cómo califica el aspecto general del equipo asistencial (modales, ropa, aseo, etc)?"
           frm_encuestas.labop4.Caption = "4)- ¿Cómo califica el aspecto general de la policlínica donde fue atendido?"
           frm_encuestas.labop5.Caption = "5)- ¿Cómo califica la atención recibida por parte del médico?"
           frm_encuestas.labop6.Caption = "6)- Cómo califica la solución ofrecida a su estado de salud?"
        Else
           If frm_encuestas.Combo1.ListIndex = 1 Then
              frm_encuestas.frame_pol.Visible = True
              frm_encuestas.labop1.Caption = "1)- ¿Cómo calificaría la atención recibida por el telefonista cuando Ud. solicitó la atención médica?"
              frm_encuestas.labop2.Caption = "2)- ¿Cómo calificaría el tiempo de llegada a su domicilio?"
              frm_encuestas.labop3.Caption = "3)- ¿Cómo califica el aspecto general del equipo asistencial (modales, ropa, aseo, etc.)?"
              frm_encuestas.labop4.Caption = "4)- ¿Cómo califica el aspecto general del móvil que concurrió a su domicilio?"
              frm_encuestas.labop5.Caption = "5)- ¿Cómo califica la atención recibida por parte del médico?"
              frm_encuestas.labop6.Caption = "6)- ¿Cómo califica la solución ofrecida a su estado de salud?"
              frm_encuestas.labop7.Visible = True
              frm_encuestas.cbop7.Visible = True
           Else
              If frm_encuestas.Combo1.ListIndex = 2 Then
                 frm_encuestas.labop1.Caption = "1)- ¿Cómo calificaría la atención recibida por el telefonista cuando Ud. solicitó la atención médica?"
                 frm_encuestas.labop2.Caption = "2)- ¿Cómo calificaría el tiempo de llegada al área protegida?"
                 frm_encuestas.labop3.Caption = "3)- ¿Cómo califica el aspecto general del equipo asistencial (modales, ropa, aseo, etc.)?"
                 frm_encuestas.labop4.Caption = "4)- ¿Cómo califica el aspecto general del móvil que concurrió al área protegida?"
                 frm_encuestas.labop5.Caption = "5)- ¿Cómo califica la atención recibida por parte del médico?"
                 frm_encuestas.labop6.Caption = "6)- ¿Cómo califica la solución ofrecida al problema de salud que motivó el llamado?"
              Else
                 frm_encuestas.labop7.Visible = False
                 frm_encuestas.cbop7.Visible = False
                 frm_encuestas.frame_pol.Visible = False
                 frm_encuestas.Frame1.Enabled = True
              End If
           End If
        End If
    End If
End If
Unload Me
End Sub

Private Sub Form_Load()
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data1.ConnectionString = "dsn=" & Xconexrmt
DBGrid1.Rows = 2
DBGrid1.Cols = 5
DBGrid1.TextMatrix(0, 0) = "FECHA"
DBGrid1.ColWidth(0) = 1300
DBGrid1.TextMatrix(0, 1) = "CLIENTE"
DBGrid1.ColWidth(1) = 1500
DBGrid1.TextMatrix(0, 2) = "NOMBRE"
DBGrid1.ColWidth(2) = 3900
DBGrid1.TextMatrix(0, 3) = "ENCUESTA DE:"
DBGrid1.ColWidth(3) = 2500
DBGrid1.TextMatrix(0, 4) = "ID"
DBGrid1.ColWidth(4) = 400


End Sub

Private Sub Form_Resize()
With Image1
     .Top = 0
     .Left = 0
     .Height = Me.Height
     .Width = Me.Width
End With

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Text1.Text <> "" Then
      If Combo1.ListIndex = 0 Then
         data1.RecordSource = "Select * from rrhh_sol where cl_fnac >='" & Format(Text1.Text, "yyyy/mm/dd") & "' order by cl_fnac"
         data1.Refresh
      Else
         If Combo1.ListIndex = 1 Then
            data1.RecordSource = "Select * from rrhh_sol where cl_nrovend =" & Text1.Text
            data1.Refresh
         End If
      End If
   End If
    DBGrid1.Rows = 2
    DBGrid1.Cols = 5
    DBGrid1.TextMatrix(0, 0) = "FECHA"
    DBGrid1.ColWidth(0) = 1300
    DBGrid1.TextMatrix(0, 1) = "CLIENTE"
    DBGrid1.ColWidth(1) = 1500
    DBGrid1.TextMatrix(0, 2) = "NOMBRE"
    DBGrid1.ColWidth(2) = 3900
    DBGrid1.TextMatrix(0, 3) = "ENCUESTA DE:"
    DBGrid1.ColWidth(3) = 2500
    DBGrid1.TextMatrix(0, 4) = "ID"
    DBGrid1.ColWidth(4) = 400
    
    Dim Xcann As Integer
    Xcann = 1
    If data1.Recordset.RecordCount > 0 Then
        data1.Recordset.MoveFirst
        Do While Not data1.Recordset.EOF
           DBGrid1.TextMatrix(Xcann, 0) = data1.Recordset("cl_fnac")
           If IsNull(data1.Recordset("cl_nrovend")) = False Then
              DBGrid1.TextMatrix(Xcann, 1) = data1.Recordset("cl_nrovend")
           End If
           If IsNull(data1.Recordset("cl_desc1")) = False Then
              DBGrid1.TextMatrix(Xcann, 2) = data1.Recordset("cl_desc1")
           End If
           If IsNull(data1.Recordset("cl_desc2")) = False Then
              DBGrid1.TextMatrix(Xcann, 3) = data1.Recordset("cl_desc2")
           End If
           DBGrid1.TextMatrix(Xcann, 4) = data1.Recordset("cl_codigo")
           DBGrid1.Rows = DBGrid1.Rows + 1
           data1.Recordset.MoveNext
           Xcann = Xcann + 1
        Loop
    End If
   
   DBGrid1.SetFocus
End If

End Sub
