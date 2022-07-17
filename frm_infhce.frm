VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infhce 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes para control de HCE MOVILES"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6000
   Icon            =   "frm_infhce.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6000
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin Crystal.CrystalReport cr1 
      Left            =   2280
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5280
      Picture         =   "frm_infhce.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_infhce.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos para el informe"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   1080
         Top             =   2520
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF0000&
         Caption         =   "Desde HCE"
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
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   3495
      End
      Begin MSAdodcLib.Adodc data_med 
         Height          =   375
         Left            =   2160
         Top             =   840
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
         Caption         =   "data_med"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc data_llam 
         Height          =   615
         Left            =   3120
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1085
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
         Caption         =   "data_llam"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2640
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox t_mov 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16744576
         ForeColor       =   65535
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16744576
         ForeColor       =   65535
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "MÉDICO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "MÓVIL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "FECHAS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1440
      Picture         =   "frm_infhce.frx":109E
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   2535
   End
End
Attribute VB_Name = "frm_infhce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
Dim Xcodelmed As Integer

MiBaseact.Execute "Delete * from inflla"
data_inf.RecordSource = "inflla"
data_inf.Refresh
If Combo1.ListIndex >= 0 Then
    data_med.RecordSource = "Select * from medicos where med_nombre ='" & Combo1.Text & "'"
    data_med.Refresh
    If data_med.Recordset.RecordCount > 0 Then
       If IsNull(data_med.Recordset("med_socnro")) = False Then
          Xcodelmed = data_med.Recordset("med_socnro")
       Else
          Xcodelmed = 0
       End If
    Else
       Xcodelmed = 0
    End If
Else
   Xcodelmed = 0
End If
If md.Text = "__/__/____" Or mh.Text = "__/__/____" Then
   MsgBox "Ingrese fechas"
Else
   frm_infhce.MousePointer = 11
   If t_mov.Text <> "" Then
      If Combo1.ListIndex >= 0 Then
         If Check1.Value = 1 Then
            If t_mov.Text = 99 Then
               If Xcodelmed = 0 Then
                  data_llam.RecordSource = "Select * from cabezal_hcdig where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
               Else
                  data_llam.RecordSource = "Select * from cabezal_hcdig where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and hc_codmed =" & Xcodelmed
               End If
            Else
               If Xcodelmed = 0 Then
                  data_llam.RecordSource = "Select * from cabezal_hcdig where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and hc_base =" & t_mov.Text
               Else
                  data_llam.RecordSource = "Select * from cabezal_hcdig where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and hc_codmed =" & Xcodelmed & " and hc_base =" & t_mov.Text
               End If
            End If
         Else
            data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nommed ='" & Combo1.Text & "' and movilpas =" & t_mov.Text & " order by fecha,hora"
         End If
         data_llam.Refresh
      Else
         If Check1.Value = 1 Then
            data_llam.RecordSource = "Select * from cabezal_hcdig where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and hc_base =" & t_mov.Text
         Else
            data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and movilpas =" & t_mov.Text & " order by fecha,hora"
         End If
         data_llam.Refresh
      End If
   Else
      MsgBox "No ingresó nro de móvil"
      data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and movilpas =" & 138 & " order by fecha,hora"
      data_llam.Refresh
   End If
   If data_llam.Recordset.RecordCount > 0 Then
      data_llam.Recordset.MoveFirst
      Do While Not data_llam.Recordset.EOF
         If Check1.Value = 1 Then
            data_inf.Recordset.AddNew
            data_inf.Recordset("fecha") = data_llam.Recordset("fecha")
            data_inf.Recordset("hora") = Mid(data_llam.Recordset("hora"), 1, 5)
            Adodc1.RecordSource = "Select * from clientes where cl_codigo =" & data_llam.Recordset("mat")
            Adodc1.Refresh
            If Adodc1.Recordset.RecordCount > 0 Then
               data_inf.Recordset("nombre") = Adodc1.Recordset("cl_apellid")
            Else
               data_inf.Recordset("nombre") = "Sin Datos"
            End If
            data_inf.Recordset("matric") = data_llam.Recordset("mat")
            data_inf.Recordset("ci") = Val(data_llam.Recordset("cedtext"))
            data_inf.Recordset("codmot") = Trim(str(data_llam.Recordset("codigo")))
            data_inf.Recordset("codmed") = data_llam.Recordset("hc_codmed")
            data_inf.Recordset("nommed") = data_llam.Recordset("hc_nommed")
            data_inf.Recordset("movilpas") = data_llam.Recordset("hc_base")
            data_inf.Recordset.Update
         Else
            data_inf.Recordset.AddNew
            data_inf.Recordset("fecha") = data_llam.Recordset("fecha")
            data_inf.Recordset("hora") = data_llam.Recordset("hora")
            data_inf.Recordset("nombre") = data_llam.Recordset("nombre")
            data_inf.Recordset("matric") = data_llam.Recordset("matric")
            data_inf.Recordset("categ") = data_llam.Recordset("categ")
            data_inf.Recordset("nomcat") = data_llam.Recordset("nomcat")
            data_inf.Recordset("ci") = data_llam.Recordset("ci")
            data_inf.Recordset("codmot") = data_llam.Recordset("codmot")
            data_inf.Recordset("colormot") = data_llam.Recordset("colormot")
            data_inf.Recordset("codmed") = data_llam.Recordset("codmed")
            data_inf.Recordset("nommed") = data_llam.Recordset("nommed")
            data_inf.Recordset("movilpas") = data_llam.Recordset("movilpas")
            data_inf.Recordset("lugar") = data_llam.Recordset("lugar")
            data_inf.Recordset.Update
         End If
         data_llam.Recordset.MoveNext
      Loop
      frm_infhce.MousePointer = 0
      MsgBox "Proceso terminado"
      data_inf.RecordSource = "select * from inflla"
      data_inf.Refresh
      cr1.ReportFileName = App.path & "\infctrolhcmov.rpt"
      If Check1.Value = 1 Then
         cr1.ReportTitle = "Informe de HCE realizadas M." & t_mov.Text & " desde: " & md.Text & " Hasta: " & mh.Text
      Else
         cr1.ReportTitle = "Informe de LLAMADOS desde: " & md.Text & " Hasta: " & mh.Text
      End If
      cr1.Action = 1
   
   End If
   frm_infhce.MousePointer = 0
      
End If

         
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()

data_med.ConnectionString = "dsn=" & Xconexrmt
data_med.RecordSource = "Select * from medicos order by med_nombre"
data_med.Refresh
If data_med.Recordset.RecordCount > 0 Then
   data_med.Recordset.MoveFirst
   Do While Not data_med.Recordset.EOF
      Combo1.AddItem data_med.Recordset("med_nombre")
      data_med.Recordset.MoveNext
   Loop
End If
data_llam.ConnectionString = "dsn=" & Xconexrmt
Adodc1.ConnectionString = "dsn=" & Xconexrmt

data_inf.DatabaseName = App.path & "\informes.mdb"


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_mov.SetFocus
End If

End Sub

Private Sub t_mov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub
