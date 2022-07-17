VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infdesp3 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes desde módulo DESPACHO"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infdesp3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10110
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr1 
      Left            =   4320
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_campos 
      Caption         =   "data_campos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton b_proc 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   8160
      Picture         =   "frm_infdesp3.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6240
      Width           =   615
   End
   Begin VB.CommandButton b_salir 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   9240
      Picture         =   "frm_infdesp3.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6240
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "Formato de Informe"
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
      Height          =   1095
      Left            =   240
      TabIndex        =   14
      Top             =   5160
      Width           =   9615
      Begin MSAdodcLib.Adodc data_llam 
         Height          =   375
         Left            =   6600
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         Caption         =   "data_llam"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF0000&
         Caption         =   "Resumen"
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
         Left            =   3360
         TabIndex        =   16
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Detalle"
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
         TabIndex        =   15
         Top             =   480
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Datos del informe"
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
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9615
      Begin VB.TextBox t_sel 
         BackColor       =   &H0080FFFF&
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
         Left            =   6480
         TabIndex        =   22
         Top             =   4200
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Selección primer filtro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   21
         Top             =   3960
         Width           =   2295
      End
      Begin MSMask.MaskEdBox mhh 
         Height          =   375
         Left            =   6600
         TabIndex        =   20
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mhd 
         Height          =   375
         Left            =   5640
         TabIndex        =   19
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton b_adddos 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   5640
         Picture         =   "frm_infdesp3.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4200
         Width           =   495
      End
      Begin VB.ListBox List3 
         Height          =   1740
         Left            =   6480
         TabIndex        =   12
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CommandButton b_adduno 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2520
         Picture         =   "frm_infdesp3.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4200
         Width           =   495
      End
      Begin VB.ListBox List2 
         Height          =   2460
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ListBox List1 
         Height          =   2460
         Left            =   3360
         TabIndex        =   8
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_infdesp3.frx":1BB2
         Left            =   2280
         List            =   "frm_infdesp3.frx":1BC2
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   3375
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         Caption         =   "PROCESANDO..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   3240
         TabIndex        =   23
         Top             =   2520
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF0000&
         Caption         =   "Filtros para el informe"
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
         Left            =   6480
         TabIndex        =   11
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Campos para el informe"
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
         Left            =   3360
         TabIndex        =   7
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Total de campos"
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
         TabIndex        =   6
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Selección de datos:"
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
         TabIndex        =   4
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "FECHAS // HORA"
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
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   480
      Picture         =   "frm_infdesp3.frx":1BFE
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   2775
   End
End
Attribute VB_Name = "frm_infdesp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_addtodo_Click()

End Sub

Private Sub b_adddos_Click()
Dim XX, Xban As Long
XX = 0
Xban = 0
On Error GoTo QuepasaPer22

If List3.ListCount > 3 Then
   MsgBox "Se excedió el límite de campos para el filtro"
Else
   If List1.List(List1.ListIndex) = "FECHA_RECIBIDO" Then
      MsgBox "El campo fecha ya está incluído por defecto"
   Else
      If List3.ListCount >= 1 Then
         For XX = 1 To List3.ListCount
             List3.ListIndex = XX - 1
             If List3.List(List3.ListIndex) = List1.List(List1.ListIndex) Then
                Xban = 1
             End If
         Next
      Else
         Xban = 0
      End If
      If List1.List(List1.ListIndex) <> "" And Xban <> 1 Then
         List3.AddItem List1.List(List1.ListIndex)
'      lper.RemoveItem lper.ListIndex
      End If
   End If
End If


Exit Sub

QuepasaPer22:
           If Err.Number = 3155 Then
              MsgBox "Error al grabar, vuelva a intentarlo"
           Else
              MsgBox "Error al agregar, verifique si seleccionó datos"
           End If



End Sub

Private Sub b_adduno_Click()
Dim XX, Xban As Long
XX = 0
Xban = 0
On Error GoTo QuepasaPer22

If List1.ListCount > 60 Then
   MsgBox "Se excedió el límite de campos para el reporte"
Else
   If List1.ListCount >= 1 Then
      For XX = 1 To List1.ListCount
          List1.ListIndex = XX - 1
          If List1.List(List1.ListIndex) = List2.List(List2.ListIndex) Then
             Xban = 1
          End If
      Next
   Else
      Xban = 0
   End If
   If List2.List(List2.ListIndex) <> "" And Xban <> 1 Then
      List1.AddItem List2.List(List2.ListIndex)
'      lper.RemoveItem lper.ListIndex
   End If
End If
List2.ListIndex = List2.ListIndex + 1

Exit Sub

QuepasaPer22:
           If Err.Number = 3155 Then
              MsgBox "Error al grabar, vuelva a intentarlo"
           Else
              MsgBox "Error al agregar, verifique si seleccionó datos"
           End If


End Sub

Private Sub b_proc_Click()
Dim Xcaden, Xmensaque, Xordena As String
Dim XX, Xban, Xtotregs As Long

Xtotregs = 0

Dim Xobjexel3 As Excel.Application
Dim Xlibexel3 As Excel.Workbook
Dim Xarchexel3 As New Excel.Worksheet

Dim Xlabrir As New Excel.Application

Dim Xarchtex As String
Dim XCol, Xlin, Xcanxdia, Xnrocan As Integer

On Error GoTo Xquepasainf33

List1.Visible = False
List3.Visible = False
Label6.Visible = True

List1.Enabled = False
List2.Enabled = False
List3.Enabled = False
frm_infdesp3.MousePointer = 11
b_proc.Enabled = False
b_salir.Enabled = False

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")

MiBaseact.Execute "Delete * from inflla"

data_inf.RecordSource = "inflla"
data_inf.Refresh

XX = 0
Xban = 0
Xcaden = ""
Xordena = ""
If md.Text <> "__/__/____" And mh.Text <> "__/__/____" And Combo1.ListIndex >= 0 Then
   If Option2.value = True Then
   Else
      Set Xobjexel3 = New Excel.Application
      Set Xlibexel3 = Xobjexel3.Workbooks.Add
      Set Xarchexel3 = Xlibexel3.Worksheets.Add
      Xarchexel3.Name = Trim(Combo1.Text)
      Xlibexel3.SaveAs ("C:\planillas\" & Trim(Combo1.Text) & "\" & Trim(Str(Month(md.Text))) & Trim(Str(Year(md.Text))) & ".xls")
      Xarchtex = "C:\planillas\" & Trim(Combo1.Text) & "\" & Trim(Str(Month(md.Text))) & Trim(Str(Year(md.Text))) & ".xls"
   End If
   If List1.ListCount > 0 Then
      For XX = 1 To List1.ListCount
          List1.ListIndex = XX - 1
          If Xcaden = "" Then
             data_campos.RecordSource = "Select * from campos where lista ='" & List1.List(List1.ListIndex) & "'"
             data_campos.Refresh
             If data_campos.Recordset.RecordCount > 0 Then
                Xcaden = data_campos.Recordset("campo")
             End If
          Else
             data_campos.RecordSource = "Select * from campos where lista ='" & List1.List(List1.ListIndex) & "'"
             data_campos.Refresh
             If data_campos.Recordset.RecordCount > 0 Then
                Xcaden = Xcaden + "," + data_campos.Recordset("campo")
             End If
          End If
          If List1.List(List1.ListIndex) = "FECHA_RECIBIDO" Or List1.List(List1.ListIndex) = "HORA_RECIBIDO" Then
             If List1.ListCount = 1 Then
             Else
                Xban = 8
             End If
          End If
      Next
      If Xban = 8 Then
         Xcaden = Xcaden + ",movilpas"
      Else
         Xcaden = Xcaden + ",fecha,hora,movilpas"
      End If
   Else
      MsgBox "Se incluyen todos los campos para el informe"
      data_campos.Recordset.MoveFirst
      Do While Not data_campos.Recordset.EOF
         If Xcaden = "" Then
            Xcaden = data_campos.Recordset("campo")
         Else
            Xcaden = Xcaden + "," + data_campos.Recordset("campo")
         End If
         data_campos.Recordset.MoveNext
      Loop
   End If
   XX = 0
   If Check1.value = 1 Then
      Xmensaque = t_sel.Text
      If List3.ListCount > 0 And Xmensaque <> "" Then
         For XX = 1 To List3.ListCount
             List3.ListIndex = XX - 1
             If Xordena = "" Then
                data_campos.RecordSource = "Select * from campos where lista ='" & List3.List(List3.ListIndex) & "'"
                data_campos.Refresh
                If data_campos.Recordset.RecordCount > 0 Then
                   Xordena = "fecha," & data_campos.Recordset("campo")
                End If
             Else
                data_campos.RecordSource = "Select * from campos where lista ='" & List3.List(List3.ListIndex) & "'"
                data_campos.Refresh
                If data_campos.Recordset.RecordCount > 0 Then
                   Xordena = Xordena + "," + data_campos.Recordset("campo")
                End If
             End If
         Next
         List3.ListIndex = 0
         data_campos.RecordSource = "Select * from campos where lista ='" & List3.List(List3.ListIndex) & "'"
         data_campos.Refresh
         If data_campos.Recordset.RecordCount > 0 Then
            If Combo1.ListIndex = 0 Then
               If data_campos.Recordset("lista") = "MATRICULA" Or _
                  data_campos.Recordset("lista") = "EDAD" Or _
                  data_campos.Recordset("lista") = "SEXO" Or _
                  data_campos.Recordset("lista") = "CEDULA" Or _
                  data_campos.Recordset("lista") = "COD_ZONA" Or _
                  data_campos.Recordset("lista") = "BASE" Or _
                  data_campos.Recordset("lista") = "MOVIL_LLAM" Or _
                  data_campos.Recordset("lista") = "MOVIL_TRASL" Or _
                  data_campos.Recordset("lista") = "COD_MEDICO" Or _
                  data_campos.Recordset("lista") = "TIPO_TRASL" Or _
                  data_campos.Recordset("lista") = "LLAM_CANCELADO" Or _
                  data_campos.Recordset("lista") = "COSTO_LLAM" Or _
                  data_campos.Recordset("lista") = "BOLETA_NRO" Then
                  data_llam.RecordSource = "Select " & Xcaden & " from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and " & data_campos.Recordset("campo") & " =" & Val(Xmensaque) & " and movilpas not in (99) order by " & Xordena
                  data_llam.Refresh
               Else
                  data_llam.RecordSource = "Select " & Xcaden & " from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and " & data_campos.Recordset("campo") & " ='" & Xmensaque & "' and movilpas not in (99) order by " & Xordena
                  data_llam.Refresh
               End If
            End If
            If Combo1.ListIndex = 1 Then
               If data_campos.Recordset("lista") = "MATRICULA" Or _
                  data_campos.Recordset("lista") = "EDAD" Or _
                  data_campos.Recordset("lista") = "SEXO" Or _
                  data_campos.Recordset("lista") = "CEDULA" Or _
                  data_campos.Recordset("lista") = "COD_ZONA" Or _
                  data_campos.Recordset("lista") = "BASE" Or _
                  data_campos.Recordset("lista") = "MOVIL_LLAM" Or _
                  data_campos.Recordset("lista") = "MOVIL_TRASL" Or _
                  data_campos.Recordset("lista") = "COD_MEDICO" Or _
                  data_campos.Recordset("lista") = "TIPO_TRASL" Or _
                  data_campos.Recordset("lista") = "LLAM_CANCELADO" Or _
                  data_campos.Recordset("lista") = "COSTO_LLAM" Or _
                  data_campos.Recordset("lista") = "BOLETA_NRO" Then
                  data_llam.RecordSource = "Select " & Xcaden & ",trasla from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and " & data_campos.Recordset("campo") & " =" & Val(Xmensaque) & " and trasla in (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15) order by " & Xordena
                  data_llam.Refresh
               Else
                  data_llam.RecordSource = "Select " & Xcaden & ",trasla from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and " & data_campos.Recordset("campo") & " ='" & Xmensaque & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15) order by " & Xordena
                  data_llam.Refresh
               End If
            End If
         
            If Combo1.ListIndex = 2 Then
               If data_campos.Recordset("lista") = "MATRICULA" Or _
                  data_campos.Recordset("lista") = "EDAD" Or _
                  data_campos.Recordset("lista") = "SEXO" Or _
                  data_campos.Recordset("lista") = "CEDULA" Or _
                  data_campos.Recordset("lista") = "COD_ZONA" Or _
                  data_campos.Recordset("lista") = "BASE" Or _
                  data_campos.Recordset("lista") = "MOVIL_LLAM" Or _
                  data_campos.Recordset("lista") = "MOVIL_TRASL" Or _
                  data_campos.Recordset("lista") = "COD_MEDICO" Or _
                  data_campos.Recordset("lista") = "TIPO_TRASL" Or _
                  data_campos.Recordset("lista") = "LLAM_CANCELADO" Or _
                  data_campos.Recordset("lista") = "COSTO_LLAM" Or _
                  data_campos.Recordset("lista") = "BOLETA_NRO" Then
                  data_llam.RecordSource = "Select " & Xcaden & ",categ from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and " & data_campos.Recordset("campo") & " =" & Val(Xmensaque) & " and categ in ('UDEMM','CERSEM','CERADU','CERDGI','CERSAP','CERHEV','CERCAS','CERMAT','CERKEV','CERIMP','CERSEV','CERVIS','CERANT','CERESS') and movilpas not in (99) order by " & Xordena
                  data_llam.Refresh
               Else
                  data_llam.RecordSource = "Select " & Xcaden & ",categ from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and " & data_campos.Recordset("campo") & " ='" & Xmensaque & "' and categ in ('UDEMM','CERSEM','CERADU','CERDGI','CERSAP','CERHEV','CERCAS','CERMAT','CERKEV','CERIMP','CERSEV','CERVIS','CERANT','CERESS') and movilpas not in (99) order by " & Xordena
                  data_llam.Refresh
               End If
            End If
         
            If Combo1.ListIndex = 3 Then
               If data_campos.Recordset("lista") = "MATRICULA" Or _
                  data_campos.Recordset("lista") = "EDAD" Or _
                  data_campos.Recordset("lista") = "SEXO" Or _
                  data_campos.Recordset("lista") = "CEDULA" Or _
                  data_campos.Recordset("lista") = "COD_ZONA" Or _
                  data_campos.Recordset("lista") = "BASE" Or _
                  data_campos.Recordset("lista") = "MOVIL_LLAM" Or _
                  data_campos.Recordset("lista") = "MOVIL_TRASL" Or _
                  data_campos.Recordset("lista") = "COD_MEDICO" Or _
                  data_campos.Recordset("lista") = "TIPO_TRASL" Or _
                  data_campos.Recordset("lista") = "LLAM_CANCELADO" Or _
                  data_campos.Recordset("lista") = "COSTO_LLAM" Or _
                  data_campos.Recordset("lista") = "BOLETA_NRO" Then
                  data_llam.RecordSource = "Select " & Xcaden & ",enfer from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and " & data_campos.Recordset("campo") & " =" & Val(Xmensaque) & " and enfer in (1) and movilpas not in (99) order by " & Xordena
                  data_llam.Refresh
               Else
                  data_llam.RecordSource = "Select " & Xcaden & ",enfer from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and " & data_campos.Recordset("campo") & " ='" & Xmensaque & "' and enfer in (1) and movilpas not in (99) order by " & Xordena
                  data_llam.Refresh
               End If
            End If
         
'in ('UDEMM','CERSEM','CERADU','CERDGI','CERSAP','CERHEV','CERCAS','CERMAT','CERKEV','CERIMP','CERSEV','CERVIS','CERANT','CERESS')
         Else
            MsgBox "No existen campos para filtrar"
         End If
      Else
         MsgBox "Seleccione los datos a filtrar"
      End If
   Else
      If List3.ListCount > 0 Then
         For XX = 1 To List3.ListCount
             List3.ListIndex = XX - 1
             If Xordena = "" Then
                data_campos.RecordSource = "Select * from campos where lista ='" & List3.List(List3.ListIndex) & "'"
                data_campos.Refresh
                If data_campos.Recordset.RecordCount > 0 Then
                   Xordena = "fecha," & data_campos.Recordset("campo")
                End If
             Else
                data_campos.RecordSource = "Select * from campos where lista ='" & List3.List(List3.ListIndex) & "'"
                data_campos.Refresh
                If data_campos.Recordset.RecordCount > 0 Then
                   Xordena = Xordena + "," + data_campos.Recordset("campo")
                End If
             End If
         Next
         List3.ListIndex = 0
      Else
         Xordena = "fecha"
      End If
      If Combo1.ListIndex = 0 Then
         data_llam.RecordSource = "Select " & Xcaden & " from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and movilpas not in (99) order by " & Xordena
         data_llam.Refresh
      End If
      If Combo1.ListIndex = 1 Then
         data_llam.RecordSource = "Select " & Xcaden & ",trasla from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15) order by " & Xordena
         data_llam.Refresh
      End If
      If Combo1.ListIndex = 2 Then
         data_llam.RecordSource = "Select " & Xcaden & ",categ from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and categ in ('UDEMM','CERSEM','CERADU','CERDGI','CERSAP','CERHEV','CERCAS','CERMAT','CERKEV','CERIMP','CERSEV','CERVIS','CERANT','CERESS') and movilpas not in (99) order by " & Xordena
         data_llam.Refresh
      End If
      If Combo1.ListIndex = 3 Then
         data_llam.RecordSource = "Select " & Xcaden & ",enfer from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and enfer in (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15) and movilpas not in (99) order by " & Xordena
         data_llam.Refresh
      End If
      
   End If
   If data_llam.Recordset.RecordCount > 0 Then
      If mhd.Text = "__:__" And mhh.Text = "__:__" Then
         XX = 0
         If Option2.value = True Then
         Else
            Xcanxdia = 0
            Xlin = 1
            XCol = 1
            Xarchexel3.Cells(Xlin, XCol) = "CENTRO DE COMPUTOS DE SAPP"
            XCol = 6
            Xarchexel3.Cells(Xlin, XCol) = "FECHA ACTUAL:"
            XCol = 7
            Xarchexel3.Cells(Xlin, XCol) = Format(Date, "dd/mm/yyyy")
            Xlin = Xlin + 1
            XCol = 2
            Xarchexel3.Range("A1", "C3").Font.Size = 16
            Xarchexel3.Cells(Xlin, XCol) = "PLANILLA DE " & Trim(Combo1.Text) & " DESDE: " & md.Text & " HASTA: " & mh.Text
            Xarchexel3.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 200)
            
            XCol = 1
            Xlin = Xlin + 2
            Xnrocan = Xnrocan + Xlin
            
            For XX = 1 To List1.ListCount
                List1.ListIndex = XX - 1
                data_campos.RecordSource = "Select * from campos where lista ='" & List1.List(List1.ListIndex) & "'"
                data_campos.Refresh
                If XX = 1 Then
                   Xarchexel3.Range("A" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 2 Then
                   Xarchexel3.Range("B" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 3 Then
                   Xarchexel3.Range("C" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 4 Then
                   Xarchexel3.Range("D" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 5 Then
                   Xarchexel3.Range("E" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 6 Then
                   Xarchexel3.Range("F" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 7 Then
                   Xarchexel3.Range("G" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 8 Then
                   Xarchexel3.Range("H" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 9 Then
                   Xarchexel3.Range("I" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 10 Then
                   Xarchexel3.Range("J" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                
                Xarchexel3.Cells(Xlin, XCol) = List1.List(List1.ListIndex)
                XCol = XCol + 1
            Next
         End If
         data_llam.Recordset.MoveFirst
         Xlin = Xlin + 1
         XCol = 1
         
         Do While Not data_llam.Recordset.EOF
            If Option2.value = True Then
               XX = 0
               data_inf.Recordset.AddNew
               data_inf.Recordset("fecha") = data_llam.Recordset("fecha")
               For XX = 1 To List3.ListCount
                   List3.ListIndex = XX - 1
                   data_campos.RecordSource = "Select * from campos where lista ='" & List3.List(List3.ListIndex) & "'"
                   data_campos.Refresh
                   If data_campos.Recordset.RecordCount > 0 Then
                      If XX = 1 Then
                         If data_campos.Recordset("lista") = "SEXO" Then
                            If IsNull(data_llam.Recordset(Trim(data_campos.Recordset("campo")))) = False Then
                               If data_llam.Recordset(Trim(data_campos.Recordset("campo"))) = 0 Then
                                  data_inf.Recordset("nombre") = "MASCULINO"
                               Else
                                  data_inf.Recordset("nombre") = "FEMENINO"
                               End If
                            Else
                               data_inf.Recordset("nombre") = "FEMENINO"
                            End If
                         Else
                            If data_campos.Recordset("lista") = "MATRICULA" Or _
                               data_campos.Recordset("lista") = "EDAD" Or _
                               data_campos.Recordset("lista") = "SEXO" Or _
                               data_campos.Recordset("lista") = "CEDULA" Or _
                               data_campos.Recordset("lista") = "COD_ZONA" Or _
                               data_campos.Recordset("lista") = "BASE" Or _
                               data_campos.Recordset("lista") = "MOVIL_LLAM" Or _
                               data_campos.Recordset("lista") = "MOVIL_TRASL" Or _
                               data_campos.Recordset("lista") = "COD_MEDICO" Or _
                               data_campos.Recordset("lista") = "TIPO_TRASL" Or _
                               data_campos.Recordset("lista") = "LLAM_CANCELADO" Or _
                               data_campos.Recordset("lista") = "COSTO_LLAM" Or _
                               data_campos.Recordset("lista") = "BOLETA_NRO" Then
                               data_inf.Recordset("nombre") = Trim(Str(data_llam.Recordset(Trim(data_campos.Recordset("campo")))))
                            Else
                               data_inf.Recordset("nombre") = Mid(data_llam.Recordset(Trim(data_campos.Recordset("campo"))), 1, 70)
                            End If
                         End If
                         data_inf.Recordset("motmov") = List3.List(List3.ListIndex)
                      End If
                      If XX = 2 Then
                         If data_campos.Recordset("lista") = "SEXO" Then
                            If IsNull(data_llam.Recordset(Trim(data_campos.Recordset("campo")))) = False Then
                               If data_llam.Recordset(Trim(data_campos.Recordset("campo"))) = 0 Then
                                  data_inf.Recordset("motcon") = "MASCULINO"
                               Else
                                  data_inf.Recordset("motcon") = "FEMENINO"
                               End If
                            Else
                               data_inf.Recordset("motcon") = "FEMENINO"
                            End If
                         Else
                            If data_campos.Recordset("lista") = "MATRICULA" Or _
                               data_campos.Recordset("lista") = "EDAD" Or _
                               data_campos.Recordset("lista") = "SEXO" Or _
                               data_campos.Recordset("lista") = "CEDULA" Or _
                               data_campos.Recordset("lista") = "COD_ZONA" Or _
                               data_campos.Recordset("lista") = "BASE" Or _
                               data_campos.Recordset("lista") = "MOVIL_LLAM" Or _
                               data_campos.Recordset("lista") = "MOVIL_TRASL" Or _
                               data_campos.Recordset("lista") = "COD_MEDICO" Or _
                               data_campos.Recordset("lista") = "TIPO_TRASL" Or _
                               data_campos.Recordset("lista") = "LLAM_CANCELADO" Or _
                               data_campos.Recordset("lista") = "COSTO_LLAM" Or _
                               data_campos.Recordset("lista") = "BOLETA_NRO" Then
                               data_inf.Recordset("motcon") = Trim(Str(data_llam.Recordset(Trim(data_campos.Recordset("campo")))))
                            Else
                               data_inf.Recordset("motcon") = Mid(data_llam.Recordset(Trim(data_campos.Recordset("campo"))), 1, 100)
                            End If
                         End If
                         data_inf.Recordset("motcance") = List3.List(List3.ListIndex)
                      End If
                      If XX = 3 Then
                         If data_campos.Recordset("lista") = "SEXO" Then
                            If IsNull(data_llam.Recordset(Trim(data_campos.Recordset("campo")))) = False Then
                               If data_llam.Recordset(Trim(data_campos.Recordset("campo"))) = 0 Then
                                  data_inf.Recordset("nomcat") = "MASCULINO"
                               Else
                                  data_inf.Recordset("nomcat") = "FEMENINO"
                               End If
                            Else
                               data_inf.Recordset("nomcat") = "FEMENINO"
                            End If
                         Else
                            If data_campos.Recordset("lista") = "MATRICULA" Or _
                               data_campos.Recordset("lista") = "EDAD" Or _
                               data_campos.Recordset("lista") = "SEXO" Or _
                               data_campos.Recordset("lista") = "CEDULA" Or _
                               data_campos.Recordset("lista") = "COD_ZONA" Or _
                               data_campos.Recordset("lista") = "BASE" Or _
                               data_campos.Recordset("lista") = "MOVIL_LLAM" Or _
                               data_campos.Recordset("lista") = "MOVIL_TRASL" Or _
                               data_campos.Recordset("lista") = "COD_MEDICO" Or _
                               data_campos.Recordset("lista") = "TIPO_TRASL" Or _
                               data_campos.Recordset("lista") = "LLAM_CANCELADO" Or _
                               data_campos.Recordset("lista") = "COSTO_LLAM" Or _
                               data_campos.Recordset("lista") = "BOLETA_NRO" Then
                               data_inf.Recordset("nomcat") = Trim(Str(data_llam.Recordset(Trim(data_campos.Recordset("campo")))))
                            Else
                               data_inf.Recordset("nomcat") = Mid(data_llam.Recordset(Trim(data_campos.Recordset("campo"))), 1, 50)
                            End If
                         End If
                         data_inf.Recordset("lugar") = List3.List(List3.ListIndex)
                      End If
                   End If
               Next
               data_inf.Recordset.Update
'               If data_llam.Recordset.EOF = True Then
'               Else
'                  data_llam.Recordset.MoveNext
'               End If
            Else
               For XX = 1 To List1.ListCount
                   List1.ListIndex = XX - 1
                   data_campos.RecordSource = "Select * from campos where lista ='" & List1.List(List1.ListIndex) & "'"
                   data_campos.Refresh
                   If data_campos.Recordset.RecordCount > 0 Then
                      XCol = XX
                      If data_campos.Recordset("lista") = "FECHA_RECIBIDO" Then
                         Xarchexel3.Cells(Xlin, XCol) = Format(data_llam.Recordset(Trim(data_campos.Recordset("campo"))), "mm/dd/yyyy")
                      Else
                         If data_campos.Recordset("lista") = "SEXO" Then
                            If IsNull(data_llam.Recordset(Trim(data_campos.Recordset("campo")))) = False Then
                               If data_llam.Recordset(Trim(data_campos.Recordset("campo"))) = 0 Then
                                  Xarchexel3.Cells(Xlin, XCol) = "MASCULINO"
                               Else
                                  Xarchexel3.Cells(Xlin, XCol) = "FEMENINO"
                               End If
                            Else
                               Xarchexel3.Cells(Xlin, XCol) = "FEMENINO"
                            End If
                         Else
                            Xarchexel3.Cells(Xlin, XCol) = data_llam.Recordset(Trim(data_campos.Recordset("campo")))
                         End If
                      End If
                   End If
               Next
            End If
            Xtotregs = Xtotregs + 1
            If data_llam.Recordset.EOF = True Then
            Else
               data_llam.Recordset.MoveNext
            End If
            Xlin = Xlin + 1
         Loop
         Xlin = Xlin + 1
         XCol = 2
         If Option2.value = True Then
         Else
            Xarchexel3.Cells(Xlin, XCol) = "TOTAL DE REGISTROS: " & Xtotregs
         End If
         List1.ListIndex = 0
         List1.Enabled = True
         List2.Enabled = True
         List3.Enabled = True
         b_proc.Enabled = False
         b_salir.Enabled = False
         frm_infdesp3.MousePointer = 0
         MsgBox "Proceso terminado"
         If Option2.value = True Then
            If List3.ListCount = 1 Then
               cr1.ReportFileName = App.Path & "\infdespn1.rpt"
               cr1.ReportTitle = "Informe de " & Combo1.Text & " desde: " & md.Text & " hasta: " & mh.Text
               cr1.Action = 1
            Else
               If List3.ListCount = 2 Then
                  cr1.ReportFileName = App.Path & "\infdespn2.rpt"
                  cr1.ReportTitle = "Informe de " & Combo1.Text & " desde: " & md.Text & " hasta: " & mh.Text
                  cr1.Action = 1
               Else
                  cr1.ReportFileName = App.Path & "\infdespn3.rpt"
                  cr1.ReportTitle = "Informe de " & Combo1.Text & " desde: " & md.Text & " hasta: " & mh.Text
                  cr1.Action = 1
               End If
            End If
         Else
            Xlibexel3.Save
            Xlibexel3.Close
            Xobjexel3.Quit
            
'            Shell frm_menu.data_usuac.Recordset("destino") & "excel.exe " & Xarchtex, vbMaximizedFocus
            Xlabrir.Workbooks.Open Xarchtex, , False
            Xlabrir.Visible = True
            Xlabrir.WindowState = xlMaximized
         End If
      Else
         XX = 0
         Xcanxdia = 0
         Xlin = 1
         XCol = 1
         If Option2.value = True Then
         Else
            Xarchexel3.Cells(Xlin, XCol) = "CENTRO DE COMPUTOS DE SAPP"
            XCol = 6
            Xarchexel3.Cells(Xlin, XCol) = "FECHA ACTUAL:"
            XCol = 7
            Xarchexel3.Cells(Xlin, XCol) = Format(Date, "dd/mm/yyyy")
            Xlin = Xlin + 1
            XCol = 2
            Xarchexel3.Range("A1", "C3").Font.Size = 16
            Xarchexel3.Cells(Xlin, XCol) = "PLANILLA DE " & Trim(Combo1.Text) & " DESDE: " & md.Text & " HASTA: " & mh.Text
            Xarchexel3.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 200)
            XCol = 1
            Xlin = Xlin + 2
            Xnrocan = Xnrocan + Xlin
            For XX = 1 To List1.ListCount
                List1.ListIndex = XX - 1
                data_campos.RecordSource = "Select * from campos where lista ='" & List1.List(List1.ListIndex) & "'"
                data_campos.Refresh
                If XX = 1 Then
                   Xarchexel3.Range("A" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 2 Then
                   Xarchexel3.Range("B" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 3 Then
                   Xarchexel3.Range("C" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 4 Then
                   Xarchexel3.Range("D" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 5 Then
                   Xarchexel3.Range("E" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 6 Then
                   Xarchexel3.Range("F" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 7 Then
                   Xarchexel3.Range("G" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 8 Then
                   Xarchexel3.Range("H" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 9 Then
                   Xarchexel3.Range("I" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                If XX = 10 Then
                   Xarchexel3.Range("J" & Trim(Str(Xlin))).ColumnWidth = 15
                End If
                
                Xarchexel3.Cells(Xlin, XCol) = List1.List(List1.ListIndex)
                XCol = XCol + 1
            Next
         End If
         data_llam.Recordset.MoveFirst
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_llam.Recordset.EOF
            If data_llam.Recordset("hora") >= mhd.Text And data_llam.Recordset("hora") <= mhh.Text Then
                If Option2.value = True Then
                   XX = 0
                   data_inf.Recordset.AddNew
                   data_inf.Recordset("fecha") = data_llam.Recordset("fecha")
                   For XX = 1 To List3.ListCount
                       List3.ListIndex = XX - 1
                       data_campos.RecordSource = "Select * from campos where lista ='" & List3.List(List3.ListIndex) & "'"
                       data_campos.Refresh
                       If data_campos.Recordset.RecordCount > 0 Then
                          If XX = 1 Then
                             If data_campos.Recordset("lista") = "SEXO" Then
                                If IsNull(data_llam.Recordset(Trim(data_campos.Recordset("campo")))) = False Then
                                   If data_llam.Recordset(Trim(data_campos.Recordset("campo"))) = 0 Then
                                      data_inf.Recordset("nombre") = "MASCULINO"
                                   Else
                                      data_inf.Recordset("nombre") = "FEMENINO"
                                   End If
                                Else
                                   data_inf.Recordset("nombre") = "FEMENINO"
                                End If
                             Else
                                If data_campos.Recordset("lista") = "MATRICULA" Or _
                                   data_campos.Recordset("lista") = "EDAD" Or _
                                   data_campos.Recordset("lista") = "SEXO" Or _
                                   data_campos.Recordset("lista") = "CEDULA" Or _
                                   data_campos.Recordset("lista") = "COD_ZONA" Or _
                                   data_campos.Recordset("lista") = "BASE" Or _
                                   data_campos.Recordset("lista") = "MOVIL_LLAM" Or _
                                   data_campos.Recordset("lista") = "MOVIL_TRASL" Or _
                                   data_campos.Recordset("lista") = "COD_MEDICO" Or _
                                   data_campos.Recordset("lista") = "TIPO_TRASL" Or _
                                   data_campos.Recordset("lista") = "LLAM_CANCELADO" Or _
                                   data_campos.Recordset("lista") = "COSTO_LLAM" Or _
                                   data_campos.Recordset("lista") = "BOLETA_NRO" Then
                                   data_inf.Recordset("nombre") = Trim(Str(data_llam.Recordset(Trim(data_campos.Recordset("campo")))))
                                Else
                                   data_inf.Recordset("nombre") = Mid(data_llam.Recordset(Trim(data_campos.Recordset("campo"))), 1, 70)
                                End If
                             End If
                             data_inf.Recordset("motmov") = List3.List(List3.ListIndex)
                          End If
                          If XX = 2 Then
                             If data_campos.Recordset("lista") = "SEXO" Then
                                If IsNull(data_llam.Recordset(Trim(data_campos.Recordset("campo")))) = False Then
                                   If data_llam.Recordset(Trim(data_campos.Recordset("campo"))) = 0 Then
                                      data_inf.Recordset("motcon") = "MASCULINO"
                                   Else
                                      data_inf.Recordset("motcon") = "FEMENINO"
                                   End If
                                Else
                                   data_inf.Recordset("motcon") = "FEMENINO"
                                End If
                             Else
                                If data_campos.Recordset("lista") = "MATRICULA" Or _
                                   data_campos.Recordset("lista") = "EDAD" Or _
                                   data_campos.Recordset("lista") = "SEXO" Or _
                                   data_campos.Recordset("lista") = "CEDULA" Or _
                                   data_campos.Recordset("lista") = "COD_ZONA" Or _
                                   data_campos.Recordset("lista") = "BASE" Or _
                                   data_campos.Recordset("lista") = "MOVIL_LLAM" Or _
                                   data_campos.Recordset("lista") = "MOVIL_TRASL" Or _
                                   data_campos.Recordset("lista") = "COD_MEDICO" Or _
                                   data_campos.Recordset("lista") = "TIPO_TRASL" Or _
                                   data_campos.Recordset("lista") = "LLAM_CANCELADO" Or _
                                   data_campos.Recordset("lista") = "COSTO_LLAM" Or _
                                   data_campos.Recordset("lista") = "BOLETA_NRO" Then
                                   data_inf.Recordset("motcon") = Trim(Str(data_llam.Recordset(Trim(data_campos.Recordset("campo")))))
                                Else
                                   data_inf.Recordset("motcon") = Mid(data_llam.Recordset(Trim(data_campos.Recordset("campo"))), 1, 100)
                                End If
                             End If
                             data_inf.Recordset("motcance") = List3.List(List3.ListIndex)
                          End If
                          If XX = 3 Then
                             If data_campos.Recordset("lista") = "SEXO" Then
                                If IsNull(data_llam.Recordset(Trim(data_campos.Recordset("campo")))) = False Then
                                   If data_llam.Recordset(Trim(data_campos.Recordset("campo"))) = 0 Then
                                      data_inf.Recordset("nomcat") = "MASCULINO"
                                   Else
                                      data_inf.Recordset("nomcat") = "FEMENINO"
                                   End If
                                Else
                                   data_inf.Recordset("nomcat") = "FEMENINO"
                                End If
                             Else
                                If data_campos.Recordset("lista") = "MATRICULA" Or _
                                   data_campos.Recordset("lista") = "EDAD" Or _
                                   data_campos.Recordset("lista") = "SEXO" Or _
                                   data_campos.Recordset("lista") = "CEDULA" Or _
                                   data_campos.Recordset("lista") = "COD_ZONA" Or _
                                   data_campos.Recordset("lista") = "BASE" Or _
                                   data_campos.Recordset("lista") = "MOVIL_LLAM" Or _
                                   data_campos.Recordset("lista") = "MOVIL_TRASL" Or _
                                   data_campos.Recordset("lista") = "COD_MEDICO" Or _
                                   data_campos.Recordset("lista") = "TIPO_TRASL" Or _
                                   data_campos.Recordset("lista") = "LLAM_CANCELADO" Or _
                                   data_campos.Recordset("lista") = "COSTO_LLAM" Or _
                                   data_campos.Recordset("lista") = "BOLETA_NRO" Then
                                   data_inf.Recordset("nomcat") = Trim(Str(data_llam.Recordset(Trim(data_campos.Recordset("campo")))))
                                Else
                                   data_inf.Recordset("nomcat") = Mid(data_llam.Recordset(Trim(data_campos.Recordset("campo"))), 1, 50)
                                End If
                             End If
                             data_inf.Recordset("lugar") = List3.List(List3.ListIndex)
                          End If
                       End If
                   Next
                   data_inf.Recordset.Update
'                   If data_llam.Recordset.EOF = True Then
'                   Else
                '      data_llam.Recordset.MoveNext
'                   End If
                Else
                   For XX = 1 To List1.ListCount
                       List1.ListIndex = XX - 1
                       data_campos.RecordSource = "Select * from campos where lista ='" & List1.List(List1.ListIndex) & "'"
                       data_campos.Refresh
                       If data_campos.Recordset.RecordCount > 0 Then
                          XCol = XX
                          If data_campos.Recordset("lista") = "FECHA_RECIBIDO" Then
                             Xarchexel3.Cells(Xlin, XCol) = Format(data_llam.Recordset(Trim(data_campos.Recordset("campo"))), "dd/mm/yyyy")
                          Else
                             If data_campos.Recordset("lista") = "SEXO" Then
                                If IsNull(data_llam.Recordset(Trim(data_campos.Recordset("campo")))) = False Then
                                   If data_llam.Recordset(Trim(data_campos.Recordset("campo"))) = 0 Then
                                      Xarchexel3.Cells(Xlin, XCol) = "MASCULINO"
                                   Else
                                      Xarchexel3.Cells(Xlin, XCol) = "FEMENINO"
                                   End If
                                Else
                                   Xarchexel3.Cells(Xlin, XCol) = "FEMENINO"
                                End If
                             Else
                                Xarchexel3.Cells(Xlin, XCol) = data_llam.Recordset(Trim(data_campos.Recordset("campo")))
                             End If
                          End If
                       End If
                   Next
                End If
            End If
            If data_llam.Recordset.EOF = True Then
            Else
               data_llam.Recordset.MoveNext
            End If
            Xtotregs = Xtotregs + 1
            Xlin = Xlin + 1
            XX = 1
         Loop
         Xlin = Xlin + 1
         XCol = 2
         If Option2.value = True Then
         Else
            Xarchexel3.Cells(Xlin, XCol) = "TOTAL DE REGISTROS: " & Xtotregs
         End If
         List1.ListIndex = 0
         List1.Enabled = True
         List2.Enabled = True
         List3.Enabled = True
         b_proc.Enabled = False
         b_salir.Enabled = False
         frm_infdesp3.MousePointer = 0
         MsgBox "Proceso terminado"
         If Option2.value = True Then
            If List3.ListCount = 1 Then
               cr1.ReportFileName = App.Path & "\infdespn1.rpt"
               cr1.ReportTitle = "Informe de " & Combo1.Text & " desde: " & md.Text & " hasta: " & mh.Text
               cr1.Action = 1
            Else
               If List3.ListCount = 2 Then
                  cr1.ReportFileName = App.Path & "\infdespn2.rpt"
                  cr1.ReportTitle = "Informe de " & Combo1.Text & " desde: " & md.Text & " hasta: " & mh.Text
                  cr1.Action = 1
               Else
                  cr1.ReportFileName = App.Path & "\infdespn3.rpt"
                  cr1.ReportTitle = "Informe de " & Combo1.Text & " desde: " & md.Text & " hasta: " & mh.Text
                  cr1.Action = 1
               End If
            End If
         Else
            Xlibexel3.Save
            Xlibexel3.Close
            Xobjexel3.Quit
'            Shell frm_menu.data_usuac.Recordset("destino") & "excel.exe " & Xarchtex, vbMaximizedFocus
            Xlabrir.Workbooks.Open Xarchtex, , False
            Xlabrir.Visible = True
            Xlabrir.WindowState = xlMaximized
         End If
      End If
   End If
Else
   MsgBox "Seleccione rango de fechas y origen del listado"

End If


List1.Enabled = True
List2.Enabled = True
List3.Enabled = True
b_proc.Enabled = True
b_salir.Enabled = True

List1.Visible = True
List3.Visible = True
Label6.Visible = False

frm_infdesp3.MousePointer = 0

Exit Sub

Xquepasainf33:
              If Err.Number = 3155 Then
                 MsgBox "Error al grabar datos, verifique"
              Else
                 MsgBox "Hay un error al generar " & Err.Number
              End If
              Xlibexel3.Close
              Xobjexel3.Quit


End Sub

Private Sub b_salir_Click()
Unload Me

End Sub

Private Sub Check1_Click()
If Check1.value = 1 Then
   If List3.ListCount > 0 Then
      t_sel.Visible = True
      t_sel.SetFocus
   Else
      MsgBox "No existen campos para seleccionar"
      t_sel.Visible = False
   End If
Else
   t_sel.Visible = False
End If

End Sub

Private Sub Form_Load()
data_inf.DatabaseName = App.Path & "\informes.mdb"

data_campos.DatabaseName = App.Path & "\campos.mdb"
data_campos.RecordSource = "Select * from campos order by id"
data_campos.Refresh

data_llam.ConnectionString = "dsn=" & Xconexrmt

List2.AddItem "FECHA_RECIBIDO"
List2.AddItem "HORA_RECIBIDO"
List2.AddItem "HORA_GRABA"
List2.AddItem "MATRICULA"
List2.AddItem "NOMBRE"
List2.AddItem "COD_CONVENIO"
List2.AddItem "DESCRIP_CNV"
List2.AddItem "EDAD"
List2.AddItem "SEXO"
List2.AddItem "CEDULA"
List2.AddItem "DIRECCION"
List2.AddItem "LOCALIDAD"
List2.AddItem "TELEFONO"
List2.AddItem "COD_ZONA"
List2.AddItem "BASE"
List2.AddItem "CLAVE_INICIAL"
List2.AddItem "CLAVE FINAL"
List2.AddItem "MOTIVO_CONS"
List2.AddItem "DIAGNOSTICO"
List2.AddItem "MOVIL_LLAM"
List2.AddItem "MOVIL_TRASL"
List2.AddItem "COD_MEDICO"
List2.AddItem "NOM_MEDICO"
List2.AddItem "HORA_PASADO"
List2.AddItem "HORA_SALIDA"
List2.AddItem "HORA_LLEGADA"
List2.AddItem "HORA_REALIZA"
List2.AddItem "HORA_TD"
List2.AddItem "TIPO_TRASL"
List2.AddItem "HORA_SOLTRASL"
List2.AddItem "LUGAR_TRASL"
List2.AddItem "HORA_SALE_TR"
List2.AddItem "HORA_LLEG_TR"
List2.AddItem "HORA_SALE_CA"
List2.AddItem "HORA_ENZONA"
List2.AddItem "USUARIO_REC"
List2.AddItem "USUARIO_DESP"
List2.AddItem "LLAM_CANCELADO"
List2.AddItem "OBS_CANCELADO"
List2.AddItem "HORA_CANCELA"
List2.AddItem "COSTO_LLAM"
List2.AddItem "BOLETA_NRO"
List2.AddItem "TIPO_BOLETA"
List2.AddItem "OBSERVACIONES"




End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub List1_DblClick()
If List1.ListIndex >= 0 Then
   List1.RemoveItem List1.ListIndex
End If

End Sub

Private Sub List3_DblClick()
If List3.ListIndex >= 0 Then
   List3.RemoveItem List3.ListIndex
End If

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub

