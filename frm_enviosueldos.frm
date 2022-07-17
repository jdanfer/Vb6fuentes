VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_enviosueldos 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Envío de correos de recibos de sueldo"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8955
   Icon            =   "frm_enviosueldos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   8955
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton b_inf 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   8040
      Picture         =   "frm_enviosueldos.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Informe de envíos"
      Top             =   2880
      Width           =   495
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   3720
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4320
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox t_asunto 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   6360
      Picture         =   "frm_enviosueldos.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Enviar los correos con los archivos de la lista"
      Top             =   2880
      Width           =   495
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_enviosueldos.frx":109E
      Height          =   2775
      Left            =   120
      OleObjectBlob   =   "frm_enviosueldos.frx":10B2
      TabIndex        =   3
      Top             =   360
      Width           =   6015
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   6360
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "Asunto del correo:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Archivos Disponibles para enviar"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Personal disponible para envíos"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frm_enviosueldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_inf_Click()
Dim Xdesde, Xhasta As String

Xdesde = InputBox("Ingrese Desde que fecha")
Xhasta = InputBox("Ingrese Hasta que fecha")
If Xdesde <> "" And Xhasta <> "" Then
   If data_inf.Recordset.RecordCount > 0 Then
      data_inf.Recordset.MoveFirst
      Do While Not data_inf.Recordset.EOF
         data_inf.Recordset.Delete
         data_inf.Recordset.MoveNext
      Loop
   End If
   Data2.RecordSource = "Select * from envioshist where fecha >=#" & Format(Xdesde, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xhasta, "yyyy/mm/dd") & "#"
   Data2.Refresh
   If Data2.Recordset.RecordCount > 0 Then
      Data2.Recordset.MoveFirst
      Do While Not Data2.Recordset.EOF
         data_inf.Recordset.AddNew
         data_inf.Recordset("fecha") = Data2.Recordset("fecha")
         data_inf.Recordset("nro") = Data2.Recordset("nro")
         Data1.Recordset.FindFirst "nro =" & Data2.Recordset("nro")
         If Not Data1.Recordset.NoMatch Then
            data_inf.Recordset("nombre") = Data1.Recordset("nombre")
         End If
         data_inf.Recordset("hora") = Data2.Recordset("hora")
         data_inf.Recordset("obs") = Data2.Recordset("obs")
         data_inf.Recordset.Update
         Data2.Recordset.MoveNext
      Loop
      MsgBox "Proceso terminado"
      data_inf.RecordSource = "Select * from envioshist order by nro"
      data_inf.Refresh
      cr1.ReportFileName = App.Path & "\infrecsue.rpt"
      cr1.ReportTitle = "Informe de envíos de correos con recibo de sueldo desde: " & Xdesde & " hasta: " & Xhasta
      cr1.Action = 1
   End If
End If


End Sub

Private Sub Command1_Click()
Dim MenCorreo As String
Dim oMail As Class1
Dim Xcanenvio As Long
Dim XX As Integer

Xcanenvio = 0
Data2.RecordSource = "envioshist"
Data2.Refresh

If List1.ListCount >= 1 Then
   Set oMail = New Class1
   For XX = 1 To List1.ListCount
       List1.ListIndex = XX - 1
       Data1.Recordset.FindFirst "nro =" & Val(List1.List(List1.ListIndex))
       If Not Data1.Recordset.NoMatch Then
          If IsNull(Data1.Recordset("correo")) = False Then
             With oMail
'                 .servidor = "adinet.com.uy"
                 .servidor = "smtp.office365.com"
'                 .servidor = "vera.com.uy"
'                 .servidor = "outlook.office365.com"
                 .puerto = 25
                 .UseAuntentificacion = True
                 .ssl = True
                 .Usuario = "gestionhumana@sapp.com.uy"
                 .PassWord = "$.Sapp1987"
                 .Asunto = t_asunto.Text
                 .de = "gestionhumana@sapp.com.uy"
                 .para = Data1.Recordset("correo")
                 .Adjunto = "c:\recibos\" & List1.List(List1.ListIndex)
                 .Mensaje = "Recibo de Sueldo"
                 .Enviar_Backup ' manda el mail
             End With
          '  MsgBox "Correo enviado..."
             Xcanenvio = Xcanenvio + 1
             Label4.Caption = "CORREOS ENVIADOS: " & Xcanenvio
             Data1.Recordset.Edit
             Data1.Recordset("fecha") = Date
             Data1.Recordset("hora") = Format(Time, "HH:mm")
             Data1.Recordset.Update
             Data2.Recordset.AddNew
             Data2.Recordset("nro") = Data1.Recordset("nro")
             Data2.Recordset("fecha") = Date
             Data2.Recordset("hora") = Format(Time, "HH:mm")
             Data2.Recordset("obs") = "ENVIO OK " & Mid(t_asunto.Text, 1, 39)
             Data2.Recordset.Update
          Else
             Data2.Recordset.AddNew
             Data2.Recordset("nro") = Data1.Recordset("nro")
             Data2.Recordset("fecha") = Date
             Data2.Recordset("hora") = Format(Time, "HH:mm")
             Data2.Recordset("obs") = "NO ENVIADO - NO FIGURA CORREO"
             Data2.Recordset.Update
          End If
       Else
          Data2.Recordset.AddNew
          Data2.Recordset("nro") = Data1.Recordset("nro")
          Data2.Recordset("fecha") = Date
          Data2.Recordset("hora") = Format(Time, "HH:mm")
          Data2.Recordset("obs") = "NO ENVIADO - NO FUNCIONARIO"
          Data2.Recordset.Update
       End If
   Next
   Set oMail = Nothing

Else
   MsgBox "No hay archivos para enviar"
End If
XX = 1

If List1.ListCount >= 1 Then
   For XX = 1 To List1.ListCount
       List1.ListIndex = XX - 1
       Kill ("c:\recibos\" & List1.List(List1.ListIndex))
   Next
   List1.Clear
End If


End Sub

Private Sub DBGrid1_DblClick()
frm_hist.Show vbModal

End Sub

Private Sub Form_Load()
Dim Elarchivo As String
Elarchivo = Dir("C:\recibos\*.pdf")

If App.PrevInstance = True Then
   MsgBox "Ya está abierto el programa", vbCritical
   End
End If

Do While Elarchivo <> ""
   List1.AddItem Elarchivo
   Elarchivo = Dir
Loop

Data1.DatabaseName = App.Path & "\enviosrs.mdb"
Data1.RecordSource = "enviosrs"
Data1.Refresh

data_inf.DatabaseName = App.Path & "\infrecsue.mdb"
data_inf.RecordSource = "envioshist"
data_inf.Refresh

Data2.DatabaseName = App.Path & "\enviosrs.mdb"

t_asunto.Text = "Recibo de Sueldo mes "

End Sub
