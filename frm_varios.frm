VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Data data_mov 
      Caption         =   "data_mov"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Disponibilidad de Móviles"
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data2.DatabaseName = App.Path & "\sapp.mdb"

Data1.DatabaseName = App.Path & "\sapp.mdb"
'Data1.RecordSource = "Select * from llamado where codzon in(1,2) and fecha >=#" & Format("01/07/2013", "yyyy/mm/dd") & "# and fecha <=#" & Format("01/07/2013", "yyyy/mm/dd") & "# and hora >='" & "09:00" & "' and hora <='" & "21:00" & "' order by hora"
Data1.RecordSource = "Select * from llamado where codzon in(1,2) and fecha >=#" & Format("01/07/2013", "yyyy/mm/dd") & "# and fecha <=#" & Format("31/07/2013", "yyyy/mm/dd") & "# order by fecha, hora"
Data1.Refresh
data_mov.DatabaseName = App.Path & "\tmoviles.mdb"
data_mov.RecordSource = "Select * from tmovil"
data_mov.Refresh
If data_mov.Recordset.RecordCount > 0 Then
   data_mov.Recordset.MoveFirst
   Do While Not data_mov.Recordset.EOF
      data_mov.Recordset.Edit
      data_mov.Recordset("hora") = "00:00"
      data_mov.Recordset.Update
      data_mov.Recordset.MoveNext
   Loop
   data_mov.Recordset.MoveFirst
End If

Dim Xcant, Xdia, Xcantf As Long

Xdia = 1
Xcant = 0
Xcantf = 0
Form1.MousePointer = 11
Do While Not Data1.Recordset.EOF
    Do While Day(Data1.Recordset("fecha")) = Xdia
       data_mov.RecordSource = "Select * from tmovil where hora <'" & Data1.Recordset("hora") & "'"
       data_mov.Refresh
       If data_mov.Recordset.RecordCount <= 0 Then
          Xcant = Xcant + 1
          MsgBox "ES:" & Data1.Recordset("fecha") & "MAT:" & Data1.Recordset("matric")
        
       End If
       data_mov.RecordSource = "Select * from tmovil where movil =" & Data1.Recordset("movilpas")
       data_mov.Refresh
       If data_mov.Recordset.RecordCount > 0 Then
          data_mov.Recordset.Edit
          If IsNull(Data1.Recordset("hzona")) = False Then
             If Data1.Recordset("hzona") <> "" Then
                data_mov.Recordset("hora") = Data1.Recordset("hzona")
             Else
                data_mov.Recordset("hora") = Data1.Recordset("hor_rea")
             End If
          Else
             data_mov.Recordset("hora") = Data1.Recordset("hor_rea")
          End If
          data_mov.Recordset.Update
       End If
       Data1.Recordset.MoveNext
    Loop
    If Xcant >= 1 Then
       Xcantf = Xcantf + 1
       MsgBox "HAY:" & Xcantf
    End If
    Xdia = Xdia + 1
    Xcant = 0
    data_mov.RecordSource = "Select * from tmovil"
    data_mov.Refresh
    If data_mov.Recordset.RecordCount > 0 Then
       data_mov.Recordset.MoveFirst
       Do While Not data_mov.Recordset.EOF
          data_mov.Recordset.Edit
          data_mov.Recordset("hora") = "00:00"
          data_mov.Recordset.Update
          data_mov.Recordset.MoveNext
       Loop
       data_mov.Recordset.MoveFirst
    End If
Loop
Form1.MousePointer = 0

MsgBox "Proceso terminado"

End Sub

