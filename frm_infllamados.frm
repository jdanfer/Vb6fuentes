VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Width           =   1980
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Width           =   2580
   End
   Begin VB.Data data_llama 
      Caption         =   "data_llama"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Width           =   2460
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xcuenta1, Xcuenta2, Xcuenta3, Xcuenta4, Xcuenta5, Xcuenta6, Xcuenta7  As Long
Dim Xcuenta8, Xcuenta9, Xcuenta10, Xcuenta11, Xcuenta12, Xcuenta13, Xcuenta14  As Long
Dim Xhd, Xhh As String
Xcuenta1 = 0
Xcuenta2 = 0
Xcuenta3 = 0
Xcuenta4 = 0
Xcuenta5 = 0
Xcuenta6 = 0
Xcuenta7 = 0
Xcuenta8 = 0
Xcuenta9 = 0
Xcuenta10 = 0
Xcuenta11 = 0
Xcuenta12 = 0
Xcuenta13 = 0

Xhd = "23:00"
Xhh = "23:59"
data_llama.RecordSource = "Select * from inflla where hor_rea >='" & Xhd & "' And hor_rea <='" & Xhh & "' And trasla =" & 1 & " order by pend"
data_llama.Refresh
If data_llama.Recordset.RecordCount > 0 Then
    data_llama.Recordset.MoveFirst
    Do While Not data_llama.Recordset.EOF
       If data_llama.Recordset("pend") = 1 Or data_llama.Recordset("pend") = 4 Then
          Xcuenta1 = Xcuenta1 + 1
       End If
       If data_llama.Recordset("pend") = 3 Then
          Xcuenta2 = Xcuenta2 + 1
       End If
       If data_llama.Recordset("pend") = 2 Or data_llama.Recordset("pend") = 7 Or data_llama.Recordset("pend") = 5 Then
          Xcuenta3 = Xcuenta3 + 1
       End If
       If data_llama.Recordset("pend") = 6 Then
          Xcuenta4 = Xcuenta4 + 1
       End If
       If data_llama.Recordset("pend") = 11 Then
          Xcuenta5 = Xcuenta5 + 1
       End If
       If data_llama.Recordset("pend") = 8 Then
          Xcuenta6 = Xcuenta6 + 1
       End If
       If data_llama.Recordset("pend") = 9 Then
          Xcuenta7 = Xcuenta7 + 1
       End If
       If data_llama.Recordset("pend") = 10 Then
          Xcuenta8 = Xcuenta8 + 1
       End If
       If data_llama.Recordset("pend") = 15 Then
          Xcuenta9 = Xcuenta9 + 1
       End If
       If data_llama.Recordset("pend") = 13 Then
          Xcuenta10 = Xcuenta10 + 1
       End If
       If data_llama.Recordset("pend") = 14 Then
          Xcuenta11 = Xcuenta11 + 1
       End If
       If data_llama.Recordset("pend") = 99 Or data_llama.Recordset("pend") = 0 Then
          If data_llama.Recordset("codzon") = 1 Then
             Xcuenta12 = Xcuenta12 + 1
          Else
             Xcuenta13 = Xcuenta13 + 1
          End If
       End If
       data_llama.Recordset.MoveNext
    Loop
End If
data_cli.Recordset.MoveFirst

data_cli.Recordset.Edit
data_cli.Recordset("hora23a24") = Xcuenta1
data_cli.Recordset.Update

data_cli.Recordset.MoveNext
data_cli.Recordset.Edit
data_cli.Recordset("hora23a24") = Xcuenta2
data_cli.Recordset.Update

data_cli.Recordset.MoveNext
data_cli.Recordset.Edit
data_cli.Recordset("hora23a24") = Xcuenta3
data_cli.Recordset.Update

data_cli.Recordset.MoveNext
data_cli.Recordset.Edit
data_cli.Recordset("hora23a24") = Xcuenta4
data_cli.Recordset.Update

data_cli.Recordset.MoveNext
data_cli.Recordset.Edit
data_cli.Recordset("hora23a24") = Xcuenta5
data_cli.Recordset.Update

data_cli.Recordset.MoveNext
data_cli.Recordset.Edit
data_cli.Recordset("hora23a24") = Xcuenta6
data_cli.Recordset.Update

data_cli.Recordset.MoveNext
data_cli.Recordset.Edit
data_cli.Recordset("hora23a24") = Xcuenta7
data_cli.Recordset.Update

data_cli.Recordset.MoveNext
data_cli.Recordset.Edit
data_cli.Recordset("hora23a24") = Xcuenta8
data_cli.Recordset.Update

data_cli.Recordset.MoveNext
data_cli.Recordset.Edit
data_cli.Recordset("hora23a24") = Xcuenta9
data_cli.Recordset.Update

data_cli.Recordset.MoveNext
data_cli.Recordset.Edit
data_cli.Recordset("hora23a24") = Xcuenta10
data_cli.Recordset.Update

data_cli.Recordset.MoveNext
data_cli.Recordset.Edit
data_cli.Recordset("hora23a24") = Xcuenta11
data_cli.Recordset.Update

data_cli.Recordset.MoveNext
data_cli.Recordset.Edit
data_cli.Recordset("hora23a24") = Xcuenta12
data_cli.Recordset.Update

data_cli.Recordset.MoveNext
data_cli.Recordset.Edit
data_cli.Recordset("hora23a24") = Xcuenta13
data_cli.Recordset.Update

MsgBox "Terminado"

End Sub

Private Sub Command2_Click()
Dim Xvalo As Integer
Xvalo = 1
data_cli.RecordSource = "Select * from porzonas"
data_cli.Refresh
Data1.DatabaseName = App.Path & "\horaslla.mdb"
Data1.RecordSource = "zonashs"
Data1.Refresh
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
   data_cli.Recordset.MoveFirst
   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora0a1")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora1a2")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora2a3")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora3a4")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora4a5")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora5a6")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora6a7")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora7a8")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora8a9")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora9a10")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora10a11")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora11a12")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora12a13")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora13a14")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora14a15")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora15a16")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora16a17")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora17a18")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora18a19")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora19a20")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora20a21")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora21a22")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora22a23")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.Edit
   data_cli.Recordset("zona" & Trim(Str(Xvalo))) = Data1.Recordset("hora23a24")
   data_cli.Recordset.Update
   data_cli.Recordset.MoveNext

   data_cli.Recordset.MoveFirst
   Xvalo = Xvalo + 1
   Data1.Recordset.MoveNext

Loop

MsgBox "Terminado"

End Sub

Private Sub Form_Load()
data_llama.DatabaseName = App.Path & "\informes.mdb"
data_llama.RecordSource = "inflla"
data_llama.Refresh
data_cli.DatabaseName = App.Path & "\horaslla.mdb"
data_cli.RecordSource = "zonashs"
data_cli.Refresh

End Sub
