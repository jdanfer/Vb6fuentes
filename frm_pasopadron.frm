VERSION 5.00
Begin VB.Form frm_pasopadron 
   Caption         =   "Paso a padron"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   5670
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Procesar SEMM"
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Arreglar CI..."
      Height          =   735
      Left            =   840
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar Ret.Mil."
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Data data_parsec 
      Caption         =   "data_parsec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Data data_clientes 
      Caption         =   "data_clientes"
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
      Top             =   600
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.Data data_nuevo 
      Caption         =   "data_nuevo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2940
   End
End
Attribute VB_Name = "frm_pasopadron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xmat As Long
Dim Xfec As Variant
Xmat = data_parsec.Recordset("ultimo_soc") + 1
' Base de datos Retirados Militares
frm_pasopadron.MousePointer = 11
Command1.Enabled = False
If data_nuevo.Recordset.RecordCount > 0 Then
   data_nuevo.Recordset.MoveFirst
   Do While Not data_nuevo.Recordset.EOF
      data_clientes.Recordset.FindFirst "cl_cedula =" & data_nuevo.Recordset("ced")
      If Not data_clientes.Recordset.NoMatch Then
         data_clientes.Recordset.AddNew
         data_clientes.Recordset("estado") = 1
        '   labestado.Caption = "ACTIVO"
         data_clientes.Recordset("cl_codigo") = Xmat
         data_clientes.Recordset("cl_codconv") = "RETMI"
         data_clientes.Recordset("cl_nomconv") = "RETIRADOS MILITARES"
         data_clientes.Recordset("cl_apellid") = data_nuevo.Recordset("apel") + " " + data_nuevo.Recordset("nom")
         data_clientes.Recordset("cl_cedula") = data_nuevo.Recordset("ced")
         data_clientes.Recordset("cl_codced") = data_nuevo.Recordset("co")
         data_clientes.Recordset("cl_edad") = 0
         data_clientes.Recordset("cl_uniedad") = "A"
         data_clientes.Recordset("cl_ultmesp") = 0
         data_clientes.Recordset("cl_ultanop") = 0
         data_clientes.Recordset("cl_atrasoa") = 0
         If IsNull(data_nuevo.Recordset("dir")) = False Then
            data_clientes.Recordset("cl_direcci") = data_nuevo.Recordset("dir")
         Else
            
         End If
         data_clientes.Recordset("cl_grupo") = 999
         data_clientes.Recordset("cl_zona") = "*TODOS"
         data_clientes.Recordset("cl_sexo") = 1
         data_clientes.Recordset("cl_telefon") = data_nuevo.Recordset("tel")
         If IsNull(data_nuevo.Recordset("mat")) = False Then
            data_clientes.Recordset("cl_nrosocm") = Trim(Str(data_nuevo.Recordset("mat")))
         End If
         data_clientes.Recordset("cl_fecing") = Format(Date, "dd/mm/yyyy")
         data_clientes.Recordset("cl_nrovend") = 102
         data_clientes.Recordset("cl_nomvend") = "RET.MILITARES"
         data_clientes.Recordset("cl_nrocobr") = 102
         data_clientes.Recordset("cl_nomcobr") = "RET.MILITARES"
         data_clientes.Recordset("cl_forpago") = 1
         data_clientes.Recordset("cl_descpag") = "Abono Mensual"
         data_clientes.Recordset("fecha_sys") = Format(Date, "dd/mm/yyyy")
         data_clientes.Recordset.Update
         data_parsec.Recordset.Edit
         data_parsec.Recordset("ultimo_soc") = Xmat
         data_parsec.Recordset.Update
         Xmat = Xmat + 1
      Else
         If IsNull(data_clientes.Recordset("estado")) = False Then
            If data_clientes.Recordset("estado") = 2 Or data_clientes.Recordset("estado") = 3 Then
               data_clientes.Recordset.Edit
               data_clientes.Recordset("estado") = 1
               '   labestado.Caption = "ACTIVO"
               data_clientes.Recordset("cl_codconv") = "RETMI"
               data_clientes.Recordset("cl_nomconv") = "RETIRADOS MILITARES"
               data_clientes.Recordset("cl_apellid") = data_nuevo.Recordset("apel") + " " + data_nuevo.Recordset("nom")
               data_clientes.Recordset("cl_cedula") = data_nuevo.Recordset("ced")
               data_clientes.Recordset("cl_codced") = data_nuevo.Recordset("co")
               data_clientes.Recordset("cl_ultmesp") = 0
               data_clientes.Recordset("cl_ultanop") = 0
               data_clientes.Recordset("cl_atrasoa") = 0
               If IsNull(data_nuevo.Recordset("dir")) = False Then
                  data_clientes.Recordset("cl_direcci") = data_nuevo.Recordset("dir")
               Else
               End If
               data_clientes.Recordset("cl_grupo") = 999
               data_clientes.Recordset("cl_zona") = "*TODOS"
               data_clientes.Recordset("cl_telefon") = data_nuevo.Recordset("tel")
               If IsNull(data_nuevo.Recordset("mat")) = False Then
                  data_clientes.Recordset("cl_nrosocm") = Trim(Str(data_nuevo.Recordset("mat")))
               End If
               data_clientes.Recordset("cl_fecing") = Format(Date, "dd/mm/yyyy")
               data_clientes.Recordset("cl_nrovend") = 102
               data_clientes.Recordset("cl_nomvend") = "RET.MILITARES"
               data_clientes.Recordset("cl_nrocobr") = 102
               data_clientes.Recordset("cl_nomcobr") = "RET.MILITARES"
               data_clientes.Recordset("cl_forpago") = 1
               data_clientes.Recordset("cl_descpag") = "Abono Mensual"
               data_clientes.Recordset("fecha_modi") = Format(Date, "dd/mm/yyyy")
               data_clientes.Recordset("fecha_baja") = Xfec
               data_clientes.Recordset.Update
            End If
         End If
      End If
      data_nuevo.Recordset.MoveNext
   Loop
End If
frm_pasopadron.MousePointer = 0
MsgBox "Terminado"


End Sub

Private Sub Command2_Click()
End

End Sub

Private Sub Command3_Click()
Dim Xlargo As Long
Dim Xcuento As Long
Dim Xcaract As String
Dim Xcedstr As String
data_nuevo.RecordSource = "semm"
data_nuevo.Refresh
frm_pasopadron.MousePointer = 11
If data_nuevo.Recordset.RecordCount > 0 Then
   data_nuevo.Recordset.MoveFirst
   Do While Not data_nuevo.Recordset.EOF
      Xlargo = Len(data_nuevo.Recordset("ced"))
      For Xcuento = 1 To Xlargo - 1
          If IsNumeric(Mid(data_nuevo.Recordset("ced"), Xcuento, 1)) = True Then
             Xcedstr = Xcedstr + Mid(data_nuevo.Recordset("ced"), Xcuento, 1)
          End If
      Next
      Xcaract = Mid(data_nuevo.Recordset("ced"), Xlargo, 1)
      Xcuento = 0
      data_nuevo.Recordset.Edit
      data_nuevo.Recordset("cednum") = Val(Xcedstr)
      data_nuevo.Recordset("codced") = Val(Xcaract)
      data_nuevo.Recordset.Update
      data_nuevo.Recordset.MoveNext
      Xcedstr = ""
      
   Loop
End If
frm_pasopadron.MousePointer = 0
MsgBox "Proceso terminado..."

End Sub

Private Sub Command4_Click()
Dim Xmat2 As Long
Dim Xfec2 As Variant
Xmat2 = data_parsec.Recordset("ultimo_soc") + 1
' Base de datos Retirados Militares
data_nuevo.RecordSource = "semm"
data_nuevo.Refresh
Command4.Enabled = False
frm_pasopadron.MousePointer = 11
If data_nuevo.Recordset.RecordCount > 0 Then
   data_nuevo.Recordset.MoveFirst
   Do While Not data_nuevo.Recordset.EOF
      data_clientes.Recordset.FindFirst "cl_cedula =" & data_nuevo.Recordset("cednum")
      If Not data_clientes.Recordset.NoMatch Then
         data_clientes.Recordset.AddNew
         data_clientes.Recordset("estado") = 1
        '   labestado.Caption = "ACTIVO"
         data_clientes.Recordset("cl_codigo") = Xmat2
         data_clientes.Recordset("cl_codconv") = "SEMM1"
         data_clientes.Recordset("cl_nomconv") = "SOCIOS SEMM"
         data_clientes.Recordset("cl_apellid") = data_nuevo.Recordset("apel") + " " + data_nuevo.Recordset("nom")
         data_clientes.Recordset("cl_cedula") = data_nuevo.Recordset("cednum")
         data_clientes.Recordset("cl_codced") = data_nuevo.Recordset("codced")
         data_clientes.Recordset("cl_edad") = 0
         data_clientes.Recordset("cl_uniedad") = "A"
         data_clientes.Recordset("cl_ultmesp") = 0
         data_clientes.Recordset("cl_ultanop") = 0
         data_clientes.Recordset("cl_atrasoa") = 0
         data_clientes.Recordset("cl_grupo") = 999
         data_clientes.Recordset("cl_zona") = "*TODOS"
         data_clientes.Recordset("cl_sexo") = 1
         data_clientes.Recordset("cl_fecing") = Format(Date, "dd/mm/yyyy")
         data_clientes.Recordset("cl_nrovend") = 101
         data_clientes.Recordset("cl_nomvend") = "SEMM"
         data_clientes.Recordset("cl_nrocobr") = 101
         data_clientes.Recordset("cl_nomcobr") = "SEMM"
         data_clientes.Recordset("cl_forpago") = 1
         data_clientes.Recordset("cl_descpag") = "Abono Mensual"
         data_clientes.Recordset("fecha_sys") = Format(Date, "dd/mm/yyyy")
         data_clientes.Recordset.Update
         data_parsec.Recordset.Edit
         data_parsec.Recordset("ultimo_soc") = Xmat2
         data_parsec.Recordset.Update
         Xmat2 = Xmat2 + 1
      Else
         If IsNull(data_clientes.Recordset("estado")) = False Then
            If data_clientes.Recordset("estado") = 2 Or data_clientes.Recordset("estado") = 3 Then
               data_clientes.Recordset.Edit
               data_clientes.Recordset("estado") = 1
               '   labestado.Caption = "ACTIVO"
               data_clientes.Recordset("cl_codconv") = "SEMM1"
               data_clientes.Recordset("cl_nomconv") = "SOCIOS SEMM"
               data_clientes.Recordset("cl_apellid") = data_nuevo.Recordset("apel") + " " + data_nuevo.Recordset("nom")
               data_clientes.Recordset("cl_cedula") = data_nuevo.Recordset("cednum")
               data_clientes.Recordset("cl_codced") = data_nuevo.Recordset("codced")
               data_clientes.Recordset("cl_ultmesp") = 0
               data_clientes.Recordset("cl_ultanop") = 0
               data_clientes.Recordset("cl_atrasoa") = 0
               data_clientes.Recordset("cl_grupo") = 999
               data_clientes.Recordset("cl_zona") = "*TODOS"
               data_clientes.Recordset("cl_fecing") = Format(Date, "dd/mm/yyyy")
               data_clientes.Recordset("cl_nrovend") = 101
               data_clientes.Recordset("cl_nomvend") = "SEMM"
               data_clientes.Recordset("cl_nrocobr") = 101
               data_clientes.Recordset("cl_nomcobr") = "SEMM"
               data_clientes.Recordset("cl_forpago") = 1
               data_clientes.Recordset("cl_descpag") = "Abono Mensual"
               data_clientes.Recordset("fecha_modi") = Format(Date, "dd/mm/yyyy")
               data_clientes.Recordset("fecha_baja") = Xfec2
               data_clientes.Recordset.Update
            End If
         End If
      End If
      data_nuevo.Recordset.MoveNext
   Loop
End If
frm_pasopadron.MousePointer = 0
MsgBox "Terminado"

End Sub

Private Sub Form_Load()
data_clientes.DatabaseName = App.Path & "\sapp.mdb"
data_clientes.RecordSource = "clientes"
data_clientes.Refresh
data_parsec.DatabaseName = App.Path & "\sapp.mdb"
data_parsec.RecordSource = "parsec0"
data_parsec.Refresh
data_nuevo.DatabaseName = App.Path & "\socnuev.mdb"
data_nuevo.RecordSource = "retmil"
data_nuevo.Refresh
'' semm

End Sub
