VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3390
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Terminar"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Nomarq As String
Nomarq = "arq0509"
createTablearqueo (Nomarq)

MsgBox "Terminado"

End Sub

Private Function createTablearqueo(tablename As String) As Boolean
    Dim conODBCDirect As DAO.Connection
    Dim rsODBCDirect As DAO.Recordset
    Dim strConn As String
    strConn = "ODBC;DSN=SAPP;"
    Set wrkODBC = CreateWorkspace("", "admin", "", dbUseODBC)
    Set conODBCDirect = wrkODBC.OpenConnection("", , , strConn)
    On Error Resume Next
'''    tablename = "ARQ0108"
    conODBCDirect.Execute ("call prcreatearq('" & tablename & "')")
    If Err <> 0 Then
        createTablearqueo = False
    Else
        createTablearqueo = True
    End If

End Function

Private Sub Command2_Click()
End

End Sub

