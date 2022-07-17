VERSION 5.00
Begin VB.Form frm_vertrasl 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton b_cerrar 
      Height          =   495
      Left            =   7080
      Picture         =   "frm_vertrasl.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cerrar"
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frm_vertrasl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_cerrar_Click()
Unload Me

End Sub

Private Sub Form_Load()
Label1.Caption = "OPCION: TRASLADOS A TERCEROS" & vbCrLf
Label1.Caption = Label1.Caption & "SELECCIONAR ÉSTA OPCIÓN CUANDO ES:" & vbCrLf
Label1.Caption = Label1.Caption & "--SOLICITUD DE MSP" & vbCrLf
Label1.Caption = Label1.Caption & "--SOLICITUD MEDICA URUGUAYA SUÁREZ,PANDO,MIGUES O ATLÁNTIDA" & vbCrLf
Label1.Caption = Label1.Caption & "--SOLICITUD CAAMEPA" & vbCrLf
Label1.Caption = Label1.Caption & "--SOLICITUD CAAMEPA MIGUES, MONTES" & vbCrLf
Label1.Caption = Label1.Caption & "--SOLICITUD CASMU SALINAS, ATLÁNTIDA" & vbCrLf
Label1.Caption = Label1.Caption & "--SOLICITUD CÍRCULO CATÓLICO ATLÁNTIDA" & vbCrLf
Label1.Caption = Label1.Caption & "--SOLICITUD CÍRCULO CATÓLICO-MSP" & vbCrLf
Label1.Caption = Label1.Caption & "--SOLICITUD RETORNOS DE CÍRCULO CATÓLICO" & vbCrLf
Label1.Caption = Label1.Caption & "--SOLICITADOS POR CÍRCULO CATÓLICO" & vbCrLf
Label1.Caption = Label1.Caption & "--SOLICITUD DE OTROS RETORNOS" & vbCrLf
Label1.Caption = Label1.Caption & "--SOLICITUD DE AS.ESPAÑOLA/OTROS RETORNOS" & vbCrLf


End Sub
