VERSION 5.00
Begin VB.Form fSplashActualizador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descargando Actualizacion: "
   ClientHeight    =   930
   ClientLeft      =   -15
   ClientTop       =   360
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ucAsyncDLHost 
      Align           =   4  'Align Right
      Height          =   930
      Left            =   15
      ScaleHeight     =   870
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   0
      Width           =   7245
   End
End
Attribute VB_Name = "fSplashActualizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private url As String
Private nombreArchivo As String

Public Property Let SetUrl(ByVal newValue As String)
    url = newValue
End Property
Public Property Let SetName(ByVal newValue As String)
    nombreArchivo = newValue
End Property

Private Sub Form_Load()
  ucAsyncDLHost.AddDownloadJob url, App.path & "/" & nombreArchivo & ".exe"
  Me.Caption = Me.Caption & " " & nombreArchivo
End Sub


Private Sub ucAsyncDLHost_DownloadComplete(Sender As ucAsyncDLStripe, ByVal TmpFileName As String)
  Debug.Print "DownloadComplete for URL: "; Sender.url
  On Error Resume Next
  Kill App.path & "/" & nombreArchivo & ".exe"
  Name TmpFileName As Sender.LocalFileName
  On Error GoTo 0
End Sub
 
Private Sub ucAsyncDLHost_DownloadProgress(Sender As ucAsyncDLStripe, ByVal BytesRead As Long, ByVal BytesTotal As Long)
  Sender.Caption = FormatBytes2KBMBGBTB(BytesRead) & " (" & FormatDLRate(BytesRead, DateDiff("s", Sender.StartDate, Now)) & ")"
End Sub

Private Sub ucAsyncDLHost_DownloadClose()
    Unload Me
End Sub

Function FormatBytes2KBMBGBTB(ByVal Bytes As Currency) As String
Dim i As Long
  Do While Bytes >= 1024: Bytes = Bytes / 1024: i = i + 1: Loop
  FormatBytes2KBMBGBTB = Int(Bytes * 10) / 10 & Split(",K,M,G,T", ",")(i) & "B"
End Function

Function FormatDLRate(ByVal Bytes As Long, ByVal Seconds As Long) As String
  If Seconds Then FormatDLRate = FormatBytes2KBMBGBTB(Bytes \ Seconds) & "/s"
End Function

Private Sub ucAsyncDLHost_Click()

End Sub
