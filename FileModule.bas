Attribute VB_Name = "FileModule"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Public Function DownloadFile(url, path)

   Dim objReq
   Dim objStream
    On Error GoTo errorDownloadFile
    Set objReq = CreateObject("Msxml2.ServerXMLHTTP")
    objReq.Open "GET", url, False, "", ""
    objReq.setRequestHeader "Cache-Control", "max-age=0"
    objReq.send

   If objReq.Status = 200 Then
       Set objStream = CreateObject("ADODB.Stream")
       objStream.Open
       objStream.Type = 1
        
       objStream.Write objReq.responseBody
       objStream.Position = 0

       objStream.SaveToFile path, 2
       objStream.Close
       Set objStream = Nothing
   End If

   Set objReq = Nothing
   
exitSub:
    Exit Function

errorDownloadFile:
    MsgBox "Ha ocurrido un error al obtener el archivo, por favor valide si no tiene el mismo archivo abierto, de lo contrario comuniquese con cómputos (" & Err.Description & ")"




End Function


Function LoadUserFile(ByVal File As String)
On Error GoTo errorLoadingFile
    ShellExecute 0, vbNullString, File, vbNullString, vbNullString, vbNormalFocus
    Exit Function

errorLoadingFile:
    MsgBox "Ha ocurrido un error al obtener el archivo, por favor valide si no tiene el mismo archivo abierto, de lo contrario comuniquese con cómputos (" & Err.Description & ")"

    
End Function
