Attribute VB_Name = "NetSyncUtils"
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Sub SetConfigVar Lib "c:\Archivos de programa\Auditores Asociados\NetSync\netsyncodbc.dll" (ByVal varName As String, ByVal varValue As String)


Public Sub SelectLimit(ByVal count As Integer)
 Dim lb As Long, pa As Long
    'map 'user32' into the address space of the calling process.
    lb = LoadLibrary(Environ("programFiles") & "\Auditores Asociados\NetSync\netsyncodbc.dll")
    Call SetConfigVar("2000", count)
    'unmap the library's address
    FreeLibrary lb
End Sub
