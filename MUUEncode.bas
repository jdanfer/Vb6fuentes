Attribute VB_Name = "MUUEncode"
Option Explicit

'Función para codificar archivos al formato MIME (Para adjuntar en un mail)
'Mejorada por Federico Colombo para reducir tiempo.
'En vez de hacer todo en memoria, lo graba a un archivo y luego lo lee

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public UUfiles(0 To 9) As String
Public indexUUfiles As Byte

Public Function TempDirectory() As String
Dim TempPath As String
Dim temp
TempPath = String(145, Chr(0))
temp = GetTempPath(145, TempPath)
TempDirectory = Left(TempPath, InStr(TempPath, Chr(0)) - 1)
End Function

Public Function UUEncodeFile(strFilePath As String) As String
    Dim intFile         As Integer      'file handler
    Dim intTempFile     As Integer      'temp file
    Dim TempFileName    As String
    Dim lFileSize       As Long         'size of the file
    Dim strFileName     As String       'name of the file
    Dim strFileData     As String       'file data chunk
    Dim lEncodedLines   As Long         'number of encoded lines
    Dim strTempLine     As String       'temporary string
    Dim i               As Long         'loop counter
    Dim j               As Integer      'loop counter
    
    'Get file name
    strFileName = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1)
    
    'Insert first marker: "begin 664 ..."
    intTempFile = FreeFile
    TempFileName = TempDirectory & GenerateCode(8)
    Open TempFileName For Binary As intTempFile
    Put intTempFile, , "begin 664 " + strFileName + vbCrLf
    
    'Get file size
    lFileSize = FileLen(strFilePath)
    lEncodedLines = lFileSize \ 45 + 1
    
    'Prepare buffer to retrieve data from the file by 45 symbols chunks
    strFileData = Space(45)
    
    intFile = FreeFile
    
    Open strFilePath For Binary As intFile
        For i = 1 To lEncodedLines
            DoEvents
            
            On Error Resume Next
            frmMain.PB.Value = i * 100 / lEncodedLines
            
            'Read file data by 45-bytes cnunks
            If i = lEncodedLines Then strFileData = Space(lFileSize Mod 45)

            'Retrieve data chunk from file to the buffer
            Get intFile, , strFileData
            
            'Add first symbol to encoded string that informs
            'about quantity of symbols in encoded string.
            'More often "M" symbol is used.
            strTempLine = Chr(Len(strFileData) + 32)
            
                
            'If the last line is processed and length of
            'source data is not a number divisible by 3, add one or two
            'blankspace symbols
            If i = lEncodedLines And (Len(strFileData) Mod 3) Then strFileData = strFileData + Space(3 - (Len(strFileData) Mod 3))
            
            For j = 1 To Len(strFileData) Step 3
                'Breake each 3 (8-bits) bytes to 4 (6-bits) bytes
                '1 byte
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j, 1)) \ 4 + 32)
                '2 byte
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j, 1)) Mod 4) * 16 _
                               + Asc(Mid(strFileData, j + 1, 1)) \ 16 + 32)
                '3 byte
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j + 1, 1)) Mod 16) * 4 _
                               + Asc(Mid(strFileData, j + 2, 1)) \ 64 + 32)
                '4 byte
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j + 2, 1)) Mod 64 + 32)
            Next j
            
            'replace " " with "`"
            strTempLine = Replace(strTempLine, " ", "`")
            'add encoded line to result buffer
            
            'strResult = strResult + strTempLine + vbCrLf
            Put intTempFile, , strTempLine + vbCrLf
            
            strTempLine = ""
        Next i
    Close intFile

    'add the end marker
    'strResult = strResult & "`" & vbCrLf + "end" + vbCrLf
    Put intTempFile, , "`" & vbCrLf + "end" + vbCrLf
    Close intTempFile
    
    
    Dim fso, f, ts
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(TempFileName)
    Set ts = f.OpenAsTextStream(1, -2)
    UUEncodeFile = ts.ReadAll
    ts.Close
    
    Kill TempFileName
End Function


