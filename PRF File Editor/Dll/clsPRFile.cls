VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim prgPath As String 'Keep the path of the AppPath (currently not used)
Dim TempDir As String 'Keep the temp dir
Dim PRFilePath As String 'Where the PRFile is
Dim UseDirPath As String 'Default path to extract file
Dim MainHeader As PRF_Header 'The main header of the PRFile
Dim FileHeader() As TempFileHeader 'Array of header  of each file in the resource
Dim ExtractedFile() As String 'Save path of all file extracted (to delete them at end)

'Init the Dll
Public Function Init(AppPath As String, DefaultUsePath As String) As Long
  Dim WinDir As String 'Keep the path of Windows
  
  'If AppPath is a bad path -> return error code
  If Dir(AppPath, vbDirectory) = "" Then Init = 999001: Exit Function
  'If the Default path is bad -> return error code
  If Dir(DefaultUsePath, vbDirectory) = "" Then Init = 999001: Exit Function
  
  'Set the path of the App
  prgPath = AppPath
  'Set the temp dir
  TempDir = prgPath & "\PRFileTemp"
  
  'Test if we can write in the temp dir ---------------------------------------------
  On Error Resume Next
  'If the temp dir does not exist create one
  If Dir(TempDir, vbDirectory) = "" Then MkDir (TempDir)
  'If an error occur (we can do the temp dir
  If Err.Number <> 0 Then
    'Get the path of Windows (without any API)
    WinDir = Environ("WinDir")
    'Create a temp dir in the temp dir of Windows
    If Dir(WinDir & "\Temp\PRFileTemp", vbDirectory) = "" Then MkDir (WinDir & "\Temp\PRFileTemp")
    TempDir = WinDir & "\Temp\PRFileTemp"
  End If
   'Return to normal error handler
  On Error GoTo 0
  '----------------------------------------------------------------------------------
  
  'Set the default path
  UseDirPath = DefaultUsePath
  
  'Init the ExtractedFile var
  ReDim ExtractedFile(0)
  
  'retrun 0 (no error)
  Init = 0
End Function

Public Function SetPRFile(PRFile As String) As Long
  'If the given PRFile does not exist -> return error code
  If Dir(PRFile) = "" Then SetPRFile = 999002: Exit Function
  PRFilePath = PRFile
  
  Dim Fno As Integer
  Dim FnoF As Integer
  Dim CharArray() As Byte
  
  'Get a free file number
  Fno = FreeFile
  'Open the PRFile in binary mode
  Open PRFilePath For Binary Access Read Lock Read Write As Fno
    'Get the mainheader of the PRFile
    Get Fno, 1, MainHeader
    
    'Check the PRFile version (cur: 1.00) (if not the same return error code and close the file)
    If MainHeader.PRFversion <> PRF_Version Then SetPRFile = 999003: Close: Exit Function
    'If the size of the file in the mainheader <> the size of file on disk (someone alter our file) -> return error code and close the file
    If MainHeader.FileLenght <> FileLen(PRFilePath) Then SetPRFile = 999004: Close: Exit Function
    
    'redim the fileheader array
    ReDim FileHeader(MainHeader.NbFile - 1)
    
    'Do a loop for each file in the file (to get each fileheader)
    For i = 0 To MainHeader.NbFile - 1
      'the a fileheader
      Get Fno, , FileHeader(i).FH
      
      'Create a not used temp file name--------------------
      Do
        FileHeader(i).TempFile = TempDir & "\TempFile" & Int(Rnd * 9) & Int(Rnd * 9) & Int(Rnd * 9) & Int(Rnd * 9) & Int(Rnd * 9) & ".tmp"
      Loop Until Dir(FileHeader(i).TempFile) = ""
      '----------------------------------------------------
      
    Next i
    
    'Do a loop for each file in the file (to get each file and copy to temp dir)
    For i = 0 To MainHeader.NbFile - 1
      'Redim CharArray for each char in the file
      ReDim CharArray(FileHeader(i).FH.FileLenght - 1)
      'Get all char of this file
      Get Fno, FileHeader(i).FH.StartAt, CharArray
      
      'Get a free file number
      FnoF = FreeFile
      'Open the temp file
      Open FileHeader(i).TempFile For Binary Access Write Lock Read Write As FnoF
        'Write the file
        Put FnoF, 1, CharArray
      'Close the file
      Close FnoF
      
    Next i
  
  'Close the file
  Close Fno
  
  'return 0 (no error)
  SetPRFile = 0
End Function

Function GetFile(FileName As String, Optional UsePath = "[UseDefault]", Optional OverWrite As Boolean = True) As Long
  Dim ExtractTo As String 'Where the file will be extracted
  Dim Index As Integer
  
  'If user don't give anything here then set the Default as the ExtractTo folder
  If UsePath = "[UseDefault]" Then ExtractTo = UseDirPath Else ExtractTo = UsePath
  
  Index = -1
  'Do a loop for each file
  For i = 0 To MainHeader.NbFile - 1
    'If the given file name = (i don't know how to explane that but check the next line :P)
    If UCase(FileName) = UCase(FileHeader(i).FH.FileName) Then Index = i: Exit For
  Next i
  'If Index = -1 -> (file in does not exist in library) retrun error code
  If Index = -1 Then GetFile = 999005: Exit Function
  
  
  On Error Resume Next
  'if Overwirte = true and the file already exist then delete it
  If Dir(ExtractTo & "\" & FileName) <> "" Then If OverWrite = True Then Kill ExtractTo & "\" & FileName Else GetFile = 999007: Exit Function
  'If error number = 70(permission denied) the file is in use so retrun the error code
  If Err.Number = 70 Then GetFile = 999008: Exit Function
  On Error GoTo 0
  
  'Copy the temp file to the destination
  FileCopy FileHeader(i).TempFile, ExtractTo & "\" & FileName
  
  'if the file was not copied return the error code
  If Dir(ExtractTo & "\" & FileName) = "" Then GetFile = 999006: Exit Function
  
  'Redim the ExtractedFile to add a path
  ReDim Preserve ExtractedFile(UBound(ExtractedFile) + 1)
  'Add the path
  ExtractedFile(UBound(ExtractedFile)) = ExtractTo & "\" & FileName
  
  'return 0 (noerror
  GetFile = 0
End Function

Function GetAllFile(Optional UsePath = "[UseDefault]", Optional OverWrite As Boolean = True) As Long
  Dim LastError As Long
  
  'Do a loop for each file(and call GetFile for each file in the ressource)
  For i = 0 To MainHeader.NbFile - 1
    LastError = GetFile(FileHeader(i).FH.FileName, UsePath, OverWrite)
  Next i
  
  GetAllFile = LastError
End Function

Private Sub Class_Terminate()
  On Error Resume Next
  
  'Remove all file in the temp dir
  Kill TempDir & "\*.*"
  'Remove the temp dir
  RmDir TempDir
  
  'Delete all extracted file
  For i = 0 To UBound(ExtractedFile)
    If Dir(ExtractedFile(i)) <> "" Then Kill ExtractedFile(i)
  Next i
End Sub
