Attribute VB_Name = "modTypeAndConst"
Public Const PRF_Version = "PRFile_v1.00"

'This is the main header of the file
Public Type PRF_Header
  PRFversion As String * 12 'Version of the file
  NbFile As Long 'Number of file in the resource
  FileLenght As Long 'The PRFile size (to check if the file is valid (compare this with the file size on disk)
End Type

'And the header for each file you put in the resource
Public Type FileHeader
  FileName As String 'Save the file name
  FileType As Byte 'Save the file type(only used in the editor and soon in the Dll)
  FileLenght As Long 'Save the file size
  StartAt As Long 'Where the file begin
  EndAt As Long 'Where the file finish (Not used because when I begin this project I think it's may be useful but I was wrong and I forgot to remove it (but it's don't matter only take 4Byte per file))
End Type

'This is used only in the code
Public Type TempFileHeader
  TempFile As String
  FH As FileHeader
End Type


