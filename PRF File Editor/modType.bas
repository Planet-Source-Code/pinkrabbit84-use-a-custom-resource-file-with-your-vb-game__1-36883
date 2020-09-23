Attribute VB_Name = "modTypeAndConst"
Public Const PRF_Version = "PRFile_v1.00"

Public Type PRF_Header
  PRFversion As String * 12
  NbFile As Long
  FileLenght As Long
End Type

Public Type FileHeader
  FileName As String
  FileType As Byte
  FileLenght As Long
  StartAt As Long
  EndAt As Long
End Type

Public Type TempFileHeader
  TempFile As String
  FH As FileHeader
End Type

