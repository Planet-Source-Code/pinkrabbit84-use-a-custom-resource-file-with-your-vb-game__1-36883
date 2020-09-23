Option Strict Off
Option Explicit On
Module modTypeAndConst
	Public Const PRF_Version As String = "PRFile_v1.00"
	
	Public Structure PRF_Header
		<VBFixedString(12),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=12)> Public PRFversion As String
		Dim NbFile As Integer
		Dim FileLenght As Integer
	End Structure
	
	Public Structure FileHeader
		Dim FileName As String
		Dim FileType As Byte
		Dim FileLenght As Integer
		Dim StartAt As Integer
		Dim EndAt As Integer
	End Structure
	
	Public Structure TempFileHeader
		Dim TempFile As String
		Dim FH As FileHeader
	End Structure
End Module