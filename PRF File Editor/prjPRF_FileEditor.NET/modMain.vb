Option Strict Off
Option Explicit On
Module modMain
	Public PictureExt() As String
	Public SoundExt() As String
	Public TextExt() As String
	Public VariousExt() As String
	Public VideoExt() As String
	
	Function CreateDefault() As Object
		ReDim PictureExt(8)
		ReDim SoundExt(3)
		ReDim TextExt(2)
		ReDim VariousExt(4)
		ReDim VideoExt(5)
		
		PictureExt(0) = "bmp"
		PictureExt(1) = "jpg"
		PictureExt(2) = "jpeg"
		PictureExt(3) = "gif"
		PictureExt(4) = "tga"
		PictureExt(5) = "iff"
		PictureExt(6) = "pcx"
		PictureExt(7) = "lbm"
		
		SoundExt(0) = "wav"
		SoundExt(1) = "mid"
		SoundExt(2) = "mp3"
		
		TextExt(0) = "txt"
		
		VariousExt(0) = "exe"
		VariousExt(1) = "zip"
		VariousExt(2) = "rar"
		VariousExt(3) = "ace"
		
		VideoExt(0) = "avi"
		VideoExt(1) = "mpg"
		VideoExt(2) = "mpeg"
		VideoExt(3) = "asf"
		VideoExt(4) = "mov"
		
		SaveBindExt()
	End Function
	
	Function SaveBindExt() As Object
		Dim Fno As Short
		
		Fno = FreeFile
		FileOpen(Fno, VB6.GetPath & "\BindExt.pfe", OpenMode.Binary, OpenAccess.Write, OpenShare.LockReadWrite)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FilePut(Fno, UBound(PictureExt))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FilePut(Fno, PictureExt)
		
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FilePut(Fno, UBound(SoundExt))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FilePut(Fno, SoundExt)
		
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FilePut(Fno, UBound(TextExt))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FilePut(Fno, TextExt)
		
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FilePut(Fno, UBound(VariousExt))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FilePut(Fno, VariousExt)
		
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FilePut(Fno, UBound(VideoExt))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FilePut(Fno, VideoExt)
		FileClose()
	End Function
	
	Function LoadBindExt() As Object
		Dim Fno As Short
		Dim ArraySize As Integer
		
		Fno = FreeFile
		FileOpen(Fno, VB6.GetPath & "\BindExt.pfe", OpenMode.Binary, OpenAccess.Read, OpenShare.LockReadWrite)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(Fno, ArraySize)
		ReDim PictureExt(ArraySize)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(Fno, PictureExt)
		
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(Fno, ArraySize)
		ReDim SoundExt(ArraySize)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(Fno, SoundExt)
		
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(Fno, ArraySize)
		ReDim TextExt(ArraySize)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(Fno, TextExt)
		
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(Fno, ArraySize)
		ReDim VariousExt(ArraySize)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(Fno, VariousExt)
		
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(Fno, ArraySize)
		ReDim VideoExt(ArraySize)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(Fno, VideoExt)
		FileClose()
	End Function
End Module