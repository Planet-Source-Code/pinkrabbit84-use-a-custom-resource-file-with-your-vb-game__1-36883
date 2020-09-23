Attribute VB_Name = "modMain"
Public PictureExt() As String
Public SoundExt() As String
Public TextExt() As String
Public VariousExt() As String
Public VideoExt() As String

Function CreateDefault()
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
  
  SaveBindExt
End Function

Function SaveBindExt()
  Dim Fno As Integer
  
  Fno = FreeFile
  Open App.Path & "\BindExt.pfe" For Binary Access Write Lock Read Write As Fno
    Put Fno, , UBound(PictureExt)
    Put Fno, , PictureExt
    
    Put Fno, , UBound(SoundExt)
    Put Fno, , SoundExt

    Put Fno, , UBound(TextExt)
    Put Fno, , TextExt

    Put Fno, , UBound(VariousExt)
    Put Fno, , VariousExt

    Put Fno, , UBound(VideoExt)
    Put Fno, , VideoExt
  Close
End Function

Function LoadBindExt()
  Dim Fno As Integer
  Dim ArraySize As Long
  
  Fno = FreeFile
  Open App.Path & "\BindExt.pfe" For Binary Access Read Lock Read Write As Fno
    Get Fno, , ArraySize
    ReDim PictureExt(ArraySize)
    Get Fno, , PictureExt
    
    Get Fno, , ArraySize
    ReDim SoundExt(ArraySize)
    Get Fno, , SoundExt

    Get Fno, , ArraySize
    ReDim TextExt(ArraySize)
    Get Fno, , TextExt
  
    Get Fno, , ArraySize
    ReDim VariousExt(ArraySize)
    Get Fno, , VariousExt
  
    Get Fno, , ArraySize
    ReDim VideoExt(ArraySize)
    Get Fno, , VideoExt
  Close
End Function
