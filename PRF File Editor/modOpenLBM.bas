Attribute VB_Name = "modOpenLBM"
Dim Header As String * 4
Dim PictureWidth As Long
Dim PictureHeight As Long
Dim BitsPerPixel As Byte
Dim NumOfColours As Integer
Dim Packed As Byte
Dim TempByte As Byte
Function IsLBM(Filename As String) As Boolean
  Dim Fno As Integer
  
  Fno = FreeFile
  Open Filename For Binary Access Read As Fno
    'Is this real a DPaint pic ?
    Get Fno, , Header
    If Header <> "FORM" Then IsLBM = False Else IsLBM = True
  Close Fno
End Function
Sub LoadLBM(Filename As String, pImage As ImageFile)
  Dim Fno As Integer
  
  Fno = FreeFile
  Open Filename For Binary Access Read As Fno
    'Check to see if file is valid
    Get Fno, , Header
    If Header <> "FORM" Then Close Fno: Exit Sub '(Not Dpaint Picture)
    'Find the Information Header
    For f = 9 To LOF(1)
      Get Fno, f, Header
      If Header = "BMHD" Then Seek Fno, f + 8: Exit For
    Next f
    'Get Picture Information
    Get Fno, , TempByte
    PictureWidth = TempByte * 256
    Get Fno, , TempByte
    PictureWidth = PictureWidth + TempByte
    Get Fno, , TempByte
    PictureHeight = TempByte * 256
    Get Fno, , TempByte
    PictureHeight = PictureHeight + TempByte
    Seek Fno, Seek(1) + 4
    Get Fno, , BitsPerPixel
    If BitsPerPixel = 1 Then NumOfColours = 2
    If BitsPerPixel = 4 Then NumOfColours = 16
    If BitsPerPixel = 8 Then NumOfColours = 256
    Seek Fno, Seek(1) + 1
    Get Fno, , Packed
    If BitsPerPixel = 8 And Packed = 1 Then
      Dim a As Byte, b As Byte, g As Byte
      Dim xX As Long, yY As Long, tlong As Long
      Dim Offset As Long
      pImage.ImageWidth = PictureWidth
      pImage.ImageHeight = PictureHeight
      pImage.ImageBPP = BitsPerPixel
      ReDim pImage.ImagePalette(0 To 255)
      Seek Fno, 1
      'get the palette (or colour map)
      For f = 1 To LOF(1)
        Get Fno, f, Header
        If Header = "CMAP" Then Seek Fno, f + 8: Exit For
      Next f
      For i = 0 To NumOfColours - 1
        Get Fno, , TempByte
        tlong = TempByte
        Get Fno, , TempByte
        tlong = tlong + CLng(TempByte) * 256
        Get Fno, , TempByte
        tlong = tlong + CLng(TempByte) * 65536
        pImage.ImagePalette(i).rgbRed = Int(tlong Mod 256)
        pImage.ImagePalette(i).rgbGreen = Int(tlong / 256) Mod 256
        pImage.ImagePalette(i).rgbBlue = Int(tlong / 65536)
      Next i
      'Find where the picture data starts.
      For f = Seek(1) To (LOF(1) - 4)
        Get Fno, f, Header
        If Header = "BODY" Then Seek Fno, f + 8: Exit For
      Next f
      Dim FileContainer() As Byte, FilePointer As Long
      ReDim FileContainer(1 To (LOF(1) - Loc(1)))
      FilePointer = 1
      Get Fno, Seek(1), FileContainer
      x = 0: y = 0
      'Decompress picture data
      ReDim pImage.ImageData(1 To (pImage.ImageWidth * pImage.ImageHeight))
      Offset = 1
      Do Until FilePointer >= UBound(FileContainer)
        a = FileContainer(FilePointer)
        FilePointer = FilePointer + 1
      If a > 128 Then
        b = FileContainer(FilePointer)
        FilePointer = FilePointer + 1
        For tlong = xX To ((xX + (257 - a)) - 1)
          pImage.ImageData(Offset) = b
          Offset = Offset + 1
        Next tlong
        xX = xX + (257 - a)
        If xX > PictureWidth - 1 Then xX = 0
        Else
          For f = 0 To a
            b = FileContainer(FilePointer)
            FilePointer = FilePointer + 1
            pImage.ImageData(Offset) = b
            Offset = Offset + 1
            xX = x + 1
            If xX > PictureWidth - 1 Then xX = 0
          Next f
        End If
      Loop
      Erase FileContainer
    End If
  Close Fno
End Sub

