Attribute VB_Name = "modOpenPCX"
'Palette type.
Type IMyRGB
  r As Byte
  g As Byte
  b As Byte
End Type
' PCX  header
Type PCXHeader
  PCXFlag As Byte
  PCXVersion As Byte
  RunLengthEncode As Byte
  BitsPerPixel As Byte
  XStart As Integer
  YStart As Integer
  XEnd As Integer
  YEnd As Integer
  HorResolution As Integer
  VerResolution As Integer
  ColorMap(0 To 15) As IMyRGB
  Reserved As Byte
  NumPlanes As Byte
  BytesPerLine As Integer
  PaletteInterp As Integer
  AlsoReserved(0 To 57) As Byte
End Type
             
Type PaletteTable
  PaletteFlag As Byte
  Palette(0 To 255) As IMyRGB
End Type
Function IsPCX(Filename As String) As Boolean
  Open Filename For Binary Access Read As #1
    Dim ph As PCXHeader
    Get #1, , ph.PCXFlag
    Get #1, , ph.PCXVersion
  Close #1
  If ph.PCXFlag <> 10 Then IsPCX = False Else IsPCX = True
End Function
'Sub LoadPCX(Filename As String, ByRef pImage As ImageFile)
  Dim xsize As Long, ysize As Long, bpl As Long, bmbpl As Long
  Dim i As Long, j As Long, Offset As Long, count As Long
  Dim run As Long, counter As Long
  Dim ph As PCXHeader
  Dim pl(0 To 768) As Byte
  Dim buf() As Byte
    
  Open Filename For Binary Access Read As #1
    Get #1, , ph.PCXFlag
    Get #1, , ph.PCXVersion
    Get #1, , ph.RunLengthEncode
    Get #1, , ph.BitsPerPixel
    ph.XStart = fgetWord(1, True)
    ph.YStart = fgetWord(1, True)
    ph.XEnd = fgetWord(1, True)
    ph.YEnd = fgetWord(1, True)
    ph.HorResolution = fgetWord(1, True)
    ph.VerResolution = fgetWord(1, True)
    Get #1, , ph.ColorMap
    Get #1, , ph.Reserved
    Get #1, , ph.NumPlanes
    ph.BytesPerLine = fgetWord(1, True)
    ph.PaletteInterp = fgetWord(1, True)
    Get #1, , ph.AlsoReserved
    If ph.PCXFlag <> 10 Then
      'invalid PCX file
      Close #1
    Exit Sub
    End If
        pImage.ImageWidth = ph.XEnd - ph.XStart + 1
        pImage.ImageHeight = ph.YEnd - ph.YStart + 1
        pImage.ImageBPP = ph.NumPlanes * 8
        Erase pImage.ImagePalette
        'Get Data
        ReDim buf(LOF(1) - Len(ph))
        Get #1, , buf
        
        If pImage.ImageBPP = 8 Then
            ' Get Palette
            Seek #1, LOF(1) - 768
            Get #1, , pl
                If pl(0) <> 12 Then
                    'invalid palette
                    Close #1
                    Exit Sub
            End If
        
            'fill out the palette
            ReDim pImage.ImagePalette(0 To 255)
            For i = 0 To 255
                pImage.ImagePalette(i).rgbRed = pl((i * 3) + 1)
                pImage.ImagePalette(i).rgbGreen = pl((i * 3) + 2)
                pImage.ImagePalette(i).rgbBlue = pl((i * 3) + 3)
            Next i
        End If
    Close #1
    
    bpl = ph.NumPlanes * ph.BytesPerLine
   
    If pImage.ImageBPP = 8 Then
        ReDim pImage.ImageData(1 To pImage.ImageWidth * pImage.ImageHeight)
        'Decompress the PCX file
        For i = 0 To pImage.ImageHeight - 1
            Do Until count >= bpl
                If (buf(Offset) And &HC0) = &HC0 Then
                    run = buf(Offset) And &H3F
                    If run + count > bpl Then run = pImage.ImageWidth - count
                    'Repeat
                    For counter = 0 To run - 1
                        pImage.ImageData(((i * pImage.ImageWidth) + (count + counter)) + 1) = buf(Offset + 1)
                    Next counter
                    count = count + counter
                    'Increase our data offset.
                    Offset = Offset + 2
                Else
                    'If this isn't a 'counter' byte
                    'Put the data straight into our image
                    pImage.ImageData(((i * pImage.ImageWidth) + count) + 1) = buf(Offset)
                    count = count + 1
                    'And increase the data count by 1
                    Offset = Offset + 1
                End If
            Loop
            count = 0
        Next i
    Else
        j = 0
        ReDim pImage.ImageData(1 To (pImage.ImageWidth * pImage.ImageHeight * 3&))
        Dim DecodePlanes() As Byte
        ReDim DecodePlanes(pImage.ImageWidth * pImage.ImageHeight * 3&)
        For i = 0 To UBound(buf)
            If (buf(i) And &HC0) = &HC0 Then
                'Encoded data
                run = buf(i) And &H3F
                
                For count = 0 To run - 1
                    DecodePlanes(j + count) = buf(i + 1)
                Next count
                
                i = i + 1 ' Double incrememnt due to coutn byte
                j = j + run 'Indicate how many pixels added
            Else
                'Not Encoded
                DecodePlanes(j) = buf(i)
                j = j + 1
            End If
        Next i
        Dim Offset2 As Long
        Offset2 = 1
        For i = 0 To pImage.ImageHeight - 1
            For j = 0 To pImage.ImageWidth - 1
                Offset = i * ph.BytesPerLine * 3 + j
                pImage.ImageData(Offset2) = DecodePlanes(Offset + ph.BytesPerLine * 2)
                pImage.ImageData(Offset2 + 1) = DecodePlanes(Offset + ph.BytesPerLine)
                pImage.ImageData(Offset2 + 2) = DecodePlanes(Offset)
                Offset2 = Offset2 + 3
            Next j
        Next i
        Erase DecodePlanes
    End If
    Erase buf
End Sub

