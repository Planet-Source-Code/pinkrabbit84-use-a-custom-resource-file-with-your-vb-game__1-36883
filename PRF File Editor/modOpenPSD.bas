Attribute VB_Name = "modOpenPSD"
Type PSDInfo
  Pixels() As Byte
  BitsPerChannel As Integer
  ColorData(0 To 767) As Byte    'For Indexed or DuoTone only
  Mode As Integer
  Width As Long
  Height As Long
  ChannelCount As Integer
  Compression As Integer
End Type
Dim ImageInfo As PSDInfo
Dim b3 As Byte, b2 As Byte, b1 As Byte, b0 As Byte
Dim FilePointer As Long

Private Function Read32(filenumber) As Long
    Get #filenumber, FilePointer, b3
    FilePointer = FilePointer + 1
    Get #filenumber, FilePointer, b2
    FilePointer = FilePointer + 1
    Get #filenumber, FilePointer, b1
    FilePointer = FilePointer + 1
    Get #filenumber, FilePointer, b0
    FilePointer = FilePointer + 1
    Read32 = LShift(b3, 24) + LShift(b2, 16) + LShift(b1, 8) + b0
End Function
Private Function Read16(filenumber) As Long
    Get #filenumber, FilePointer, b1
    FilePointer = FilePointer + 1
    Get #filenumber, FilePointer, b0
    FilePointer = FilePointer + 1
    Read16 = b0 + LShift(b1, 8)
End Function
Function IsPSD(Filename As String) As Boolean
On Error GoTo errorout:
    Open Filename For Binary Access Read As #1
        FilePointer = 1
        IType = Read32(1)
        If IType <> 943870035 Then IsPSD = False Else IsPSD = True
    Close #1
    Exit Function
errorout:
IsPSD = False
Close #1
End Function

Sub LoadPSD(Filename As String, pImage As ImageFile)

    Dim i As Long, j As Long, k As Long
    Dim IType As Long
    Dim ModeDataCount As Long, ResourceDataCount As Long, ReservedDataCount As Long
    Dim PSDVersion As Integer
    FilePointer = 1
    ' Firstopen the file and get for us important entries from the header...
    Open Filename For Binary Access Read As #1
        IType = Read32(1)
        If IType <> 943870035 Then Close #1: Exit Sub       'Not a PSD File.
        PSDVersion = Read16(1)
        If PSDVersion <> 1 Then Close #1: Exit Sub          'Incorrect PSD Version, MUST be 1.
        ' Skip 6 Bytes, irrelevant info. Must be 0
        FilePointer = FilePointer + 6
        ImageInfo.ChannelCount = Read16(1)
        If ImageInfo.ChannelCount < 0 Or ImageInfo.ChannelCount > 16 Then Close #1: Exit Sub 'Incorrect Channel Count
        ImageInfo.Height = Read32(1)
        ImageInfo.Width = Read32(1)
        ImageInfo.BitsPerChannel = Read16(1)      'Supported values are 1,8 or 16
        If ImageInfo.BitsPerChannel <> 8 Then Close #1: Exit Sub  'NO RGB COLOURS
        ' Make sure the color mode is RGB.
        ' Supported Modes are Bitmap=0, Grayscale=1, Indexed=2,RGB=3,CMYK=4,MultiChannel=7
        ' Duotone=8,Lab=9
        ImageInfo.Mode = Read16(1)
        If ImageInfo.Mode <> 3 Then Close #1: Exit Sub      'ColorMode is Not RGB
        ' Skip the Mode Data. (It's the palette for indexed color; other info for other modes.)
        ModeDataCount = Read32(1)
        If ModeDataCount <> 0 Then FilePointer = FilePointer + ModeDataCount
        ' Skip the image resources. (resolution, pen tool paths, etc)
        ResourceDataCount = Read32(1)
        If ResourceDataCount <> 0 Then FilePointer = FilePointer + ResourceDataCount
        ' Skip the reserved data.
        ReservedDataCount = Read32(1)
        If ReservedDataCount <> 0 Then FilePointer = FilePointer + ReservedDataCount
        ' Find out if the data is compressed.
        ImageInfo.Compression = Read16(1)
        'Compression Type 0=Raw Data, RLE Compressed = 1
        If ImageInfo.Compression > 1 Then Close #1: Exit Sub  'Compression Type Not Supported
        ' Decode Image...
        ReDim ImageInfo.Pixels(0 To (4 * ImageInfo.Height * ImageInfo.Width) + 2) As Byte
        DecodePSD 1
        Close #1
        'Copy this data into our custom image object (which was passed).
        pImage.ImageBPP = 24
        pImage.ImageWidth = ImageInfo.Width
        pImage.ImageHeight = ImageInfo.Height
        Erase pImage.ImagePalette
        ReDim pImage.ImageData(1 To ((pImage.ImageWidth * pImage.ImageHeight) * 3))
        Dim offset1 As Long, Offset2 As Long
        offset1 = 1
        Offset2 = 0
        For i = 0 To ImageInfo.Height - 1
            For j = 0 To ImageInfo.Width - 1
                pImage.ImageData(offset1) = ImageInfo.Pixels((Offset2 * 4))
                offset1 = offset1 + 1
                pImage.ImageData(offset1) = ImageInfo.Pixels((Offset2 * 4) + 1)
                offset1 = offset1 + 1
                pImage.ImageData(offset1) = ImageInfo.Pixels((Offset2 * 4) + 2)
                offset1 = offset1 + 1
                'Skip Alpha Pixel (which would be (offset2 *4) +1)
                Offset2 = Offset2 + 1
            Next j
        Next i
        Erase ImageInfo.Pixels
End Sub
Sub SavePSD(Filename As String, pImage As ImageFile)

End Sub

Sub DecodePSD(filenumber As Integer)
    'NOTE: This function (including the DecodePSD function) are VERY slow whilst running in the IDE
    '      For accurate load time results, please compile the EXE first.
    Dim Default(0 To 3) As Long
    Dim chn(0 To 3) As Long
    Dim PixelCount As Long
    Dim c As Long, n As Long, pn As Long, channel As Long, count As Long, ilen As Long, ival As Byte
    Default(0) = 0
    Default(1) = 0
    Default(2) = 0
    Default(3) = 255
    chn(0) = 2
    chn(1) = 1
    chn(2) = 0
    chn(3) = 3
    Dim FileContainer() As Byte
    ReDim FileContainer(0 To LOF(filenumber) - FilePointer)
    Get #1, FilePointer, FileContainer
    FilePointer = 0
    PixelCount = ImageInfo.Width * ImageInfo.Height
    If ImageInfo.Compression Then
        FilePointer = FilePointer + ImageInfo.Height * ImageInfo.ChannelCount * 2
        For c = 0 To 3
            pn = 0
            channel = chn(c)
            If channel >= ImageInfo.ChannelCount Then
                For pn = 0 To PixelCount - 1
                    ImageInfo.Pixels((pn * 4) + channel) = Default(channel)
                Next pn
            Else
                count = 0
                Do Until (count >= PixelCount)
                    ilen = FileContainer(FilePointer)
                    FilePointer = FilePointer + 1
                    If ilen = 128 Then
                    ElseIf ilen < 128 Then
                        ilen = ilen + 1
                        count = count + ilen
                        Do Until ilen = 0
                            ImageInfo.Pixels((pn * 4) + channel) = FileContainer(FilePointer)
                            FilePointer = FilePointer + 1
                            pn = pn + 1
                            ilen = ilen - 1
                        Loop
                    ElseIf ilen > 128 Then
                        ilen = ilen Xor 255
                        ilen = ilen + 2
                        ival = FileContainer(FilePointer)
                        FilePointer = FilePointer + 1
                        count = count + ilen
                        Do Until ilen = 0
                            ImageInfo.Pixels((pn * 4) + channel) = ival
                            pn = pn + 1
                            ilen = ilen - 1
                        Loop
                    End If
                Loop
            End If
        Next c
    Else
        For c = 0 To 3
            channel = chn(c)
            If channel > ImageInfo.ChannelCount Then
                For pn = 0 To PixelCount - 1
                    ImageInfo.Pixels((pn * 4) + channel) = Default(channel)
                Next pn
            Else
                For n = 0 To PixelCount - 1
                    ImageInfo.Pixels((n * 4) + channel) = FileContainer(FilePointer)
                    FilePointer = FilePointer + 1
                Next n
            End If
        Next c
    End If
    Erase FileContainer
End Sub


