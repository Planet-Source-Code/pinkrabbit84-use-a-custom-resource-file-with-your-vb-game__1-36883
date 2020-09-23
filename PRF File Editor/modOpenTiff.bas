Attribute VB_Name = "modOpenTiff"
'TIF_header_struct defined
Type TIF_HEADER
     byte_order(2) As Byte
     version As Integer
     Offset As Long
End Type
'Directory struct define
Type TIF_ENTRY
      tag As Integer
      type As Integer
      length As Integer
      Offset As Integer
End Type

'  tiff tag names
Const NewSubfile = 254
Const SubfileType = 255
Const ImageWidth = 256
Const ImageLength = 257
Const BitsPerSample = 258
Const Compression = 259
Const StripOffsets = 273
Const RowsPerStrip = 278
Const StripByteCounts = 279
Const SamplesPerPixel = 277
Const PlanarConfiguration = 284
Const Group3Options = 292
Const Group4Options = 293
Const FillOrder = 266
Const Threshholding = 263
Const CellWidth = 264
Const CellLength = 265
Const MinSampleValue = 280
Const MaxSampleValue = 281
Const PhotometricInterp = 262
Const GrayResponseUnit = 290
Const GrayResponseCurve = 291
Const ColorResponseUnit = 300
Const ColorResponseCurves = 301
Const XResolution = 282
Const YResolution = 283
Const ResolutionUnit = 296
Const Orientation = 274
Const DocumentName = 269
Const PageName = 285
Const XPosition = 286
Const YPosition = 287
Const PageNumber = 297
Const ImageDescription = 270
Const Make = 271
Const Model = 272
Const FreeOffsets = 288
Const FreeByteCounts = 289
Const Predictor = 317
Const tagPALETTE = 320

' tiff size
Const TIFFbyte = 1
Const TIFFascii = 2
Const TIFFshort = 3
Const TIFFlong = 4
Const TIFFrational = 5

'  tiff compression types
Const COMPnone = 1
Const COMPhuff = 2
Const COMPfax3 = 3
Const COMPfax4 = 4
Const COMPwrd1 = 32771   '0x8003
Const COMPmpnt = 32773   '0x8005

Function IsTIFF(Filename As String) As Boolean
    Dim i As Long
    FilePointer = 1
    Open Filename For Binary Access Read As #1
        i = fgetWord(1, True)
        If i = 18761 Then                       '"II" or 0x4949
            IsTIFF = True
        ElseIf i = 19789 Then                   '"MM" or 0x4d4d
            IsTIFF = True
        Else                                    ' Not a TIFF file
            IsTIFF = False
        End If
    Close #1
End Function
Sub LoadTIFF(Filename As String, ByRef pImage As ImageFile)
    Dim buf() As Byte
    Dim i As Long, k As Long, j As Long, entry As Long, nc As Long, nr As Long
    Dim tif_head As TIF_HEADER
    Dim tifen As TIF_ENTRY
    Dim bw As Long, unreg As Long, tmp As Long
    Dim intel As Boolean
    Dim offset1 As Long, Offset2 As Long
    FilePointer = 1
    Open Filename For Binary Access Read As #1
        i = fgetWord(1, True)
        If i = 18761 Then                       '"II" or 0x4949
            intel = True
        ElseIf i = 19789 Then                   '"MM" or 0x4d4d
            intel = False
        Else                                    ' Not a TIFF file
            Close #1
            Exit Sub
        End If
        tif_head.version = fgetWord(1, intel)   ' read filehead
        If tif_head.version <> 42 Then          ' Not a TIFF file
            Close #1
            Exit Sub
        End If
        
        tif_head.Offset = fgetLong(1, intel)    'read filehead
        offset1 = tif_head.Offset               'get the direction offset
        FilePointer = offset1 + 1
        entry = fgetWord(1, intel)
        For i = 0 To entry - 1                 'Deal with Entry
            tifen.tag = fgetWord(1, intel)
            tifen.type = fgetWord(1, intel)
            If tifen.type = TIFFlong Then
                tifen.length = fgetLong(1, intel)
                tifen.Offset = fgetLong(1, intel)
            Else
                tifen.length = fgetWord(1, intel)
                fgetWord 1, intel
                tifen.Offset = fgetWord(1, intel)
                tmp = fgetWord(1, intel)
            End If

            Select Case tifen.tag
                Case SubfileType, Compression, PlanarConfiguration
                    If tifen.Offset <> 1 Then 'TIFF file not supported
                        Close #1
                        Exit Sub
                    End If
                Case ImageWidth
                    nc = tifen.Offset
                Case ImageLength
                    nr = tifen.Offset
                Case BitsPerSample
                    If tifen.length <> 3 Then ' Not a 24 bits TIFF file
                        Close #1
                        Exit Sub
                    End If
                Case PhotometricInterp
                    tmp = tifen.Offset
                    If (tmp <> 2) And (tmp <> 3) And (tmp <> 1) Then unreg = 1
                    If tmp = 1 Then bw = 1 ' reversed 0 white 255 black
                Case StripOffsets
                    Offset2 = tifen.Offset
            End Select
        Next i
        If nc = 0 Or nr = 0 Then             'not a valid TIFF file
            Close #1
            Exit Sub
        End If
        pImage.ImageBPP = 24
        pImage.ImageWidth = nc
        pImage.ImageHeight = nr
        Erase pImage.ImagePalette
        ReDim pImage.ImageData((pImage.ImageWidth * 3) * pImage.ImageHeight)
        ReDim buf(0 To nc * 3) As Byte
        Dim Offset As Long
        If unreg = 0 Then                     'Read image data
            FilePointer = nc + Offset2 - 1
            Offset = 1
            For i = 0 To nr - 1
                Get #1, FilePointer, buf
                FilePointer = FilePointer + (nc * 3)
                k = 0
                For j = 0 To nc - 1
                    pImage.ImageData(Offset) = buf(k + 2)
                    pImage.ImageData(Offset + 1) = buf(k + 1)
                    pImage.ImageData(Offset + 2) = buf(k)
                    k = k + 3
                    Offset = Offset + 3
                Next j
            Next i
        End If
    Close #1
End Sub

