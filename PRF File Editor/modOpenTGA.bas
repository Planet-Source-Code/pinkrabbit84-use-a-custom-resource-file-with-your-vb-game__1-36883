Attribute VB_Name = "modOpenTGA"
' Tagar file header
Type TRGHDR
    bIDFieldSize As Byte           'Characters in ID field
    bClrMapType As Byte            'Color map type
    bImageType As Byte             'Image type
    lClrMapSpec(0 To 4) As Byte    'Color map specification
    wXOrigin As Integer              'X origin
    wYOrigin As Integer              'Y origin
    wWidth As Long                 'Bitmap width
    wHeight As Long                'Bitmap height
    bBitsPixel As Byte             'Bits per pixel
    bImageDescriptor As Byte       'Image descriptor
End Type
Dim TempInt As Integer
Dim thead As TRGHDR

Function IsTGA(Filename As String)
    Open Filename For Binary Access Read As #1
        'read filehead
        Get #1, , thead.bIDFieldSize
        Get #1, , thead.bClrMapType
        Get #1, , thead.bImageType
        Get #1, , thead.lClrMapSpec
        Get #1, , thead.wXOrigin
        Get #1, , thead.wYOrigin
        Get #1, , TempInt
        Get #1, , TempInt
        Get #1, , thead.bBitsPixel
        Get #1, , thead.bImageDescriptor
        If (thead.bImageType <> 2) Or (thead.bBitsPixel <> 24) Then IsTGA = False Else IsTGA = True 'not an TGA 24 bits file
    Close #1
End Function
Sub LoadTGA(Filename As String, ByRef pImage As ImageFile)
    Dim buf() As Byte
    Dim i As Long, j As Long
    Open Filename For Binary Access Read As #1
        'read filehead
        Get #1, , thead.bIDFieldSize
        Get #1, , thead.bClrMapType
        Get #1, , thead.bImageType
        Get #1, , thead.lClrMapSpec
        Get #1, , thead.wXOrigin
        Get #1, , thead.wYOrigin
        Get #1, , TempInt
        thead.wWidth = TempInt
        Get #1, , TempInt
        thead.wHeight = TempInt
        Get #1, , thead.bBitsPixel
        Get #1, , thead.bImageDescriptor

        If (thead.bImageType <> 2) Or (thead.bBitsPixel <> 24) Then Close #1: Exit Sub   'not an TGA 24 bits file
        
        pImage.ImageBPP = 24
        pImage.ImageWidth = thead.wWidth
        pImage.ImageHeight = thead.wHeight
        Erase pImage.ImagePalette
        ReDim pImage.ImageData(1 To (CLng(thead.wWidth) * 3) * thead.wHeight) As Byte
        ReDim buf(1 To thead.wWidth * 3)
        Dim offset1 As Long, Offset2 As Long
        offset1 = UBound(pImage.ImageData)
        offset1 = offset1 - ((pImage.ImageWidth * 3) - 1)
        For i = thead.wHeight - 1 To 0 Step -1
            Get #1, , buf
            CopyMemory pImage.ImageData(offset1), buf(1), (pImage.ImageWidth * 3)
            offset1 = offset1 - (pImage.ImageWidth * 3)
        Next i
    Close #1
End Sub
Sub SaveTGA(Filename As String, ByRef pImage As ImageFile)
    Dim buf() As Byte
    Dim i As Long, j As Long
    ReDim buf(1 To pImage.ImageWidth * 3)
    If Dir(Filename) <> "" Then Kill Filename
    Open Filename For Binary Access Write As #1
        thead.bIDFieldSize = 0
        thead.bClrMapType = 0
        thead.bImageType = 2
        For i = 0 To 4
            thead.lClrMapSpec(i) = 0
        Next i
        thead.wXOrigin = 0
        thead.wYOrigin = 0
        thead.wWidth = pImage.ImageWidth
        thead.wHeight = pImage.ImageHeight
        thead.bBitsPixel = 24
        thead.bImageDescriptor = 32         '0x20

        'write filehead'
        Put #1, , thead.bIDFieldSize
        Put #1, , thead.bClrMapType
        Put #1, , thead.bImageType
        Put #1, , thead.lClrMapSpec
        Put #1, , thead.wXOrigin
        Put #1, , thead.wYOrigin
        TempInt = thead.wWidth
        Put #1, , TempInt
        TempInt = thead.wHeight
        Put #1, , TempInt
        Put #1, , thead.bBitsPixel
        Put #1, , thead.bImageDescriptor
        
        Dim offset1 As Long, Offset2 As Long
        Select Case pImage.ImageBPP
            Case 24
                offset1 = 1
                For i = 0 To pImage.ImageHeight - 1
                    Offset2 = 1
                    For j = 0 To pImage.ImageWidth - 1
                        buf(Offset2) = pImage.ImageData(offset1)
                        buf(Offset2 + 1) = pImage.ImageData(offset1 + 1)
                        buf(Offset2 + 2) = pImage.ImageData(offset1 + 2)
                        offset1 = offset1 + 3
                        Offset2 = Offset2 + 3
                   Next j
                   Put #1, , buf
                Next i
            Case 8
                offset1 = 1
                offset1 = UBound(pImage.ImageData)
                offset1 = offset1 - (pImage.ImageWidth - 1)
                For i = 0 To pImage.ImageHeight - 1
                    Offset2 = 1
                    ''bgr
                    For j = 0 To pImage.ImageWidth - 1
                        buf(Offset2 + 2) = pImage.ImagePalette(pImage.ImageData(offset1)).rgbBlue
                        buf(Offset2 + 1) = pImage.ImagePalette(pImage.ImageData(offset1)).rgbGreen
                        buf(Offset2) = pImage.ImagePalette(pImage.ImageData(offset1)).rgbRed
                        offset1 = offset1 + 1
                        Offset2 = Offset2 + 3
                    Next j
                    offset1 = offset1 - (pImage.ImageWidth)
                    offset1 = offset1 - (pImage.ImageWidth)
                    Put #1, , buf
                Next i
        End Select
    Close #1
End Sub

