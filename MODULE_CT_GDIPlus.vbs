Option Compare Database
Option Explicit
' HC 5/2010 - added for handling the icon sizing on the ribbon bars
'-------------------------------------------------
' Picture functions using then GDIPlus-API (GDIP) |
'-------------------------------------------------
'    *  Office 2007/2010 version  *               |
'-------------------------------------------------
'   (c) mossSOFT / Sascha Trowitzsch rev. 11/2009 |
'-------------------------------------------------

' # IDBE Avenius
' # 11/2009 angepasst an Office x64
' # http://www.accessribbon.de

'- Reference to library "OLE Automation" (stdole) needed!



'-----------------------------------------------------------------------------------------
'Global module variable:
Private lGDIP As LongPtr
Private bSharedLoad As Boolean
'-----------------------------------------------------------------------------------------


'Initialize GDI+
Function InitGDIP() As Boolean
    Dim TGDP As GDIPStartupInput
    Dim hMod As Long
    
    If lGDIP = 0 Then
        If IsNull(TempVars("GDIPlusHandle")) Then   'If lGDIP is broken due to unhandled errors restore it from the TempVars collection
            TGDP.GdiplusVersion = 1
            hMod = GetModuleHandle("gdiplus.dll")   'gdiplus.dll not yet loaded?
            If hMod = 0 Then
                hMod = LoadLibrary("gdiplus.dll")
                bSharedLoad = False
            Else
                bSharedLoad = True
            End If
            GdiplusStartup lGDIP, TGDP  'Get a personal instance of gdiplus
            TempVars("GDIPlusHandle") = lGDIP
        Else
            lGDIP = TempVars("GDIPlusHandle")
        End If
        AutoShutDown
    End If
    'InitGDIP = (lGDIP > 0)
    ' HC invalid handle = 0
    InitGDIP = (lGDIP <> 0)
End Function

'Clear GDI+
Sub ShutDownGDIP()
    If lGDIP <> 0 Then
        If KillTimer(0&, CLng(TempVars("TimerHandle"))) Then TempVars("TimerHandle") = 0
        GdiplusShutdown lGDIP
        lGDIP = 0
        TempVars("GDIPlusHandle") = Null
        If Not bSharedLoad Then FreeLibrary GetModuleHandle("gdiplus.dll")
    End If
End Sub

'Scheduled ShutDown of GDI+ handle to avoid memory leaks
Private Sub AutoShutDown()
    'Set to 5 seconds for next shutdown
    'That's IMO appropriate for looped routines  - but configure for your own purposes
    If lGDIP <> 0 Then
        TempVars("TimerHandle") = SetTimer(0&, 0&, 5000, AddressOf TimerProc)
    End If
End Sub

'Callback for AutoShutDown
Private Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    'Debug.Print "GDI+ AutoShutDown", idEvent
    If TempVars("TimerHandle") <> 0 Then
        If KillTimer(0&, CLng(TempVars("TimerHandle"))) Then TempVars("TimerHandle") = 0
    End If
    ShutDownGDIP
End Sub

'Load image file with GDIP
'It's equivalent to the method LoadPicture() in OLE-Automation library (stdole2.tlb)
'Allowed format: bmp, gif, jp(e)g, tif, png, wmf, emf, ico
Function LoadPictureGDIP(SFileName As String) As StdPicture
    Dim hBmp As LongPtr
    Dim hPic As LongPtr

    If Not InitGDIP Then Exit Function
    If GdipCreateBitmapFromFile(StrPtr(SFileName), hPic) = 0 Then
        GdipCreateHBITMAPFromBitmap hPic, hBmp, 0&
        If hBmp <> 0 Then
            Set LoadPictureGDIP = BitmapToPicture(hBmp)
            GdipDisposeImage hPic
        End If
    End If

End Function

'Scale picture with GDIP
'A Picture object is commited, also the return value
'Width and Height of generatrix pictures in Width, Height
'bSharpen: TRUE=Thumb is additional sharpened
Function ResampleGDIP(ByVal Image As StdPicture, ByVal Width As Long, ByVal Height As Long, _
                      Optional bSharpen As Boolean = True) As StdPicture
    Dim lRes As Long
    Dim lBitmap As LongPtr

    If Not InitGDIP Then Exit Function
    
    If Image.Type = 1 Then
        lRes = GdipCreateBitmapFromHBITMAP(Image.handle, 0, lBitmap)
    Else
        lRes = GdipCreateBitmapFromHICON(Image.handle, lBitmap)
    End If
    If lRes = 0 Then
        Dim lThumb As LongPtr
        Dim hBitmap As LongPtr

        lRes = GdipGetImageThumbnail(lBitmap, Width, Height, lThumb, 0, 0)
        If lRes = 0 Then
            If Image.Type = 3 Then  'Image-Type 3 is named : Icon!
                'Convert with these GDI+ method :
                lRes = GdipCreateHICONFromBitmap(lThumb, hBitmap)
                Set ResampleGDIP = BitmapToPicture(hBitmap, True)
            Else
                lRes = GdipCreateHBITMAPFromBitmap(lThumb, hBitmap, 0)
                Set ResampleGDIP = BitmapToPicture(hBitmap)
            End If
            
            GdipDisposeImage lThumb
        End If
        GdipDisposeImage lBitmap
    End If

End Function


'Retrieve Width and Height of a pictures in Pixel with GDIP
'Return value as user/defined type TSize (X/Y als Long)
Function GetDimensionsGDIP(ByVal Image As StdPicture) As TSize
    Dim lRes As Long
    Dim lBitmap As LongPtr
    Dim X As LongPtr, Y As LongPtr

    If Not InitGDIP Then Exit Function
    If Image Is Nothing Then Exit Function
    lRes = GdipCreateBitmapFromHBITMAP(Image.handle, 0, lBitmap)
    If lRes = 0 Then
        GdipGetImageHeight lBitmap, Y
        GdipGetImageWidth lBitmap, X
        GetDimensionsGDIP.X = CDbl(X)
        GetDimensionsGDIP.Y = CDbl(Y)
        GdipDisposeImage lBitmap
    End If

End Function

'Save a bitmap as file (with format conversion!)
'image = StdPicture object
'sFile = complete file path
'PicType = pictypeBMP, pictypeGIF, pictypePNG oder pictypeJPG
'Quality: 0...100; (works only with pictypeJPG!)
'Returns TRUE if successful
Function SavePicGDIPlus(ByVal Image As StdPicture, sFile As String, _
                        PicType As PicFileType, Optional Quality As Long = 80) As Boolean
    Dim lBitmap As LongPtr
    Dim TEncoder As GUID
    Dim ret As Long
    Dim TParams As EncoderParameters
    Dim sType As String

    If Not InitGDIP Then Exit Function

    If GdipCreateBitmapFromHBITMAP(Image.handle, 0, lBitmap) = 0 Then
        Select Case PicType
        Case pictypeBMP: sType = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypeGIF: sType = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypePNG: sType = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypeJPG: sType = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
        End Select
        CLSIDFromString StrPtr(sType), TEncoder
        If PicType = pictypeJPG Then
            TParams.Count = 1
            With TParams.Parameter    ' Quality
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .UUID
                .NumberOfValues = 1
                .Type = 4
                .Value = VarPtr(CLng(Quality))
            End With
        Else
            'Different numbers of parameter between GDI+ 1.0 and GDI+ 1.1 on GIFs!!
            If (PicType = pictypeGIF) Then TParams.Count = 1 Else TParams.Count = 0
        End If
        'Save GDIP-Image to file :
        ret = GdipSaveImageToFile(lBitmap, StrPtr(sFile), TEncoder, TParams)
        GdipDisposeImage lBitmap
        DoEvents
        'Function returns True, if generated file actually exists:
        SavePicGDIPlus = (Dir(sFile) <> "")
    End If

End Function

'This procedure is similar to the above (see Parameter), the different is,
'that nothing is stored as a file, but a conversion is executed
'using a OLE-Stream-Object to an Byte-Array .
Function ArrayFromPicture(ByVal Image As Object, PicType As PicFileType, Optional Quality As Long = 80) As Byte()
    Dim lBitmap As LongPtr
    Dim TEncoder As GUID
    Dim ret As Long
    Dim TParams As EncoderParameters
    Dim sType As String
    Dim IStm As IUnknown

    If Not InitGDIP Then Exit Function

    If GdipCreateBitmapFromHBITMAP(Image.handle, 0, lBitmap) = 0 Then
        Select Case PicType    'Choose GDIP-Format-Encoders CLSID:
        Case pictypeBMP: sType = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypeGIF: sType = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypePNG: sType = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypeJPG: sType = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
        End Select
        CLSIDFromString StrPtr(sType), TEncoder

        If PicType = pictypeJPG Then    'If JPG, then set additional parameter
                                        ' to apply quality level
            TParams.Count = 1
            With TParams.Parameter    ' Quality
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .UUID
                .NumberOfValues = 1
                .Type = 4
                .Value = VarPtr(CLng(Quality))
            End With
        Else
            'Different number of parameters between GDI+ 1.0 and GDI+ 1.1 on GIFs!!
            If (PicType = pictypeGIF) Then TParams.Count = 1 Else TParams.Count = 0
        End If

        ret = CreateStreamOnHGlobal(0&, 1, IStm)    'Create stream
        'Save GDIP-Image to stream :
        ret = GdipSaveImageToStream(lBitmap, IStm, TEncoder, TParams)
        If ret = 0 Then
            Dim hMem As LongPtr, LSize As Long, lpMem As Long
            Dim abData() As Byte

            ret = GetHGlobalFromStream(IStm, hMem)    'Get memory handle from stream
            If ret = 0 Then
                LSize = GlobalSize(hMem)
                lpMem = GlobalLock(hMem)   'Get access to memory
                ReDim abData(LSize - 1)    'Arrays dimension
                'Commit memory stack from streams :
                CopyMemory abData(0), ByVal lpMem, LSize
                GlobalUnlock hMem   'Lock memory
                ArrayFromPicture = abData   'Result
            End If

            Set IStm = Nothing  'Clean
        End If

        GdipDisposeImage lBitmap    'Clear GDIP-Image-Memory
    End If

End Function

'Create a picture object from an Access 2007 attachment
'strTable:              Table containing picture file attachments
'strAttachmentField:    Name of the attachment column in the table
'strImage:              Name of the image to search in the attachment records
'? AttachmentToPicture("ribbonimages","imageblob","cloudy.png").Width
Public Function AttachmentToPicture(strTable As String, strAttachmentField As String, strImage As String) As StdPicture
    Dim strSQL As String
    Dim bin() As Byte
    Dim nOffset As Long
    Dim nSize As Long
    
    strSQL = "SELECT " & strTable & "." & strAttachmentField & ".FileData AS data " & _
             "FROM " & strTable & _
             " WHERE " & strTable & "." & strAttachmentField & ".FileName='" & strImage & "'"
    On Error Resume Next
    bin = DBEngine(0)(0).OpenRecordSet(strSQL, dbOpenSnapshot)(0)
    If Err.Number = 0 Then
        Dim bin2() As Byte
        nOffset = bin(0)    'First byte of Field2.FileData identifies offset to the file data block
        nSize = UBound(bin)
        ReDim bin2(nSize - nOffset)
        CopyMemory bin2(0), bin(nOffset), nSize - nOffset   'Copy file into new byte array starting at nOffset
        Set AttachmentToPicture = ArrayToPicture(bin2)
        Erase bin2
        Erase bin
    End If
End Function

'Create an OLE-Picture from Byte-Array PicBin()
Public Function ArrayToPicture(ByRef PicBin() As Byte) As StdPicture
    Dim IStm As IUnknown
    Dim lBitmap As LongPtr
    Dim hBmp As LongPtr
    Dim ret As Long

    If Not InitGDIP Then Exit Function

    ret = CreateStreamOnHGlobal(VarPtr(PicBin(0)), 0, IStm)  'Create stream from memory stack
    If ret = 0 Then    'OK, start GDIP :
        'Convert stream to GDIP-Image :
        ret = GdipLoadImageFromStream(IStm, lBitmap)
        If ret = 0 Then
            'Get Windows-Bitmap from GDIP-Image:
            GdipCreateHBITMAPFromBitmap lBitmap, hBmp, 0&
            If hBmp <> 0 Then
                'Convert bitmap to picture object :
                Set ArrayToPicture = BitmapToPicture(hBmp)
            End If
        End If
        'Clear memory ...
        GdipDisposeImage lBitmap
    End If

End Function

'Help function to get a OLE-Picture from Windows-Bitmap-Handle
'If bIsIcon = TRUE, an Icon-Handle is commited
Function BitmapToPicture(ByVal hBmp As LongPtr, Optional bIsIcon As Boolean = False) As StdPicture
    Dim TPicConv As PICTDESC, uid As GUID

    With TPicConv
        If bIsIcon Then
            .cbSizeOfStruct = 16
            .PicType = 3    'PicType Icon
        Else
            .cbSizeOfStruct = Len(TPicConv)
            .PicType = 1    'PicType Bitmap
        End If
        .hImage = hBmp
    End With

    CLSIDFromString StrPtr(GUID_IPicture), uid
    OleCreatePictureIndirect TPicConv, uid, True, BitmapToPicture

End Function