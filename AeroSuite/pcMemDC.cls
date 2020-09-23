VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcMemDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' cMemDC - flicker free drawing

'--- Raster Operation Codes
Private Const DSna                      As Long = &H220326
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Long
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type RGBTRIPLE
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
End Type
Private Type PALETTEENTRY
    peRed               As Byte
    peGreen             As Byte
    peBlue              As Byte
    peFlags             As Byte
End Type

Private Type LOGPALETTE
    palVersion          As Integer
    palNumEntries       As Integer
    palPalEntry(255)    As PALETTEENTRY
End Type
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
           lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long

Private m_hDC As Long
Private m_hBmp As Long
Private m_hBmpOld As Long
Private m_lWidth As Long
Private m_lHeight As Long

Public Property Get Width() As Long
   Width = m_lWidth
End Property
Public Property Let Width(ByVal Value As Long)
Dim lJunk As Long
   If (Value > m_lWidth) Then
      m_lWidth = Value
      pCreate m_lWidth, m_lHeight
   End If
End Property
Public Property Get Height() As Long
   Height = m_lHeight
End Property
Public Property Let Height(ByVal Value As Long)
Dim lJunk As Long
   If (Value > m_lHeight) Then
      m_lHeight = Value
      pCreate m_lWidth, m_lHeight
   End If
End Property

Public Property Get hDC() As Long
   hDC = m_hDC
End Property
Public Sub StretchDraw( _
      ByVal hDC As Long, _
      Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0, _
      Optional ByVal WidthDst As Long = 0, Optional ByVal HeightDst As Long = 0, _
      Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
      Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0 _
   )
   If WidthSrc <= 0 Then WidthSrc = m_lWidth
   If HeightSrc <= 0 Then HeightSrc = m_lHeight
   StretchBlt hDC, xDst, yDst, WidthDst, HeightDst, m_hDC, xSrc, ySrc, WidthSrc, HeightSrc, vbSrcCopy
End Sub
Public Sub Draw( _
      ByVal hDC As Long, _
      Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
      Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, _
      Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0, _
      Optional ByVal UseMask As Boolean _
   )
  If WidthSrc <= 0 Then WidthSrc = m_lWidth
  If HeightSrc <= 0 Then HeightSrc = m_lHeight
  If UseMask Then
    pvTransBlt hDC, xDst, yDst, WidthSrc, HeightSrc, m_hDC, xSrc, ySrc
  Else
    BitBlt hDC, xDst, yDst, WidthSrc, HeightSrc, m_hDC, xSrc, ySrc, vbSrcCopy
  End If
End Sub
Public Sub CreateFromPicture(sPic As IPicture)
Dim tB As BITMAP
Dim lhDCC As Long, lhDC As Long
Dim lhBmpOld As Long
   GetObjectAPI sPic.Handle, Len(tB), tB
   Width = tB.bmWidth
   Height = tB.bmHeight
   lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   lhDC = CreateCompatibleDC(lhDCC)
   lhBmpOld = SelectObject(lhDC, sPic.Handle)
   BitBlt m_hDC, 0, 0, tB.bmWidth, tB.bmHeight, lhDC, 0, 0, vbSrcCopy
   SelectObject lhDC, lhBmpOld
   DeleteDC lhDC
   DeleteDC lhDCC
End Sub
Private Sub pCreate(ByVal Width As Long, ByVal Height As Long)
Dim lhDCC As Long
   pDestroy
   lhDCC = CreateDC("DISPLAY", "", "", ByVal 0&)
   If Not (lhDCC = 0) Then
      m_hDC = CreateCompatibleDC(lhDCC)
      If Not (m_hDC = 0) Then
         m_hBmp = CreateCompatibleBitmap(lhDCC, Width, Height)
         If Not (m_hBmp = 0) Then
            m_hBmpOld = SelectObject(m_hDC, m_hBmp)
            If Not (m_hBmpOld = 0) Then
               m_lWidth = Width
               m_lHeight = Height
               DeleteDC lhDCC
               Exit Sub
            End If
         End If
      End If
      DeleteDC lhDCC
      pDestroy
   End If
End Sub
Private Sub pDestroy()
   If Not m_hBmpOld = 0 Then
      SelectObject m_hDC, m_hBmpOld
      m_hBmpOld = 0
   End If
   If Not m_hBmp = 0 Then
      DeleteObject m_hBmp
      m_hBmp = 0
   End If
   If Not m_hDC = 0 Then
      DeleteDC m_hDC
      m_hDC = 0
   End If
   m_lWidth = 0
   m_lHeight = 0
End Sub

Private Sub Class_Terminate()
   pDestroy
End Sub

Private Sub pvTransBlt( _
            ByVal hdcDest As Long, _
            ByVal xDest As Long, _
            ByVal yDest As Long, _
            ByVal nWidth As Long, _
            ByVal nHeight As Long, _
            ByVal hdcSrc As Long, _
            Optional ByVal xSrc As Long = 0, _
            Optional ByVal ySrc As Long = 0, _
            Optional ByVal clrMask As OLE_COLOR = vbMagenta, _
            Optional ByVal hPal As Long = 0)
    Dim hdcMask             As Long ' hDC of the created mask image
    Dim hdcColor            As Long ' hDC of the created color image
    Dim hbmMask             As Long ' Bitmap handle to the mask image
    Dim hbmColor            As Long ' Bitmap handle to the color image
    Dim hbmColorOld         As Long
    Dim hbmMaskOld          As Long
    Dim hpalOld             As Long
    Dim hdcScreen           As Long
    Dim hdcScnBuffer        As Long ' Buffer to do all work on
    Dim hbmScnBuffer        As Long
    Dim hbmScnBufferOld     As Long
    Dim hPalBufferOld       As Long
    Dim lMaskColor          As Long
    Dim hpalHalftone        As Long

    hdcScreen = GetDC(0&)
    ' Validate palette
    If hPal = 0 Then
        hpalHalftone = CreateHalftonePalette(hdcScreen)
        hPal = hpalHalftone
    End If
    OleTranslateColor clrMask, hPal, lMaskColor
    lMaskColor = lMaskColor And &HFFFFFF
    ' Create a color bitmap to server as a copy of the destination
    ' Do all work on this bitmap and then copy it back over the destination
    ' when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
    ' Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hdcScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    ' Copy the destination to the screen buffer
    BitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcDest, xDest, yDest, vbSrcCopy
    ' Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    ' hdcSrc, because this will create a DIB section if the original bitmap
    ' is a DIB section)
    hbmColor = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
    ' Now create a monochrome bitmap for the mask
    hbmMask = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
    ' First, blt the source bitmap onto the cover.  We do this first
    ' and then use it instead of the source bitmap
    ' because the source bitmap may be
    ' a DIB section, which behaves differently than a bitmap.
    ' (Specifically, copying from a DIB section to a monochrome bitmap
    ' does a nearest-color selection rather than painting based on the
    ' backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hdcScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hpalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    ' In case hdcSrc contains a monochrome bitmap, we must set the destination
    ' foreground/background colors according to those currently set in hdcSrc
    ' (because Windows will associate these colors with the two monochrome colors)
    Call SetBkColor(hdcColor, GetBkColor(hdcSrc))
    Call SetTextColor(hdcColor, GetTextColor(hdcSrc))
    Call BitBlt(hdcColor, 0, 0, nWidth, nHeight, hdcSrc, xSrc, ySrc, vbSrcCopy)
    ' Paint the mask.  What we want is white at the transparent color
    ' from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hdcScreen)
    hbmMaskOld = SelectObject(hdcMask, hbmMask)
    ' When BitBlt'ing from color to monochrome, Windows sets to 1
    ' all pixels that match the background color of the source DC.  All
    ' other bits are set to 0.
    Call SetBkColor(hdcColor, lMaskColor)
    Call SetTextColor(hdcColor, vbWhite)
    Call BitBlt(hdcMask, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcCopy)
    ' Paint the rest of the cover bitmap.
    '
    ' What we want here is black at the transparent color, and
    ' the original colors everywhere else.  To do this, we first
    ' paint the original onto the cover (which we already did), then we
    ' AND the inverse of the mask onto that using the DSna ternary raster
    ' operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    ' Operation Codes", "Ternary Raster Operations", or search in MSDN
    ' for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    ' When BitBlt'ing from monochrome to color, Windows transforms all white
    ' bits (1) to the background color of the destination hDC.  All black (0)
    ' bits are transformed to the foreground color.
    Call SetTextColor(hdcColor, vbBlack)
    Call SetBkColor(hdcColor, vbWhite)
    Call BitBlt(hdcColor, 0, 0, nWidth, nHeight, hdcMask, 0, 0, DSna)
    ' Paint the Mask to the Screen buffer
    Call BitBlt(hdcScnBuffer, 0, 0, nWidth, nHeight, hdcMask, 0, 0, vbSrcAnd)
    ' Paint the Color to the Screen buffer
    Call BitBlt(hdcScnBuffer, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcPaint)
    ' Copy the screen buffer to the screen
    Call BitBlt(hdcDest, xDest, yDest, nWidth, nHeight, hdcScnBuffer, 0, 0, vbSrcCopy)
    ' All done!
    Call DeleteObject(SelectObject(hdcColor, hbmColorOld))
    Call SelectPalette(hdcColor, hpalOld, True)
    Call RealizePalette(hdcColor)
    Call DeleteDC(hdcColor)
    Call DeleteObject(SelectObject(hdcScnBuffer, hbmScnBufferOld))
    Call SelectPalette(hdcScnBuffer, hPalBufferOld, 0)
    Call RealizePalette(hdcScnBuffer)
    Call DeleteDC(hdcScnBuffer)
    Call DeleteObject(SelectObject(hdcMask, hbmMaskOld))
    Call DeleteDC(hdcMask)
    Call ReleaseDC(0&, hdcScreen)
    If hpalHalftone <> 0 Then
        Call DeleteObject(hpalHalftone)
    End If
End Sub


