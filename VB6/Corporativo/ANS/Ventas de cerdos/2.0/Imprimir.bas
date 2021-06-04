Attribute VB_Name = "Module1"

Option Explicit
Private Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type
Private Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   palPalEntry(255) As PALETTEENTRY
End Type
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Private Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type
Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetWindowDC Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function ReleaseDC Lib "USER32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Public Function CaptureClient(frmSrc As Form) As Picture
On Error Resume Next
Set CaptureClient = CaptureWindow(frmSrc.hWnd, True, 0, 0, frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
End Function
Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
On Error Resume Next
Dim hDCMemory As Long
Dim hBmp As Long
Dim hBmpPrev As Long
Dim r As Long
Dim hDCSrc As Long
Dim hPal As Long
Dim hPalPrev As Long
Dim RasterCapsScrn As Long
Dim HasPaletteScrn As Long
Dim PaletteSizeScrn As Long
Dim LogPal As LOGPALETTE
If Client Then
    hDCSrc = GetDC(hWndSrc)
Else
    hDCSrc = GetWindowDC(hWndSrc)
End If
hDCMemory = CreateCompatibleDC(hDCSrc)
hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
hBmpPrev = SelectObject(hDCMemory, hBmp)
RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
HasPaletteScrn = RasterCapsScrn And RC_PALETTE
PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    LogPal.palVersion = &H300
    LogPal.palNumEntries = 256
    r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
    hPal = CreatePalette(LogPal)
    hPalPrev = SelectPalette(hDCMemory, hPal, 0)
    r = RealizePalette(hDCMemory)
End If
r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
hBmp = SelectObject(hDCMemory, hBmpPrev)
If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    hPal = SelectPalette(hDCMemory, hPalPrev, 0)
End If
r = DeleteDC(hDCMemory)
r = ReleaseDC(hWndSrc, hDCSrc)
Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
On Error Resume Next
Dim r As Long
Dim Pic As PicBmp
Dim IPic As IPicture
Dim IID_IDispatch As GUID
With IID_IDispatch
   .Data1 = &H20400
   .Data4(0) = &HC0
   .Data4(7) = &H46
End With
With Pic
   .Size = Len(Pic)
   .Type = vbPicTypeBitmap
   .hBmp = hBmp
   .hPal = hPal
End With
r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
Set CreateBitmapPicture = IPic
End Function


