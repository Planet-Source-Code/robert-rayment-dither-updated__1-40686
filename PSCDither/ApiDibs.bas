Attribute VB_Name = "ApiDibs"
'ApiDibs.bas


Option Base 1  ' Arrays base 1
DefLng A-W     ' All variable Long
DefSng X-Z     ' unless singles
               ' unless otherwise defined
               

' -----------------------------------------------------------
'  This required instead of Screen.Height & Width for resizing

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXSCREEN = 0 'X Size of screen
Public Const SM_CYSCREEN = 1 'Y Size of Screen

' -----------------------------------------------------------
Public Declare Function GetPixel Lib "gdi32" _
(ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

' -----------------------------------------------------------
' APIs for getting DIB bits to PicMem

Public Declare Function GetDIBits Lib "gdi32" _
(ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" _
(ByVal hdc As Long) As Long

Public Declare Function SelectObject Lib "gdi32" _
(ByVal hdc As Long, ByVal hObject As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" _
(ByVal hdc As Long) As Long

'------------------------------------------------------------------------------

'To fill BITMAP structure
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal Lenbmp As Long, dimbmp As Any) As Long

Public Type BITMAP
   bmType As Long              ' Type of bitmap
   bmWidth As Long             ' Pixel width
   bmHeight As Long            ' Pixel height
   bmWidthBytes As Long        ' Byte width = 3 x Pixel width
   bmPlanes As Integer         ' Color depth of bitmap
   bmBitsPixel As Integer      ' Bits per pixel, must be 16 or 24
   bmBits As Long              ' This is the pointer to the bitmap data  !!!
End Type
Public bmp As BITMAP

'------------------------------------------------------------------------------

' Structures for StretchDIBits
Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
   Colors(0 To 1) As RGBQUAD
End Type
Public bm As BITMAPINFO

' For transferring drawing in an array to Form or PicBox 7 printing
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal x As Long, ByVal y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcXOffset As Long, ByVal SrcYOffset As Long, _
ByVal PICWW As Long, ByVal PICHH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long
'wUsage is one of:-
Public Const DIB_PAL_COLORS = 1 '  uses system
Public Const DIB_RGB_COLORS = 0 '  uses RGBQUAD
'dwRop is vbSrcCopy
'------------------------------------------------------------------------------
' To invert picture box
Public Declare Function SetRect Lib "user32" (lpRect As RECT, _
ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function InvertRect Lib "user32" _
(ByVal hdc As Long, lpRect As RECT) As Long

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public IR As RECT
'------------------------------------------------------------------------------

Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" _
(ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Public Declare Function SetCursor Lib "user32" _
(ByVal hCursor As Long) As Long
'Use  SetCursor LoadCursor(0, IDC_XXXXX)

Public Const IDC_ARROW = 32512&
Public Const IDC_HAND = 32649&
Public Const IDC_LRTRI = 32653&     ' Left-Right triangle (Win98)

'------------------------------------------------------------------------------
'Copy one array to another of same number of bytes

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)


'Public pic1mem() As Byte  ' Holds loaded picture
'Public pic2mem() As Byte  ' Holds dithered picture

Public Sub FillBMPStruc(ByVal bwidth, ByVal bheight)
  
  With bm.bmiH
   .biSize = 40
   .biwidth = bwidth
   .biheight = bheight
   .biPlanes = 1
   .biBitCount = 32     ' always 32 in this prog
   .biCompression = 0
   
   BScanLine = (((bwidth * .biBitCount) + 31) \ 32) * 4
   ' Ensure expansion to 4B boundary
   BScanLine = (Int((BScanLine + 3) \ 4)) * 4

   .biSizeImage = BScanLine * Abs(.biheight)
   .biXPelsPerMeter = 0
   .biYPelsPerMeter = 0
   .biClrUsed = 0
   .biClrImportant = 0
 End With

End Sub

Public Sub GETDIBS(ByVal PICIM As Long)


' PICIM is picbox.Image - handle to picbox memory
' from which pixels will be extracted and
' stored in pic1mem()

On Error GoTo DIBError

'Get info on picture loaded into PIC
GetObjectAPI PICIM, Len(bmp), bmp
' which fills:-
'Public Type BITMAP
'   bmType As Long              ' Type of bitmap
'   bmWidth As Long             ' Pixel width
'   bmHeight As Long            ' Pixel height
'   bmWidthBytes As Long        ' Byte width = 3 x Pixel width
'   bmPlanes As Integer         ' Color depth of bitmap
'   bmBitsPixel As Integer      ' Bits per pixel, must be 16 or 24
'   bmBits As Long              ' This is the pointer to the bitmap data  !!!
'End Type
'Public bmp As BITMAP


NewDC = CreateCompatibleDC(0&)
OldH = SelectObject(NewDC, PICIM)

FillBMPStruc bmp.bmWidth, bmp.bmHeight
' which fills:-
' Structures for StretchDIBits
'Public Type BITMAPINFOHEADER ' 40 bytes
'   biSize As Long
'   biwidth As Long
'   biheight As Long
'   biPlanes As Integer
'   biBitCount As Integer
'   biCompression As Long
'   biSizeImage As Long
'   biXPelsPerMeter As Long
'   biYPelsPerMeter As Long
'   biClrUsed As Long
'   biClrImportant As Long
'End Type

' Set PicMem to receive color bytes or indexes or bits
Culb = 4
PICH = bmp.bmHeight
PICW = bmp.bmWidth
ReDim pic1mem(Culb, PICW, PICH)
ReDim pic2mem(Culb, PICW, PICH)

' Load color bytes to pic1mem
ret = GetDIBits(NewDC, PICIM, 0, PICH, pic1mem(1, 1, 1), bm, 1)

' Clear mem
SelectObject NewDC, OldH
DeleteDC NewDC

picmemPtr = VarPtr(pic1mem(1, 1, 1))
picmemSize = BScanLine * PICH

Exit Sub
'==========
DIBError:
  MsgBox "DIB Error in GETDIBS", , "Dither"
  DoEvents
  Unload Form1
  End
End Sub

