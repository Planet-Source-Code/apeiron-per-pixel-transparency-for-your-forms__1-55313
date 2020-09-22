VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   2415
      Left            =   -240
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This is an example of how to load a alpha transparent png with GDI+
' and get windows to display it using the UpdateLayeredWindows call.
'  I haven't seen anything this simple like this on the net.
'  Really neat effect.  This could be used to msake the coolest splash screens.
'  I say splash screens because all the drawing is taken over by windows so none of the child
' controls are shown, though they
' can still respond to mouse/keyboard events.  Try it out, drag the top of the circle
' and you can use the titlebar or click in a non-transparent area and the button will react.
'  With some careful planning you could use this not just for splash screens but a really interesting UI.
'  Have fun with it. (Comments/Votes welcome)

' I tried to remove most everything but the basic loading of a 32 Bpp .png image with alpha transparency
' and setting it s display with the UpdateLayred windows calls.  It's easier for me to see an example when all the
' (most of) the extraneous stuff is taken out.

' For additional info check out these two submissions, both very good projects.

'  Thanks to the submissions at Using 32pp GDI Bitmaps and UpdateLayeredWindow at
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=52629&lngWId=1 for an example of per pixel alpha using 24 bpp bitmaps.

' and Use GDI+ at http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
' for a really nice, extensive GDI plus api declarations and examples.

' Also close form by right clicking and selecting close rather than shutting down
' with the button in the IDE.  Otherwise there is sometimes a memory problem that shuts down VB.
' Probably something simple I'm just not seeing.

Private Const ULW_OPAQUE = &H4
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER = &H0
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function AlphaBlend Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal lnYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal bf As Long) As Boolean
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Dim mDC As Long  ' Memory hDC
Dim mainBitmap As Long ' Memory Bitmap
Dim blendFunc32bpp As BLENDFUNCTION
Dim token As Long ' Needed to close GDI+
Dim oldBitmap As Long

Private Sub Command1_Click()
  MsgBox "Button has been clicked even though it's not being drawn!"
  MakeTrans (App.Path & "\test2.png")
End Sub

Private Sub Form_Initialize()
   ' Start up GDI+
   Dim GpInput As GdiplusStartupInput
   GpInput.GdiplusVersion = 1
   If GdiplusStartup(token, GpInput) <> 0 Then
     MsgBox "Error loading GDI+!", vbCritical
     Unload Me
   End If
   MakeTrans (App.Path & "\test.png")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Cleanup everything
    Call GdiplusShutdown(token)
    SelectObject mDC, oldBitmap
    DeleteObject mainBitmap
    DeleteObject oldBitmap
    DeleteDC mDC
End Sub

Private Function MakeTrans(pngPath As String) As Boolean
   Dim tempBI As BITMAPINFO
   Dim tempBlend As BLENDFUNCTION      ' Used to specify what kind of blend we want to perform
   Dim lngHeight As Long, lngWidth As Long
   Dim curWinLong As Long
   Dim img As Long
   Dim graphics As Long
   Dim winSize As Size
   Dim srcPoint As POINTAPI
   
   With tempBI.bmiHeader
      .biSize = Len(tempBI.bmiHeader)
      .biBitCount = 32    ' Each pixel is 32 bit's wide
      .biHeight = Me.ScaleHeight  ' Height of the form
      .biWidth = Me.ScaleWidth    ' Width of the form
      .biPlanes = 1   ' Always set to 1
      .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
   End With
   mDC = CreateCompatibleDC(Me.hdc)
   mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
   oldBitmap = SelectObject(mDC, mainBitmap)   ' Select the new bitmap, track the old that was selected
    
   ' GDI Initializations
   Call GdipCreateFromHDC(mDC, graphics)
   Call GdipLoadImageFromFile(StrConv(pngPath, vbUnicode), img)  ' Load Png
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)
   Call GdipDrawImageRect(graphics, img, 0, 0, lngWidth, lngHeight)

   ' Change windows extended style to be used by updatelayeredwindow
   curWinLong = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    ' Accidently did This line below which flipped entire form, it's neat so I left it in
    ' Comment out the line above and uncomment line below.
    'curWinLong = GetWindowLong(Me.hwnd, GWL_STYLE)
   SetWindowLong Me.hwnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED
   
    ' Make the window a top-most window so we can always see the cool stuff
   SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
   
   ' Needed for updateLayeredWindow call
   srcPoint.x = 0
   srcPoint.y = 0
   winSize.cx = Me.ScaleWidth
   winSize.cy = Me.ScaleHeight
    
   With blendFunc32bpp
      .AlphaFormat = AC_SRC_ALPHA ' 32 bit
      .BlendFlags = 0
      .BlendOp = AC_SRC_OVER
      .SourceConstantAlpha = 255
   End With
    
   Call GdipDisposeImage(img)
   Call GdipDeleteGraphics(graphics)
   Call UpdateLayeredWindow(Me.hwnd, Me.hdc, ByVal 0&, winSize, mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
End Function
