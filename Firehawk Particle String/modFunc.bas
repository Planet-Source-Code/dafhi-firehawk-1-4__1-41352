Attribute VB_Name = "modFunc"
Option Explicit

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(1) As SAFEARRAYBOUND
End Type

Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLbound As Long
End Type

Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Type PicBmp
    Size As Long
    Type As PictureTypeConstants
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Type RGBQUAD
    Red As Byte
    Green As Byte
    Blue As Byte
    Reserved As Byte
End Type

Type BITMAPINFOHEADER
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

Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type PointAPI
 X As Long
 Y As Long
End Type

Private Type DimsAPI
 Width As Long
 Height As Long
End Type

Public Type AnimSurf1D
 Dib() As RGBQUAD
 LDib() As Long
 TopRight As PointAPI
 Dims As DimsAPI
 halfW As Single
 halfH As Single
 TotalPixels As Long
 UBSurf As Long
 SafeAry1D As SAFEARRAY1D
 SafeAry1D_L As SAFEARRAY1D
 EraseDib() As Long
 LBotLeftErase() As Long
 LTopLeftErase() As Long
 LEraseWidth() As Long
 EraseSpriteCount As Long
End Type

Public tSA As SAFEARRAY1D
Public tSA2 As SAFEARRAY1D

Public BM As BITMAP
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy&)

Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject&, ByVal nCount&, lpObject As Any) As Long
Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC&, pBitmapInfo As BITMAPINFO, ByVal un&, lplpVoid&, ByVal handle&, ByVal dw&) As Long

Declare Function RedrawWindow Lib "user32" (ByVal hwnd&, lprcUpdate As RECT, ByVal hrgnUpdate&, ByVal fuRedraw&) As Long

Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle&, IPic As IPicture) As Long

Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

Function CreatePicture(ByVal nWidth&, ByVal nHeight&, ByVal BitDepth&) As Picture
Dim Pic As PicBmp, IID_IDispatch As GUID
Dim BMI As BITMAPINFO
With BMI.bmiHeader
.biSize = Len(BMI.bmiHeader)
.biWidth = nWidth
.biHeight = nHeight
.biPlanes = 1
.biBitCount = BitDepth
End With
Pic.hBmp = CreateDIBSection(0, BMI, 0, 0, 0, 0)
IID_IDispatch.Data1 = &H20400: IID_IDispatch.Data4(0) = &HC0: IID_IDispatch.Data4(7) = &H46
Pic.Size = Len(Pic)
Pic.Type = vbPicTypeBitmap
OleCreatePictureIndirect Pic, IID_IDispatch, 1, CreatePicture
If CreatePicture = 0 Then Set CreatePicture = Nothing
End Function

Function GetPicture(ByVal Pic&, outAry() As Byte) As Boolean
GetObjectAPI Pic, Len(BM), BM
ReDim outAry(BM.bmWidthBytes * BM.bmHeight - 1)
CopyMemory outAry(0), ByVal BM.bmBits, BM.bmWidthBytes * BM.bmHeight
GetPicture = True
End Function

Function SetPicture(ByVal Pic&, inAry() As Byte) As Boolean
GetObjectAPI Pic, Len(BM), BM
If LBound(inAry) <> 0 Or UBound(inAry) <> BM.bmWidthBytes * BM.bmHeight - 1 Then Exit Function
CopyMemory ByVal BM.bmBits, inAry(0), BM.bmWidthBytes * BM.bmHeight
SetPicture = True
End Function

Public Sub InitAnimSurf1D(Obj As Object, Surf As AnimSurf1D)
 
 Surf.Dims.Width = Obj.ScaleWidth
 Surf.Dims.Height = Obj.ScaleHeight
 Surf.TopRight.X = Surf.Dims.Width - 1
 Surf.TopRight.Y = Surf.Dims.Height - 1
 
 Surf.halfW = Surf.Dims.Width / 2&
 Surf.halfH = Surf.Dims.Height / 2&
 
 Surf.TotalPixels = Surf.Dims.Width * Surf.Dims.Height
 
 If Surf.TotalPixels > 0 Then
 
  CleanUp Surf
 
  Surf.UBSurf = Surf.TotalPixels - 1
  
  'Allocate memory to the Refresh buffer
  Obj.Picture = CreatePicture(Surf.Dims.Width, Surf.Dims.Height, 32)
  GetObjectAPI Obj.Picture, Len(BM), BM
  With Surf.SafeAry1D
  .cbElements = 4
  .cDims = 1
  .lLbound = 0
  .cElements = BM.bmHeight * BM.bmWidth
  .pvData = BM.bmBits
  End With
  CopyMemory ByVal VarPtrArray(Surf.Dib), VarPtr(Surf.SafeAry1D), 4
 
  With Surf.SafeAry1D_L
  .cbElements = 4
  .cDims = 1
  .lLbound = 0
  .cElements = BM.bmHeight * BM.bmWidth
  .pvData = BM.bmBits
  ReDim Surf.EraseDib(.cElements)
  End With
  CopyMemory ByVal VarPtrArray(Surf.LDib), VarPtr(Surf.SafeAry1D_L), 4
 
 End If 'Surf.TotalPixels > 0
 
End Sub

Public Sub CleanUp(Surf As AnimSurf1D)
 CopyMemory ByVal VarPtrArray(Surf.Dib), 0&, 4
 CopyMemory ByVal VarPtrArray(Surf.LDib), 0&, 4
 Erase Surf.EraseDib
 Erase Surf.LBotLeftErase
 Erase Surf.LTopLeftErase
 Erase Surf.LEraseWidth
 Surf.EraseSpriteCount = 0
End Sub

