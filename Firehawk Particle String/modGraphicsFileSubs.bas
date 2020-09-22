Attribute VB_Name = "modGraphicsFileSubs"
Option Explicit

Dim AlphaIndex&
Dim N&
Dim N1&

Dim sR!
Dim sG!
Dim sB!

Dim BMPadBytes&
Dim SpriteLastBlueRight&
Dim BMPFileHeader As BITMAPFILEHEADER   'Holds the file header
Dim BMPInfoHeader As BITMAPINFOHEADER   'Holds the info header
Dim BMPData() As Byte                   'Holds the pixel data

'Bitmap file format structures
Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Type DafhiSprite
 BGRAPixels() As Byte
 LngPixels() As Long
 precisionBGR() As Single
 Wid As Long
 Hei As Long
 wDiv2 As Long
 hDiv2 As Long
 SHigh As Long
 SRight As Long
 WidexHigh As Long
 SWidthBytes As Long
 SBlueRight As Long
 SLastBlue As Long
 SLast As Long
End Type

Public Tile1 As DafhiSprite

Dim ClipWidthLeft&
Dim CLipWidthTop&
Dim CLipWidthBot&
Dim SpriteClipLeftBytes&
Dim DrawLeftBase&
Dim AddDrawWidth&
Dim SpriteWidth&
Dim SpriteHeight&
Dim DrawRight&
Dim AddDrawWidthBytes&
Dim SpriteX&
Dim SpriteY&
Dim DrawY&
Dim DrawBot&
Dim DrawTop&
Dim SpriteArrayDepth&
Dim SpriteWidthBytes&
Dim DrawWidthBytes&
Dim Left_&
Dim Top_&
Dim Right_&
Dim Bot_&
Dim DrawWidth_&

Public Sub SetSpriteFromFile(DS1 As DafhiSprite, strFileName$, Optional intensity! = 1)
Dim XLngSSFF&
Dim YLngSSFF&
Dim TopLeft&
 
 Open (App.Path & "\" & strFileName) For Binary As #1
   Get #1, 1, BMPFileHeader
   Get #1, , BMPInfoHeader
   With BMPInfoHeader
    N = .biWidth * 3
    BMPadBytes = ((N + 3) And &HFFFFFFFC) - N
    SpriteWidth& = .biWidth
    SpriteHeight& = .biHeight
    ReDim BMPData(.biHeight * (BMPadBytes + .biWidth * .biBitCount / 8))
   End With
   Get #1, , BMPData
 Close #1
   
 EnumDafhiSprite DS1
  
 ReDim DS1.BGRAPixels(DS1.SLast)
 ReDim DS1.precisionBGR(DS1.SLast)
 'ReDim DS1.AlphaElem(SpriteWidth * SpriteHeight - 1)
 DrawRight = DS1.SRight * 3
 N1 = 0
 N& = SpriteWidth * 3 + BMPadBytes
 TopLeft = N * (SpriteHeight - 1)
 'AlphaIndex = 0
 For YLngSSFF& = 0 To TopLeft Step N
  AddDrawWidthBytes& = YLngSSFF& + DrawRight&
  For XLngSSFF& = YLngSSFF& To AddDrawWidthBytes& Step 3&
   sB = BMPData(XLngSSFF&)
   sG = BMPData(XLngSSFF& + 1&)
   sR = BMPData(XLngSSFF& + 2&)
   DS1.BGRAPixels(N1&) = sB!
   DS1.BGRAPixels(N1& + 1&) = sG!
   DS1.BGRAPixels(N1& + 2&) = sR!
   DS1.precisionBGR(N1&) = sB
   DS1.precisionBGR(N1& + 1&) = sG
   DS1.precisionBGR(N1& + 2&) = sR
   'DS1.AlphaElem(AlphaIndex) = intensity
   N1& = N1& + 4&
   'AlphaIndex = AlphaIndex + 1
  Next XLngSSFF&
 Next YLngSSFF&
 
End Sub
Private Sub EnumDafhiSprite(DS1 As DafhiSprite)
With DS1
 .Wid = SpriteWidth&
 .wDiv2 = .Wid / 2&
 .Hei = SpriteHeight&
 .hDiv2 = .Hei / 2& '+ SH
 .SHigh = .Hei - 1&
 .SRight = .Wid - 1&
 .WidexHigh = .Wid * .Hei
 .SWidthBytes = SpriteWidth * 4&
 .SBlueRight = .SWidthBytes& - 4&
 .SLast = SpriteHeight& * .SWidthBytes& - 1&
 .SLastBlue = .SLast - 2&
End With
End Sub

Public Sub BLitAsTile(DS1 As DafhiSprite, ByVal x!, ByVal y!, Optional blnErasing As Boolean = False, Optional CopyToBuf As Boolean = False)
Dim LWidthRemaining&
Dim DrawDisplacement&
Dim DrawX&

 SpriteX = x
 SpriteY = y
 
  dsc2 DS1
  
  For DrawY& = DrawBot& To DrawTop& Step StepX&
   'CopyMemory bDib(DrawY), DS1.BGRAPixels(SpriteArrayDepth), DrawWidthBytes
   DrawDisplacement = SpriteArrayDepth
   DrawRight = DrawY + AddDrawWidthBytes
   For DrawX = DrawY To DrawRight Step 4&
   bDib(DrawX) = DS1.BGRAPixels(DrawDisplacement)
   bDib(DrawX + 1&) = DS1.BGRAPixels(DrawDisplacement + 1&)
   bDib(DrawX + 2&) = DS1.BGRAPixels(DrawDisplacement + 2&)
   DrawDisplacement = DrawDisplacement + 4&
   Next
   SpriteArrayDepth& = SpriteArrayDepth& + SpriteWidthBytes
  Next DrawY
 
    
End Sub
Private Sub dsc2(DS1 As DafhiSprite)
 
 With DS1
  'independent
  Left_& = SpriteX& - .wDiv2
  Top_& = SH - SpriteY - .hDiv2
 
  'dependent
  Right_& = Left_& + .SRight
  Bot_& = Top_& - .SHigh
  
  SpriteWidthBytes& = .SWidthBytes
  SpriteWidth& = .Wid
  SpriteHeight& = .Hei
 End With
 
 CLipWidthTop& = Top_& - Ymax&
  
 If Bot_ < 0& Then
  CLipWidthBot& = -Bot_&
  DrawBot& = 0&
 Else
  CLipWidthBot = 0&
  DrawBot& = Bot_& * StepX&
 End If
 
 If Left_& < SW& Then
  If Left_& < 0& Then
   ClipWidthLeft& = -Left_&
   DrawLeftBase& = 0&
  Else
   ClipWidthLeft = 0&
   DrawLeftBase = 4& * Left_
  End If
  If Right_& > Xmax Then
   DrawWidth_& = SW& - Left_& - ClipWidthLeft&
  ElseIf Right_& > -1& Then
   DrawWidth_& = SpriteWidth& - ClipWidthLeft&
  Else
   DrawWidth_& = 0&
  End If
 Else
  DrawWidth_& = 0&
 End If
 
 SpriteClipLeftBytes& = ClipWidthLeft& * 4&
 
 DrawWidthBytes& = DrawWidth_& * 4&
 AddDrawWidthBytes& = DrawWidthBytes& - 4&
 
 If CLipWidthTop& < 0& Then
  DrawTop& = Top_& * StepX&
 Else
  DrawTop = ViewPort_TopLeft
 End If
 
 DrawBot& = DrawBot& + DrawLeftBase&
 DrawTop& = DrawTop& + DrawLeftBase&
 
 SpriteArrayDepth& = CLipWidthBot& * SpriteWidthBytes& + SpriteClipLeftBytes&

End Sub

