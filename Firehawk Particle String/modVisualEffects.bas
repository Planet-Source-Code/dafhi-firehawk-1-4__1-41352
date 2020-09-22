Attribute VB_Name = "modVisualEffects"
Option Explicit

Public iR!
Public iG!
Public iB!

Dim N&

'color changing vars
Public maximu!
Public minimu!
Dim iSubt!
Dim bytMaxMin_diff As Byte
Dim SHF_Intensity!
Public intensity!

Dim i6i!
Dim i1!
Dim i2!
Dim i3!
Dim i4!

Dim SpriteX&
Dim SpriteY&

Dim SpriteArrayDepth&

Dim ClipWidthLeft&
Dim CLipWidthBot&
Dim CLipWidthTop&
Dim DrawWidth&

Dim initial_offset_X!
Dim initial_offset_Y!
Public delta_x!
Public delta_y!
Dim delta_ySq!
Dim half!
Dim Rounded&
Dim pp5!
Dim mult2!
Dim add2!
Dim baseleft!
Dim Bright&

Dim right_side_scale!
Dim left_side_scale!
Dim top_edge_scale!
Dim bottom_edge_scale!

Dim Left_&
Dim Right_&
Dim Top_&
Dim Bot_&

Public Const ParticlesPerGroupUB As Long = 499

Private Type ParticlePrecisionDims
 diameter As Single
 radius As Single
End Type

Private Type ParticleSystemVariables
 intens As Single
 def As Single
End Type

Public Type ParticleSprite
 Dims As ParticlePrecisionDims
 px As Single
 py As Single
 ix As Single
 iy As Single
 def As Single
 sRed As Single
 sGrn As Single
 sBlu As Single
 iLum As Single
 iFrame As Single
 sFrame As Single
 maxFrame As Single
End Type

Public Type ParticleStruct
 chIntensity As Byte
 definition As Single
 centerX As Single
 centerY As Single
 cdx As Single
 cdy As Single
 sRed As Single
 sGrn As Single
 sBlu As Single
 sBright As Single
 iBright As Single
 maxRadius As Single
 IProcess As ParticleSystemVariables
End Type

Public Type ParticleGroup
 chRed As Byte
 chGreen As Byte
 chBlue As Byte
 emitRate As Single
 explosiveness As Single
 IPS As ParticleStruct
 bErasing As Boolean
 bCopyToEraseBuf As Boolean
End Type

Dim WhichAry() As SAFEARRAY1D

Type StackedArray
 StackUB As Long
 Stack() As Long
End Type

Type PGroupStack
 GroupsUB As Long
End Type

Type ParticleStack
 PeasUB As Long
 PeaAvail As StackedArray
 PeaInUse As StackedArray
End Type

Public PGStack As PGroupStack
Public NumParticleGroups As Long

Public PStack As ParticleStack
Public MaxParticles As Long

Dim AddDrawWidthBytes&
Dim DrawTop&
Dim DrawBot&
Dim DrawLeft&
Dim DrawRight&
Dim DrawWidthBytes&
Dim DrawWidth_&
'Dim AddDrawHeight&
Dim AddDrawWidth&
Dim AlphaArrayDepth&
Dim AlphaIndex&
Dim DrawY&
Dim DrawX&
Dim BGBlu&
Dim BGGrn&
Dim BGRed&

Dim CheckErase() As Single

Private Type EraseInfo
 FirstEraseByt As Long
 TopLeftEraseByt As Long
 EraseWidth As Long
 'DoMaskErase As Boolean
 'sngMaskAlpha As Single
 'sprWidthBytes As Long
 'SprArrayDep As Long
 'SprAlphaDep As Long
 'SprWidth As Long
End Type

Public EraseData() As EraseInfo
Public FrameCount&

Public Const BYT2 As Byte = 2
Public Const Pi! = 3.141592654
Public Const TwoPi! = 2 * Pi
Public Const B255 As Byte = 255

Public SW&        '& means  As Long
Public HalfWidth! '! means  As Single
Public SH&
Public HalfHeight!

Public Xmax&, Ymax&

Public blnRunning As Boolean
Public EraseBuf() As Long
Public Loca&, Loca1&, Loca2&
Public StepX&
Public Last_Blue_Byte&
Public Last_Alpha_byte&
Public ViewPort_Right&
Public ViewPort_TopLeft&
Public ViewPort_Right_Blue&

Public StandardSpeed As Single
Public bErasing&

Public Sub EraseSprites(Surf As AnimSurf1D)
Dim EraseWidth&
Dim BotLeftErase&
Dim TopLeftErase&
 
 If Surf.TotalPixels > 0 Then
 
 For N& = 1& To Surf.EraseSpriteCount
  
  EraseWidth = EraseData(N).EraseWidth
  BotLeftErase = EraseData(N).FirstEraseByt
  TopLeftErase = EraseData(N).TopLeftEraseByt
    
  DrawBot& = BotLeftErase
  DrawTop& = TopLeftErase
  AddDrawWidth = EraseWidth - 1&
  For Loca& = DrawBot& To DrawTop& Step Surf.Dims.Width
   DrawRight = Loca + AddDrawWidth
   For DrawX = Loca To DrawRight
    Surf.LDib(DrawX) = Surf.EraseDib(DrawX)
   Next DrawX
  Next
  
 Next N& 'Next Sprite
 
 Surf.EraseSpriteCount = 0&
 
 End If

End Sub
Public Sub LumenParticleGroup(PGroup As ParticleGroup, definition!)
Dim tmpIntensity!
 
 PGroup.IPS.chIntensity = PGroup.IPS.sBright
 PGroup.IPS.IProcess.def = definition * PGroup.IPS.chIntensity
 
 tmpIntensity = PGroup.IPS.chIntensity / B255
 PGroup.IPS.IProcess.intens = tmpIntensity / B255
 
End Sub
Public Sub NormaliseSHF(CurrentValue!, maximum!)

 SHF_Intensity = Abs(CurrentValue) / (Abs(maximum) / 6)
 
 If maximum! < 0 Or CurrentValue < 0 Then
  SHF_Intensity = 6 - SHF_Intensity
 End If

 If SHF_Intensity >= 1 Then
  If SHF_Intensity < 2 Then
  i1 = SHF_Intensity - 1
  ElseIf SHF_Intensity < 3 Then
  i2 = SHF_Intensity - 2
  ElseIf SHF_Intensity < 4 Then
  i3 = SHF_Intensity - 3
  ElseIf SHF_Intensity < 5 Then
  i4 = SHF_Intensity - 4
  ElseIf SHF_Intensity < 6 Then
  i6i = 6 - SHF_Intensity
  End If
 End If

End Sub
Public Sub HueShift()

 If iR < iB Then
  If iR < iG Then
   minimu = iR
   If iG < iB Then
    maximu = iB
   Else
    maximu = iG
   End If
  Else
   maximu = iB
   minimu = iG
  End If
 ElseIf iR > iG Then
  maximu = iR
  If iB < iG Then
   minimu = iB
  Else
   minimu = iG
  End If
 Else
  maximu = iG
  minimu = iB
 End If
 bytMaxMin_diff = maximu - minimu
 
 If iR = maximu Then
  If iB = minimu Then 'g +
   If intensity < 1 Then
    iG = iG + bytMaxMin_diff * intensity
    If iG > maximu Then
     iSubt = iG - maximu
     iG = maximu
     iR = maximu - iSubt
    End If
   ElseIf intensity < 2 Then
    iR = maximu + minimu - iG - i1 * bytMaxMin_diff
    iG = maximu
    If iR < minimu Then
     iSubt = minimu - iR
     iR = minimu
     iB = minimu + iSubt
    Else
     iB = minimu
    End If
   ElseIf intensity < 3 Then
    iB = iG + i2 * bytMaxMin_diff
    iR = minimu
    If iB > maximu Then
     iSubt = iB - maximu
     iB = maximu
     iG = maximu - iSubt
    Else
     iG = maximu
    End If
   ElseIf intensity < 4 Then
    iG = maximu + minimu - iG - i3 * bytMaxMin_diff
    iB = maximu
    If iG < minimu Then
     iSubt = minimu - iG
     iG = minimu
     iR = minimu + iSubt
    Else
     iR = minimu
    End If
   ElseIf intensity < 5 Then
    iR = iG + i4 * bytMaxMin_diff
    iG = minimu
    If iR > maximu Then
     iSubt = iR - maximu
     iR = maximu
     iB = maximu - iSubt
    Else
     iB = maximu
    End If
   ElseIf intensity < 6 Then
    iG = iG - i6i * bytMaxMin_diff
    If iG < minimu Then
     iSubt = minimu - iG
     iG = minimu
     iB = minimu + iSubt
    Else
     iB = minimu
    End If
   End If
  Else 'r max, g min, blue -
   If intensity < 1 Then
    iB = iB - bytMaxMin_diff * intensity
    If iB < minimu Then
     iSubt = minimu - iB
     iB = minimu
     iG = iG + iSubt
    End If
   ElseIf intensity < 2 Then
    iG = maximu + minimu - iB + bytMaxMin_diff * i1
    iB = minimu
    If iG > maximu Then
     iSubt = iG - maximu
     iG = maximu
     iR = maximu - iSubt
    Else
     iR = maximu
    End If
   ElseIf intensity < 3 Then
    iR = iB - bytMaxMin_diff * i2
    iG = maximu
    If iR < minimu Then
     iSubt = minimu - iR
     iR = minimu
     iB = minimu + iSubt
    Else
     iB = minimu
    End If
   ElseIf intensity < 4 Then
    iB = maximu + minimu - iB + bytMaxMin_diff * i3
    iR = minimu
    If iB > maximu Then
     iSubt = iB - maximu
     iB = maximu
     iG = maximu - iSubt
    Else
     iG = maximu
    End If
   ElseIf intensity < 5 Then
    iG = iB - bytMaxMin_diff * i4
    iB = maximu
    If iG < minimu Then
     iSubt = minimu - iG
     iG = minimu
     iR = minimu + iSubt
    Else
     iR = minimu
    End If
   ElseIf intensity < 6 Then
    iB = iB + bytMaxMin_diff * i6i
    If iB > maximu Then
     iSubt = iB - maximu
     iB = maximu
     iR = maximu - iSubt
    Else
     iR = maximu
    End If
   End If
  End If
 ElseIf iR = minimu Then
  If iG = maximu Then 'blue +
   If intensity < 1 Then
    iB = iB + bytMaxMin_diff * intensity
    If iB > maximu Then
     iSubt = iB - maximu
     iB = maximu
     iG = maximu - iSubt
    End If
   ElseIf intensity < 2 Then
    iG = maximu + minimu - iB - bytMaxMin_diff * i1
    iB = maximu
    If iG < minimu Then
     iSubt = minimu - iG
     iG = minimu
     iR = minimu + iSubt
    Else
     iR = minimu
    End If
   ElseIf intensity < 3 Then
    iR = iB + bytMaxMin_diff * i2
    iG = minimu
    If iR > maximu Then
     iSubt = iR - maximu
     iR = maximu
     iB = maximu - iSubt
    Else
     iB = maximu
    End If
   ElseIf intensity < 4 Then
    iB = maximu + minimu - iB - bytMaxMin_diff * i3
    iR = maximu
    If iB < minimu Then
     iSubt = minimu - iB
     iB = minimu
     iG = minimu + iSubt
    Else
     iG = minimu
    End If
   ElseIf intensity < 5 Then
    iG = iB + bytMaxMin_diff * i4
    iB = minimu
    If iG > maximu Then
     iSubt = iG - maximu
     iG = maximu
     iR = maximu - iSubt
    Else
     iR = maximu
    End If
   ElseIf intensity < 6 Then
    iB = iB - bytMaxMin_diff * i6i
    If iB < minimu Then
     iSubt = minimu - iB
     iB = minimu
     iR = minimu + iSubt
    Else
     iR = minimu
    End If
   End If
  Else 'blue max green -
   If intensity < 1 Then
    iG = iG - bytMaxMin_diff * intensity
    If iG < minimu Then
     iSubt = minimu - iG
     iG = minimu
     iR = minimu + iSubt
    End If
   ElseIf intensity < 2 Then
    iR = maximu + minimu - iG + bytMaxMin_diff * i1
    iG = minimu
    If iR > maximu Then
     iSubt = iR - maximu
     iR = maximu
     iB = maximu - iSubt
    Else
     iB = maximu
    End If
   ElseIf intensity < 3 Then
    iB = iG - bytMaxMin_diff * i2
    iR = maximu
    If iB < minimu Then
     iSubt = minimu - iB
     iB = minimu
     iG = minimu + iSubt
    Else
     iG = minimu
    End If
   ElseIf intensity < 4 Then
    iG = maximu + minimu - iG + bytMaxMin_diff * i3
    iB = minimu
    If iG > maximu Then
     iSubt = iG - maximu
     iG = maximu
     iR = maximu - iSubt
    Else
     iR = maximu
    End If
   ElseIf intensity < 5 Then
    iR = iG - bytMaxMin_diff * i4
    iG = maximu
    If iR < minimu Then
     iSubt = minimu - iR
     iR = minimu
     iB = minimu + iSubt
    Else
     iB = minimu
    End If
   ElseIf intensity < 6 Then
    iG = iG + bytMaxMin_diff * i6i
    If iG > maximu Then
     iSubt = iG - maximu
     iG = maximu
     iB = maximu - iSubt
    Else
     iB = maximu
    End If
   End If
  End If
 ElseIf iB = maximu Then 'green min, red +
  If intensity < 1 Then
   iR = iR + bytMaxMin_diff * intensity
   If iR > maximu Then
    iSubt = iR - maximu
    iR = maximu
    iB = maximu - iSubt
   End If
  ElseIf intensity < 2 Then
   iB = maximu + minimu - iR - bytMaxMin_diff * i1
   iR = maximu
   If iB < minimu Then
    iSubt = minimu - iB
    iB = minimu
    iG = minimu + iSubt
   Else
    iG = minimu
   End If
  ElseIf intensity < 3 Then
   iG = iR + bytMaxMin_diff * i2
   iB = minimu
   If iG > maximu Then
    iSubt = iG - maximu
    iG = maximu
    iR = maximu - iSubt
   Else
    iR = maximu
   End If
  ElseIf intensity < 4 Then
   iR = maximu + minimu - iR - bytMaxMin_diff * i3
   iG = maximu
   If iR < minimu Then
    iSubt = minimu - iR
    iR = minimu
    iB = minimu + iSubt
   Else
    iB = minimu
   End If
  ElseIf intensity < 5 Then
   iB = iR + bytMaxMin_diff * i4
   iR = minimu
   If iB > maximu Then
    iSubt = iB - maximu
    iB = maximu
    iG = maximu - iSubt
   Else
    iG = maximu
   End If
  ElseIf intensity < 6 Then
   iR = iR - bytMaxMin_diff * i6i
   If iR < minimu Then
    iSubt = minimu - iR
    iR = minimu
    iG = minimu + iSubt
   Else
    iG = minimu
   End If
  End If
 Else 'blue min, green max, red -
  If intensity < 1 Then
   iR = iR - bytMaxMin_diff * intensity
   If iR < minimu Then
    iSubt = minimu - iR
    iR = minimu
    iB = minimu + iSubt
   End If
  ElseIf intensity < 2 Then
   iB = minimu - iR + maximu + bytMaxMin_diff * i1
   iR = minimu
   If iB > maximu Then
    iSubt = iB - maximu
    iB = maximu
    iG = maximu - iSubt
   Else
    iG = maximu
   End If
  ElseIf intensity < 3 Then
   iG = iR - bytMaxMin_diff * i2
   iB = maximu
   If iG < minimu Then
    iSubt = minimu - iG
    iG = minimu
    iR = minimu + iSubt
   Else
    iR = minimu
   End If
  ElseIf intensity < 4 Then
   iR = maximu + minimu - iR + bytMaxMin_diff * i3
   iG = minimu
   If iR > maximu Then
    iSubt = iR - maximu
    iR = maximu
    iB = maximu - iSubt
   Else
    iB = maximu
   End If
  ElseIf intensity < 5 Then
   iB = iR - bytMaxMin_diff * i4
   iR = maximu
   If iB < minimu Then
    iSubt = minimu - iB
    iB = minimu
    iG = minimu + iSubt
   Else
    iG = minimu
   End If
  ElseIf intensity < 6 Then
   iR = iR + bytMaxMin_diff * i6i
   If iR > maximu Then
    iSubt = iR - maximu
    iR = maximu
    iG = maximu - iSubt
   Else
    iG = maximu
   End If
  End If
 End If
 
End Sub
Public Sub BlitParticle(Surf As AnimSurf1D, Pea As ParticleSprite)
Dim b_slope!
Dim height_sq!
Dim dx_dy_Sq!
Dim inv_press!
Dim inv_press_sq!

 Left_ = Int(Pea.px - Pea.Dims.radius)
 Right_ = Int(Pea.px + Pea.Dims.radius)
 
 Bot_ = Int(Pea.py - Pea.Dims.radius)
 Top_ = Int(Pea.py + Pea.Dims.radius)
 
 If Left_ < 0 Then
  DrawLeft = -Left_
  Left_ = 0
 Else
  DrawLeft = Left_
 End If
 
 If Bot_ < 0 Then
  DrawBot = -Bot_ * Surf.Dims.Width
  Bot_ = 0
 Else
  DrawBot = Bot_ * Surf.Dims.Width
 End If
 
 If Top_ > Surf.TopRight.Y Then Top_ = Surf.TopRight.Y
 If Right_ > Surf.TopRight.X Then Right_ = Surf.TopRight.X
 
 DrawTop = Top_ * Surf.Dims.Width
 AddDrawWidth = Right_ - Left_
 
 DrawBot = DrawBot + Left_
 DrawTop = DrawTop + Left_
 
 If bErasing Then
  Surf.EraseSpriteCount = Surf.EraseSpriteCount + 1
  ReDim Preserve EraseData(1 To Surf.EraseSpriteCount)
  With EraseData(Surf.EraseSpriteCount)
   .FirstEraseByt = DrawBot&
   .TopLeftEraseByt = DrawTop&
   .EraseWidth = AddDrawWidth + 1
  End With
 End If
  
 b_slope = Pea.def / Pea.Dims.radius
 
 delta_y = (Bot_ - Pea.py) * b_slope
 baseleft = (Left_ - Pea.px) * b_slope
 
 height_sq = Pea.def * Pea.def
 inv_press = Pea.def - 255
 
 If inv_press < 0 Then inv_press = 0
 
 inv_press_sq = inv_press * inv_press
  
 For DrawY = DrawBot To DrawTop Step Surf.Dims.Width
  delta_ySq! = delta_y * delta_y
  delta_x! = baseleft
  DrawRight = DrawY + AddDrawWidth
  For DrawX = DrawY To DrawRight
   dx_dy_Sq = delta_x * delta_x + delta_ySq
   If dx_dy_Sq < height_sq Then
    If dx_dy_Sq > inv_press_sq Then
     Bright = Pea.def - Sqr(dx_dy_Sq)
    Else
     Bright = 255
    End If
    BGBlu = Surf.Dib(DrawX).Blue
    BGGrn = Surf.Dib(DrawX).Green
    BGRed = Surf.Dib(DrawX).Red
    Surf.Dib(DrawX).Blue = BGBlu + Bright * (Pea.sBlu - BGBlu) / 255
    Surf.Dib(DrawX).Green = BGGrn + Bright * (Pea.sGrn - BGGrn) / 255
    Surf.Dib(DrawX).Red = BGRed + Bright * (Pea.sRed - BGRed) / 255
   End If
   delta_x = delta_x + b_slope
  Next DrawX
  delta_y = delta_y + b_slope
 Next DrawY
  
  
  
  
End Sub
Public Sub BlitBrightParticle(Surf As AnimSurf1D, Pea As ParticleSprite)
Dim b_slope!
Dim height_sq!
Dim dx_dy_Sq!
Dim inv_press!
Dim inv_press_sq!

 Left_ = Int(Pea.px - Pea.Dims.radius)
 Right_ = Int(Pea.px + Pea.Dims.radius)
 
 Bot_ = Int(Pea.py - Pea.Dims.radius)
 Top_ = Int(Pea.py + Pea.Dims.radius)
 
 If Left_ < 0 Then
  DrawLeft = -Left_
  Left_ = 0
 Else
  DrawLeft = Left_
 End If
 
 If Bot_ < 0 Then
  DrawBot = -Bot_ * Surf.Dims.Width
  Bot_ = 0
 Else
  DrawBot = Bot_ * Surf.Dims.Width
 End If
 
 If Top_ > Surf.TopRight.Y Then Top_ = Surf.TopRight.Y
 If Right_ > Surf.TopRight.X Then Right_ = Surf.TopRight.X
 
 DrawTop = Top_ * Surf.Dims.Width
 AddDrawWidth = Right_ - Left_
 
 DrawBot = DrawBot + Left_
 DrawTop = DrawTop + Left_
 
 If bErasing Then
  Surf.EraseSpriteCount = Surf.EraseSpriteCount + 1
  ReDim Preserve EraseData(1 To Surf.EraseSpriteCount)
  With EraseData(Surf.EraseSpriteCount)
   .FirstEraseByt = DrawBot&
   .TopLeftEraseByt = DrawTop&
   .EraseWidth = AddDrawWidth + 1
  End With
 End If
  
 b_slope = Pea.def / Pea.Dims.radius
 
 delta_y = (Bot_ - Pea.py) * b_slope
 baseleft = (Left_ - Pea.px) * b_slope
 
 height_sq = Pea.def * Pea.def
 inv_press = Pea.def - 255
 
 If inv_press < 0 Then inv_press = 0
 
 inv_press_sq = inv_press * inv_press
  
 For DrawY = DrawBot To DrawTop Step Surf.Dims.Width
  delta_ySq! = delta_y * delta_y
  delta_x! = baseleft
  DrawRight = DrawY + AddDrawWidth
  For DrawX = DrawY To DrawRight
   dx_dy_Sq = delta_x * delta_x + delta_ySq
   If dx_dy_Sq < height_sq Then
    If dx_dy_Sq > inv_press_sq Then
     Bright = Pea.def - Sqr(dx_dy_Sq)
    Else
     Bright = 255
    End If
    BGBlu = Surf.Dib(DrawX).Blue + Bright& * Pea.sBlu / 255&
    BGGrn = Surf.Dib(DrawX).Green + Bright& * Pea.sGrn / 255&
    BGRed = Surf.Dib(DrawX).Red + Bright& * Pea.sRed / 255&
    If BGBlu > 255 Then BGBlu = 255
    If BGGrn > 255 Then BGGrn = 255
    If BGRed > 255 Then BGRed = 255
    Surf.Dib(DrawX).Blue = BGBlu
    Surf.Dib(DrawX).Green = BGGrn
    Surf.Dib(DrawX).Red = BGRed
   End If
   delta_x = delta_x + b_slope
  Next DrawX
  delta_y = delta_y + b_slope
 Next DrawY
 
End Sub
Public Sub DimensionParticle(AS1 As ParticleSprite, particDiameter!)
 
 AS1.Dims.diameter = particDiameter
 AS1.Dims.radius = particDiameter / 2
 
End Sub

Public Function RealRound(ByVal sngValue!) As Long
Dim diff!
 'This function rounds .5 up
 
 RealRound = Int(sngValue)
 diff = sngValue - RealRound
 If diff >= 0.5! Then RealRound = RealRound + 1&

End Function
Public Function RealRound2(ByVal sngValue!) As Long
Dim diff!
 'This function rounds .5 down
 
 RealRound2 = Int(sngValue)
 diff = sngValue - RealRound2
 If diff > 0.5! Then RealRound2 = RealRound2 + 1&

End Function

