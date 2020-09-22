VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Firehawk"
   ClientHeight    =   5490
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'FireHawk 1.4
'by dafhi

'inspired by Yat Seng's  Sprite Tracking

Dim N&
Dim N1&

Dim homeX As Single         ' X coordinate for 1st circle's destination
Dim homeY As Single         ' Y coordinate for 1st circle's destination
Dim tempX!        ' X coordinate for subsequent circles
Dim tempY!        ' Y coordinate for subsequent circles

Dim move_Slowness!   ' Sprite speed
Dim track_Slowness!  ' Tracking Speed

'General
Dim sing!

Dim standardSpeedControl!

Dim ParticleLightness&

Dim PGroupsOrderArray&()

Dim explosiveBase!
Dim explosiveVariance!
Dim ParticleSizes!
Dim emissionsBase!
Dim emissionsVariance!
Dim durationsBase!
Dim durationsVariance!

Dim PBLoops&
Dim NFrames&
Dim radius!
Dim sineparm!

Dim boolA As Boolean 'raise particleSizes
Dim boolZ As Boolean 'lower

Dim boolS As Boolean 'raise explosiveBase
Dim boolX As Boolean 'lower

Dim boolD As Boolean 'raise explosiveVariance
Dim boolC As Boolean 'lower

Dim boolF As Boolean 'raise emissionsBase
Dim boolV As Boolean 'lower

Dim boolG As Boolean 'raise emissionsVariance
Dim boolB As Boolean 'lower

Dim boolH As Boolean 'raise durationsBase
Dim boolN As Boolean 'lower

Dim boolJ As Boolean 'raise durationsVariance
Dim boolM As Boolean 'lower

Dim bool1 As Boolean 'raise sGroupCount
Dim boolQ As Boolean 'lower

Dim bool2 As Boolean 'raise sMaxParticles
Dim boolW As Boolean 'lower

Dim bool3 As Boolean 'first group spectrum shift
Dim boolE As Boolean 'shift reverse

Dim bool4 As Boolean 'increase color spread
Dim boolR As Boolean 'decrease

Const SIZES_MAX! = 15
Const SIZES_MIN! = 0

Const XPLOBASE_MIN! = 0
Const XPLOBASE_MAX! = 250
Const XPLOVARI_MIN! = 0
Const XPLOVARI_MAX! = XPLOBASE_MAX

Const EMSNBASE_MIN! = 0.1
Const EMSNBASE_MAX! = 80
Const EMSNVARI_MIN! = 0
Const EMSNVARI_MAX! = 2 * EMSNBASE_MAX

Const DURABASE_MIN! = 0.05!
Const DURABASE_MAX! = 5!
Const DURAVARI_MIN! = 0!
Const DURAVARI_MAX! = 5!

Const PARTICLES_MIN& = 5
Const PARTICLES_MAX& = 50000
Const P_MAX_UB& = PARTICLES_MAX - 1&

Const GROUPS_MIN& = 1
Const GROUPS_MAX& = 100

Const LIGHTNESS_MAX& = 60
Const LIGHTNESS_MIN& = 2

Dim sizes_adjustRate!
Dim iSizesRate!
Dim iSizesRateAccel!

Dim xploBase_adjustRate!
Dim xploVari_adjustRate!
Dim emsnBase_adjustRate!
Dim emsnVari_adjustRate!
Dim duraBase_adjustRate!
Dim duraVari_adjustRate!

Dim sGroupCount!
Dim iGroupCount!
Dim iGroupCountAccel!
Dim iGroupCountAccelA!

Dim sMaxParticles!
Dim iMaxParticles!
Dim iMaxParticlesAccel!
Dim iMaxParticlesAccelA!

Dim s1Red!
Dim s1Grn!
Dim s1Blu!

Dim s1Intensity!
Dim s2Intensity!

Dim sShiftSpeed!  'rate of change for first group's color
Dim sSpreadSpeed! 'when user changes group spread

Dim pCountBasedRateMult!

Dim iTrackSlow!

Dim trackSlowMax!
Dim trackSlowMin!

Dim PGroups(GROUPS_MAX - 1&) As ParticleGroup
Dim Peas(P_MAX_UB) As ParticleSprite

Dim sParticleSizesLUT(P_MAX_UB) As Single

Dim bBlitBright As Boolean
Dim bBackground2 As Boolean

Dim Elapsed&
Dim LastTic&
Dim StandardSpeed!
Dim Tick&
Dim TickSum&
Dim fps!
Dim FrameCount&

'When to change target point
Dim sFrame!
Dim maxFrame!

Dim FormDib As AnimSurf1D

Dim bMovingGroup1&

Dim CSeL&

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub Form_Load()
ReDim PStack.PeaAvail.Stack(P_MAX_UB)
ReDim PStack.PeaInUse.Stack(P_MAX_UB)

 PStack.PeasUB = -1

 Randomize
 ForeColor = vbWhite
 ScaleMode = vbPixels
 Left = 1000
 Top = 1000
 
 CSeL = Int(Rnd * 6.99)
 Show

 standardSpeedControl = 0.0065

 SystemTest
 
 NumParticleGroups = 50

 MaxParticles = 500
 
 explosiveBase = 0
 explosiveVariance = 0

'shower rate.  affects particle groups
 emissionsBase = 10
 emissionsVariance = 10

 'These are multiplied by StandardSpeed later
 durationsBase = 2
 durationsVariance = 1

 'variance
 ParticleSizes = 2.5

 'Adjust this for a different throw
 ParticleLightness = 7

 blnRunning = True

 sGroupCount = NumParticleGroups
 sMaxParticles = MaxParticles

 s1Red = 0   'keep one of these 0 and one 255 for maximum saturation
 s1Grn = 0
 s1Blu = 255

 s1Intensity = 6& * Rnd! 'maximum of 6
 s2Intensity = 6! 'maximum of 6

 'when a new set of movement points is randomized
 maxFrame = 4
 'this is closer to 100x what you put it, with increment being Single

 bBlitBright = False
 bErasing = True
 bMovingGroup1 = True

 InitParticles
 InitParticleGroups
 InitStandardSpeedVariables
 InitParticleSizesLUT
 ParticleSizesOnce
 
 While blnRunning 'press Esc to terminate this loop
 
  InitTrack PGroups(0&)
  
  For N = 1& To PGStack.GroupsUB
   Call track(PGroups(N - 1&), PGroups(N)) ' following circles tracks their previous circle
  Next N

  HandleKeys
  
  If SH > 0& Then DrawExplosions
  
  FrameCount = FrameCount + 1
  
  If bMovingGroup1 Then
   If sFrame > maxFrame Then
    homeX = SW * (0.5 + (Rnd - 0.5) / 0.8!)
    homeY = SH * (0.5 + (Rnd - 0.5) / 0.8!)
    sFrame = 0
   End If
   sFrame = sFrame + StandardSpeed
  End If
  
  Refresh
  
  CurrentY = 5
  Print "FPS: " & Round(fps, 1)
  Print
  Print "1  2  3  4  5  6"
  Print "Q W E R T Y"
  Print
  Print "A S D F G H J"
  Print "Z X C V B N M"
  Print
  Print " i, o, p"
  Print "spacebar"
  
  CalcTick
  CalcFPS
  
  If bErasing Then EraseSprites FormDib
  
  DoEvents
   
 Wend
 
 'We've only reached this point if the user has quit
 
 CleanUp FormDib
 
 'Set Form1 = Nothing
 
 Unload Me
 
 End

End Sub
Private Sub track(PGLeader As ParticleGroup, PG As ParticleGroup)
 PG.IPS.cdx = (PGLeader.IPS.centerX - PG.IPS.centerX) / track_Slowness
 PG.IPS.cdy = (PGLeader.IPS.centerY - PG.IPS.centerY) / track_Slowness
 PG.IPS.centerX = PG.IPS.centerX + PG.IPS.cdx
 PG.IPS.centerY = PG.IPS.centerY + PG.IPS.cdy
End Sub
Private Sub InitTrack(PG As ParticleGroup)
 PG.IPS.cdx = (homeX - PG.IPS.centerX) / move_Slowness
 PG.IPS.cdy = (homeY - PG.IPS.centerY) / move_Slowness
 PG.IPS.centerX = PG.IPS.centerX + PG.IPS.cdx
 PG.IPS.centerY = PG.IPS.centerY + PG.IPS.cdy
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
homeX = X
homeY = FormDib.Dims.Height - Y
sFrame = 0
End Sub


Private Sub HandleKeys()

pCountBasedRateMult = PStack.PeaInUse.StackUB / 1000&

'particleSizes
If boolA Xor boolZ Then 'if one xor the other is pressed
 Caption = "Particle Size Variance: " & ParticleSizes
 If boolA Then
  If ParticleSizes >= SIZES_MAX Then
   ParticleSizes = SIZES_MAX
   InitParticleSizes PStack.PeaInUse.StackUB
  ElseIf ParticleSizes < SIZES_MAX Then
   InitParticleSizes PStack.PeaInUse.StackUB
   ParticleSizes = ParticleSizes + sizes_adjustRate
  End If
 Else
  If ParticleSizes <= SIZES_MIN Then
   ParticleSizes = SIZES_MIN
   InitParticleSizes PStack.PeaInUse.StackUB
  ElseIf ParticleSizes > SIZES_MIN Then
   InitParticleSizes PStack.PeaInUse.StackUB
   ParticleSizes = ParticleSizes - sizes_adjustRate
  End If
 End If
 sizes_adjustRate = sizes_adjustRate + iSizesRate
 iSizesRate = iSizesRate + iSizesRateAccel
End If

'explosiveBase
If boolS Xor boolX Then 'if one xor the other is pressed
 Caption = "explosiveBase: " & explosiveBase
 If boolS Then
  If explosiveBase > XPLOBASE_MAX Then
   explosiveBase = XPLOBASE_MAX
   InitExplosives
  ElseIf explosiveBase < XPLOBASE_MAX Then
   InitExplosives
   explosiveBase = explosiveBase + xploBase_adjustRate * pCountBasedRateMult
  End If
 Else
  If explosiveBase < XPLOBASE_MIN Then
   explosiveBase = XPLOBASE_MIN
   InitExplosives
  ElseIf explosiveBase > XPLOBASE_MIN Then
   InitExplosives
   explosiveBase = explosiveBase - xploBase_adjustRate * pCountBasedRateMult
  End If
 End If
End If

'explosiveVariance
If boolD Xor boolC Then
 Caption = "explosiveVariance: " & explosiveVariance
 If boolD Then
  If explosiveVariance > XPLOVARI_MAX Then
   explosiveVariance = XPLOVARI_MAX
   InitExplosives
  ElseIf explosiveVariance < XPLOVARI_MAX Then
   InitExplosives
   explosiveVariance = explosiveVariance + xploVari_adjustRate * pCountBasedRateMult
  End If
 Else
  If explosiveVariance < XPLOVARI_MIN Then
   explosiveVariance = XPLOVARI_MIN
   InitExplosives
  ElseIf explosiveVariance > XPLOVARI_MIN Then
   InitExplosives
   explosiveVariance = explosiveVariance - xploVari_adjustRate * pCountBasedRateMult
  End If
 End If
End If

'emissionsBase
If boolF Xor boolV Then
 Caption = "emissionsBase: " & FormatPercent((emissionsBase - EMSNBASE_MIN) / (EMSNBASE_MAX - EMSNBASE_MIN), 1)
 If boolF Then
  If emissionsBase > EMSNBASE_MAX Then
   emissionsBase = EMSNBASE_MAX
   InitEmitRates
  ElseIf emissionsBase < EMSNBASE_MAX Then
   InitEmitRates
   emissionsBase = emissionsBase + emsnBase_adjustRate * pCountBasedRateMult
  End If
 Else
  If emissionsBase < EMSNBASE_MIN Then
   emissionsBase = EMSNBASE_MIN
   InitEmitRates
  ElseIf emissionsBase > EMSNBASE_MIN Then
   InitEmitRates
   emissionsBase = emissionsBase - emsnBase_adjustRate * pCountBasedRateMult
  End If
 End If
End If

'emissionsVariance
If boolG Xor boolB Then
 Caption = "emissionsVariance: " & FormatPercent((emissionsVariance - EMSNVARI_MIN) / (EMSNVARI_MAX - EMSNVARI_MIN), 1)
 If boolG Then
  If emissionsVariance > EMSNVARI_MAX Then
   emissionsVariance = EMSNVARI_MAX
   InitEmitRates
  ElseIf emissionsVariance < EMSNVARI_MAX Then
   InitEmitRates
   emissionsVariance = emissionsVariance + emsnVari_adjustRate * pCountBasedRateMult
  End If
 Else
  If emissionsVariance < EMSNVARI_MIN Then
   emissionsVariance = EMSNVARI_MIN
   InitEmitRates
  ElseIf emissionsVariance > EMSNVARI_MIN Then
   InitEmitRates
   emissionsVariance = emissionsVariance - emsnVari_adjustRate * pCountBasedRateMult
  End If
 End If
End If

'durationsBase
If boolH Xor boolN Then
 Caption = "durationsBase: " & FormatPercent((durationsBase - DURABASE_MIN) / (DURABASE_MAX - DURABASE_MIN), 1)
 If boolH Then
  durationsBase = durationsBase + duraBase_adjustRate '* pCountBasedRateMult
  If durationsBase > DURABASE_MAX Then
   durationsBase = DURABASE_MAX
  End If
 Else
  durationsBase = durationsBase - duraBase_adjustRate '* pCountBasedRateMult
  If durationsBase < DURABASE_MIN Then
   durationsBase = DURABASE_MIN
  End If
 End If
End If

'durationsVariance
If boolJ Xor boolM Then
 Caption = "durationsVariance: " & FormatPercent((durationsVariance - DURAVARI_MIN) / (DURAVARI_MAX - DURAVARI_MIN), 1)
 If boolJ Then
  durationsVariance = durationsVariance + duraVari_adjustRate
  If durationsVariance > DURAVARI_MAX Then
   durationsVariance = DURAVARI_MAX
  End If
 Else
  durationsVariance = durationsVariance - duraVari_adjustRate
  If durationsVariance < DURAVARI_MIN Then
   durationsVariance = DURAVARI_MIN
  End If
 End If
End If

'particle group count
If bool1 Xor boolQ Then
 Caption = "Num Particle Groups: " & NumParticleGroups
End If

'MaxParticles
If bool2 Xor boolW Then
 Caption = "Max Particle Count: " & MaxParticles
 If bool2 Then
  If MaxParticles > PARTICLES_MAX Then
   MaxParticles = PARTICLES_MAX
   InitParticles
  ElseIf MaxParticles < PARTICLES_MAX Then
   InitParticles
   MaxParticles = MaxParticles + iMaxParticles
  End If
 Else
  If MaxParticles <= PARTICLES_MIN Then
   MaxParticles = PARTICLES_MIN
   InitParticles
  Else
   InitParticles
   MaxParticles = MaxParticles - iMaxParticles
  End If
 End If
 iMaxParticles = iMaxParticles + iMaxParticlesAccel
 iMaxParticlesAccel = iMaxParticlesAccel + iMaxParticlesAccelA
End If

'shift color of first particle group
If bool3 Xor boolE Then
 Caption = "First Group RGB: " & Int(s1Red) & " " & Int(s1Grn) & " " & Int(s1Blu)
 If bool3 Then
  iR = s1Red
  iG = s1Grn
  iB = s1Blu
  intensity = sShiftSpeed '* PStack.PeaInUse.StackUB
  NormaliseSHF intensity, 6!
  HueShift
  s1Red = iR
  s1Grn = iG
  s1Blu = iB
  ColorGroups
 Else
  iR = s1Red
  iG = s1Grn
  iB = s1Blu
  intensity = 6! - sShiftSpeed '* PStack.PeaInUse.StackUB
  NormaliseSHF intensity, 6!
  HueShift
  s1Red = iR
  s1Grn = iG
  s1Blu = iB
  ColorGroups
 End If
End If

'color spread
If bool4 Xor boolR Then
 Caption = "Color Spread: " & FormatPercent(s2Intensity / 6!)
 If bool4 Then
  s2Intensity = s2Intensity + sSpreadSpeed * pCountBasedRateMult * 64&
  If s2Intensity > 6! Then s2Intensity = 6!
  ColorGroups
 Else
  s2Intensity = s2Intensity - sSpreadSpeed * pCountBasedRateMult * 64&
  If s2Intensity < 0! Then s2Intensity = 0!
  ColorGroups
 End If
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
 Case vbKeyA
  boolA = True
 Case vbKeyZ
  boolZ = True
 Case vbKeyS
  boolS = True
 Case vbKeyX
  boolX = True
 Case vbKeyD
  boolD = True
 Case vbKeyC
  boolC = True
 Case vbKeyF
  boolF = True
 Case vbKeyV
  boolV = True
 Case vbKeyG
  boolG = True
 Case vbKeyB
  boolB = True
 Case vbKeyH
  boolH = True
 Case vbKeyN
  boolN = True
 Case vbKeyJ
  boolJ = True
 Case vbKeyM
  boolM = True
 Case vbKey1
  bool1 = True
  sGroupCount = sGroupCount + iGroupCount
  If sGroupCount > GROUPS_MAX Then
   sGroupCount = GROUPS_MAX
  End If
  iGroupCount = iGroupCount + iGroupCountAccel
  NumParticleGroups = sGroupCount
  InitParticleGroups
 Case vbKeyQ
  boolQ = True
  sGroupCount = sGroupCount - iGroupCount
  If sGroupCount < GROUPS_MIN Then
   sGroupCount = GROUPS_MIN
  End If
  iGroupCount = iGroupCount + iGroupCountAccel
  NumParticleGroups = sGroupCount
  InitParticleGroups
 Case vbKey2
  bool2 = True
 Case vbKeyW
  boolW = True
 Case vbKey3
  bool3 = True
 Case vbKeyE
  boolE = True
 Case vbKey4
  bool4 = True
 Case vbKeyR
  boolR = True
 Case vbKey5
  If ParticleLightness < LIGHTNESS_MAX Then
   ParticleLightness = ParticleLightness + 1
  End If
  Caption = "Particle Lightness: " & ParticleLightness
 Case vbKeyT
  If ParticleLightness > LIGHTNESS_MIN Then
   ParticleLightness = ParticleLightness - 1
  End If
  Caption = "Particle Lightness: " & ParticleLightness
 Case vbKey6
  track_Slowness = track_Slowness + iTrackSlow
  If track_Slowness > trackSlowMax Then track_Slowness = trackSlowMax
  Caption = "Tracking Slowness: " & FormatPercent((track_Slowness - trackSlowMin) / (trackSlowMax - trackSlowMin), 1)
 Case vbKeyY
  track_Slowness = track_Slowness - iTrackSlow
  If track_Slowness < trackSlowMin Then track_Slowness = trackSlowMin
  Caption = "Tracking Slowness: " & FormatPercent((track_Slowness - trackSlowMin) / (trackSlowMax - trackSlowMin), 1)
 Case vbKeySpace
  CSeL = Int(Rnd * 6.99)
  bBackground2 = Not bBackground2
  GenerateBackground FormDib
 Case vbKeyO
  bErasing = Not bErasing
 Case vbKeyI
  bMovingGroup1 = Not bMovingGroup1
 Case vbKeyP
  bBlitBright = Not bBlitBright
 Case vbKeyEscape
  blnRunning = False
 End Select
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
 Case vbKeyA
  boolA = False
  SetSizesRateInc
 Case vbKeyZ
  boolZ = False
  SetSizesRateInc
 Case vbKeyS
  boolS = False
 Case vbKeyX
  boolX = False
 Case vbKeyD
  boolD = False
 Case vbKeyC
  boolC = False
 Case vbKeyF
  boolF = False
 Case vbKeyV
  boolV = False
 Case vbKeyG
  boolG = False
 Case vbKeyB
  boolB = False
 Case vbKeyH
  boolH = False
 Case vbKeyN
  boolN = False
 Case vbKeyJ
  boolJ = False
 Case vbKeyM
  boolM = False
 Case vbKey1
  bool1 = False
  iGroupCount = 1!
 Case vbKeyQ
  boolQ = False
  iGroupCount = 1!
 Case vbKey2
  bool2 = False
  SetMaxParticlesInc
 Case vbKeyW
  boolW = False
  SetMaxParticlesInc
 Case vbKey3
  bool3 = False
 Case vbKeyE
  boolE = False
 Case vbKey4
  bool4 = False
 Case vbKeyR
  boolR = False
 End Select
 Caption = "Firehawk"
End Sub
Public Sub DrawExplosions()
Dim L&, sing!, N2&

 N = -1&
 
 'This designates which particle groups get the most
 'particles based on Rnd probability
 Do While N < PGStack.GroupsUB
  N = N + 1&
  L = PGroupsOrderArray(N)
  sing = PGroups(L).emitRate
  N2 = Int(sing)
  sing = sing - N2 'take the fractional part of emitRate
  'if Rnd < .37, for example
  If Rnd < sing Then
   AddParticle PGroups(L)
  End If
  'If the Whole part of emitRate was 2,
  '2 chances to add a particle to this group
  For N1 = 1& To N2
   If Rnd < 0.5! Then AddParticle PGroups(L)
  Next
 Loop
 
 'shift the order for who gets first dibs on
 'available particles in above loop
 For N = 1& To PGStack.GroupsUB
  If Rnd < 0.5! Then
   L = PGroupsOrderArray(0&)
   PGroupsOrderArray(0&) = PGroupsOrderArray(N)
   PGroupsOrderArray(N) = L
  End If
 Next
 
 If FormDib.TotalPixels > 0 Then
  N = -1&
  If bBlitBright Then
  Do While N < PStack.PeaInUse.StackUB
   N = N + 1&
   L = PStack.PeaInUse.Stack(N)
   AdvanceBrightParticle Peas(L), N
  Loop
  Else
  Do While N < PStack.PeaInUse.StackUB
   N = N + 1&
   L = PStack.PeaInUse.Stack(N)
   AdvanceParticle Peas(L), N
  Loop
  End If
 End If
 
End Sub

Private Sub AdvanceBrightParticle(Pea As ParticleSprite, Elem&)
Dim APGLoop&, cdxs!, cdys!
 If Pea.sFrame < Pea.maxFrame Then
  BlitBrightParticle FormDib, Pea
  Pea.px = Pea.px + Pea.ix
  Pea.py = Pea.py + Pea.iy
  Pea.def = Pea.def - Pea.iLum
  Pea.sFrame = Pea.sFrame + Pea.iFrame
 Else
  DeleteParticle Elem
 End If
End Sub
Private Sub AdvanceParticle(Pea As ParticleSprite, Elem&)
Dim APGLoop&, cdxs!, cdys!
 If Pea.sFrame < Pea.maxFrame Then
  BlitParticle FormDib, Pea
  Pea.px = Pea.px + Pea.ix
  Pea.py = Pea.py + Pea.iy
  Pea.def = Pea.def - Pea.iLum
  Pea.sFrame = Pea.sFrame + Pea.iFrame
 Else
  DeleteParticle Elem
 End If
End Sub


Private Sub SystemTest()
Dim NumPasses&

 Elapsed = 0&
  
 LastTic = timeGetTime
 
 While Elapsed < 200&
  LoopTestCode
  Elapsed = timeGetTime - LastTic
  NumPasses = NumPasses + 1&
 Wend
 
 StandardSpeed = (Elapsed / NumPasses) * standardSpeedControl
 
 iGroupCount! = 0.08!
 iGroupCountAccel! = 0.05!

 SetMaxParticlesInc
 
 move_Slowness = 0.7! / StandardSpeed
 track_Slowness = 0.8! / StandardSpeed
 
 trackSlowMax = 2! / StandardSpeed
 trackSlowMin = 0.1! / StandardSpeed
 
 If track_Slowness < 1 Then track_Slowness = 1
 If trackSlowMin < 1 Then trackSlowMin = 1
 
 iTrackSlow = StandardSpeed * 4&
      
End Sub
Private Sub SetMaxParticlesInc()
 iMaxParticlesAccel! = StandardSpeed / 128&
 iMaxParticlesAccelA! = StandardSpeed / 32&
 iMaxParticles = iMaxParticlesAccel
End Sub
Private Sub SetSizesRateInc()
 iSizesRate = StandardSpeed / 500
 iSizesRateAccel = StandardSpeed / 750
 sizes_adjustRate = iSizesRate
End Sub
Private Sub LoopTestCode()
Dim N102&, N40&, A41&(100)
 
 For N102 = 1& To 2&

  For N40 = 1& To 4000
  A41(Rnd * 100) = Rnd * 1200
  radius = Sqr(A41(Rnd)) ^ 2
  Next
 
 Next N102
End Sub
Private Sub InitParticleGroups()
Dim IPLoop&
 
 If NumParticleGroups > GROUPS_MAX Then NumParticleGroups = GROUPS_MAX
 
 PGStack.GroupsUB = NumParticleGroups - 1
 
 ReDim PGroupsOrderArray(PGStack.GroupsUB)
 
 For IPLoop = 0& To PGStack.GroupsUB
  PGroupsOrderArray(IPLoop) = IPLoop
 Next
 
 ColorGroups
 
 For IPLoop = 0& To PGStack.GroupsUB
  LumenParticleGroup PGroups(IPLoop), 1.6! + Rnd / 2
 Next
  
 InitExplosives
 InitEmitRates
 
End Sub
Private Sub ColorGroups()
 'used by HueShift
 iR = s1Red
 iG = s1Grn
 iB = s1Blu
 intensity = s1Intensity
 
 'intensity adjusts the rainbow spread
 
 NormaliseSHF intensity, 6
 HueShift 'HueShift changes the values of iR, iG, iB
 
 If PGStack.GroupsUB > 0 Then
  intensity = s2Intensity / PGStack.GroupsUB
  NormaliseSHF intensity, 6
 End If
 
 For N = 0& To PGStack.GroupsUB
  With PGroups(N)
  .IPS.sBright = 255!
  .chBlue = iB
  .chGreen = iG
  .chRed = iR
  End With
  HueShift
 Next
End Sub
Private Sub CalcTick()

 'With, say, 200Ghz systems, Elapsed will = 0,
 'and nothing will move if calling CalcSpeed
 'in Render() in modMainRendering
 
 Tick = timeGetTime
 Elapsed = Tick - LastTic
 LastTic = Tick

End Sub
Private Sub CalcFPS()
 
 TickSum = TickSum + Elapsed
 If TickSum& > 1000& Then
  fps = 1000& * FrameCount / TickSum
  FrameCount = 0&
  TickSum = 0&
 End If

End Sub
Private Sub InitParticles()

 PStack.PeasUB = MaxParticles - 1
 
 PStack.PeaAvail.StackUB = PStack.PeasUB
 PStack.PeaInUse.StackUB = -1
 
 For N = 0& To PStack.PeasUB
  PStack.PeaAvail.Stack(N) = N
 Next
  
End Sub
Private Sub DefParticle(Pea As ParticleSprite, PGroup As ParticleGroup)

 Pea.iFrame = StandardSpeed / 4&
 Pea.maxFrame = durationsBase + Rnd * durationsVariance
 NFrames = Pea.maxFrame / Pea.iFrame
 Pea.sBlu = PGroup.chBlue
 Pea.sGrn = PGroup.chGreen
 Pea.sRed = PGroup.chRed
 Pea.def = PGroup.IPS.IProcess.def
 Pea.iLum = Pea.def / NFrames
 sineparm = Rnd * TwoPi
 radius = (Rnd + 0.1!) * PGroup.IPS.maxRadius / Pea.maxFrame
 delta_x = Cos(sineparm) * radius
 delta_y = Sin(sineparm) * radius
 Pea.ix = delta_x + PGroup.IPS.cdx / ParticleLightness&
 Pea.iy = delta_y + PGroup.IPS.cdy / ParticleLightness&
 radius = Rnd / 2&
 Pea.px = PGroup.IPS.centerX + Pea.ix * radius
 Pea.py = PGroup.IPS.centerY + Pea.iy * radius
 Pea.sFrame = 0!

End Sub
Private Sub InitParticleSizes(ByVal PeasUBound&)
 sing = ParticleSizes / SIZES_MAX
 PeasUBound = PeasUBound * 1.5!
 If PeasUBound > PStack.PeasUB Then PeasUBound = PStack.PeasUB
 For N = 0& To PeasUBound
  N1 = PStack.PeaInUse.Stack(N)
  DimensionParticle Peas(N1), 1! + sing * sParticleSizesLUT(N1)
 Next
End Sub
Private Sub InitParticleSizesLUT()
 For N = 0& To P_MAX_UB
  sParticleSizesLUT(N) = Rnd * SIZES_MAX! + 0.01!
 Next
End Sub
Private Sub InitExplosives()
 For N = 0& To PGStack.GroupsUB
  PGroups(N).IPS.maxRadius = StandardSpeed * (explosiveBase! + explosiveVariance * Rnd)
 Next
End Sub
Private Sub InitEmitRates()
 For N = 0& To PGStack.GroupsUB
  PGroups(N).emitRate = StandardSpeed * (emissionsBase + emissionsVariance * Rnd)
 Next
End Sub
Private Sub AddParticle(PGroup As ParticleGroup)
If PStack.PeaAvail.StackUB > -1& Then
 PStack.PeaInUse.StackUB = PStack.PeaInUse.StackUB + 1&
 PStack.PeaInUse.Stack(PStack.PeaInUse.StackUB) = _
  PStack.PeaAvail.Stack(PStack.PeaAvail.StackUB)
 PStack.PeaAvail.StackUB = PStack.PeaAvail.StackUB - 1&
 DefParticle Peas(PStack.PeaInUse.Stack(PStack.PeaInUse.StackUB)), PGroup
End If
End Sub
Public Sub DeleteParticle(Element&)
 'Increase the OffScreen bullet stack pointer
 PStack.PeaAvail.StackUB = PStack.PeaAvail.StackUB + 1&
 'Copy PGInUse.Stack(Element) to PGAvail.Stack(OffScrUBound)
 PStack.PeaAvail.Stack(PStack.PeaAvail.StackUB) = PStack.PeaInUse.Stack(Element)
 'Copy PGInUse.Stack(UBound) to PGInUse.Stack(Element)
 PStack.PeaInUse.Stack(Element) = PStack.PeaInUse.Stack(PStack.PeaInUse.StackUB)
 'Lower the PGInUse bullet stack pointer
 PStack.PeaInUse.StackUB = PStack.PeaInUse.StackUB - 1&
 Element = Element - 1&
End Sub
Private Sub InitStandardSpeedVariables()
Dim ISVLoop&, ssv1!

 For ISVLoop = 0& To PGStack.GroupsUB
  PGroups(ISVLoop).emitRate = StandardSpeed * (emissionsBase + emissionsVariance * Rnd)
  PGroups(ISVLoop).IPS.maxRadius = StandardSpeed * (explosiveBase! + explosiveVariance * Rnd)
 Next
 durationsBase = durationsBase / StandardSpeed / 16&
 durationsVariance = durationsVariance / StandardSpeed / 16&
  
 ssv1 = StandardSpeed / 100&
 sizes_adjustRate = (SIZES_MAX - SIZES_MIN) * ssv1
 SetSizesRateInc
 xploBase_adjustRate = (XPLOBASE_MAX - XPLOBASE_MIN) * ssv1
 xploVari_adjustRate = (XPLOVARI_MAX - XPLOVARI_MIN) * ssv1
 emsnBase_adjustRate = (EMSNBASE_MAX - EMSNBASE_MIN) * ssv1
 emsnVari_adjustRate = (EMSNVARI_MAX - EMSNVARI_MIN) * ssv1
 duraBase_adjustRate = (DURABASE_MAX - DURABASE_MIN) * ssv1 * 8
 duraVari_adjustRate = (DURAVARI_MAX - DURAVARI_MIN) * ssv1 * 8

 sShiftSpeed = ssv1 * 36&
 sSpreadSpeed = ssv1 / 10&
 
End Sub
Private Sub GenerateBackground(Surf As AnimSurf1D)
Dim XTrack&
Dim x_position!
Dim y_position!
Dim ix!
Dim iy!
Dim B&
Dim right_side!
Dim top_edge!
Dim left_side!
Dim bottom_edge!
Dim N&
   
  'position along scanline
  XTrack = 0
  
  left_side = 16& * (Rnd - 0.5)
  bottom_edge = 16& * (Rnd - 0.5)
  
  If Not bBackground2 Then
   right_side = left_side + SW * ((Rnd - 0.5) + 1.2) / 377
   top_edge = bottom_edge + SH * ((Rnd - 0.5) + 0.9) / 472
  Else
   right_side = left_side + SW * ((Rnd - 0.5) * 0.001 + 350.1) / 377
   top_edge = bottom_edge + SH * ((Rnd - 0.5) * 0.001 + 532.45) / 472
  End If

  ix = (right_side - left_side) / SW
  iy = (top_edge - bottom_edge) / SH
  
  x_position = left_side
  y_position = bottom_edge
  
  If Not bBackground2 Then

  For N& = 0& To Surf.UBSurf Step 1&
  
   B& = 25& * Sin(x_position! * y_position - Cos(x_position + y_position!)) + 55!
   
   Select Case CSeL
   Case 0
    Surf.Dib(N&).Blue = B&
    Surf.Dib(N&).Green = B&
    Surf.Dib(N&).Red = 0&
   Case 1
    Surf.Dib(N&).Blue = B&
    Surf.Dib(N&).Green = 0&
    Surf.Dib(N&).Red = 0&
   Case 2
    Surf.Dib(N&).Blue = B&
    Surf.Dib(N&).Green = 0&
    Surf.Dib(N&).Red = B&
   Case 3
    Surf.Dib(N&).Blue = 0&
    Surf.Dib(N&).Green = B&
    Surf.Dib(N&).Red = 0&
   Case 4
    Surf.Dib(N&).Blue = 0&
    Surf.Dib(N&).Green = B&
    Surf.Dib(N&).Red = B&
   Case 5
    Surf.Dib(N&).Blue = B&
    Surf.Dib(N&).Green = B&
    Surf.Dib(N&).Red = B&
   Case 6
    Surf.Dib(N&).Blue = 0&
    Surf.Dib(N&).Green = 0&
    Surf.Dib(N&).Red = B&
   End Select
   
   x_position! = x_position! + ix!
   
   XTrack& = XTrack& + 1&
   If XTrack& >= Surf.Dims.Width& Then
    x_position! = left_side!
    y_position! = y_position! + iy!
    XTrack& = 0&
   End If
   
  Next
  
  Else
  
  For N& = 0& To Surf.UBSurf Step 1
  
   B& = 57& * Sin(x_position! * y_position - Cos(x_position * y_position!)) + 75!
   
   Select Case CSeL
   Case 0
    Surf.Dib(N&).Blue = B&
    Surf.Dib(N&).Green = B&
    Surf.Dib(N&).Red = 0&
   Case 1
    Surf.Dib(N&).Blue = B&
    Surf.Dib(N&).Green = 0&
    Surf.Dib(N&).Red = 0&
   Case 2
    Surf.Dib(N&).Blue = B&
    Surf.Dib(N&).Green = 0&
    Surf.Dib(N&).Red = B&
   Case 3
    Surf.Dib(N&).Blue = 0&
    Surf.Dib(N&).Green = B&
    Surf.Dib(N&).Red = 0&
   Case 4
    Surf.Dib(N&).Blue = 0&
    Surf.Dib(N&).Green = B&
    Surf.Dib(N&).Red = B&
   Case 5
    Surf.Dib(N&).Blue = B&
    Surf.Dib(N&).Green = B&
    Surf.Dib(N&).Red = B&
   Case 6
    Surf.Dib(N&).Blue = 0&
    Surf.Dib(N&).Green = 0&
    Surf.Dib(N&).Red = B&
   End Select
   
   x_position! = x_position! + ix!
   
   XTrack& = XTrack& + 1&
   If XTrack& >= Surf.Dims.Width Then
    x_position! = left_side!
    y_position! = y_position! + iy!
    XTrack& = 0&
   End If
   
  Next
  End If
  
  'EraseBuf is what EraseSprites copies from to 'Erase'
  For N = 0 To Surf.UBSurf
   Surf.EraseDib(N) = Surf.LDib(N)
  Next
  
End Sub
Private Sub Form_Resize()

 SW = ScaleWidth
 SH = ScaleHeight
 HalfWidth! = SW / 2
 HalfHeight! = SH / 2
 
 InitAnimSurf1D Me, FormDib
 
 If SW > 0 And SH > 0 Then
  GenerateBackground FormDib
 End If
 
End Sub
Private Sub ParticleSizesOnce()
 sing = ParticleSizes / SIZES_MAX
 For N = 0& To P_MAX_UB
  DimensionParticle Peas(N), 1! + sing * sParticleSizesLUT(N)
  PStack.PeaInUse.Stack(N) = P_MAX_UB - N
 Next
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
blnRunning = False
End Sub

