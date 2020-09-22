Attribute VB_Name = "modMainRendering"
Option Explicit

'This module is pretty straight-forward

'-----------------------------------------
' This is a good spot for your variables
'-----------------------------------------

Public SW&        '& means  As Long
Public HalfWidth! '! means  As Single
Public SH&
Public HalfHeight!

Public Xmax&, Ymax&

Public blnRunning As Boolean
Public EraseBuf() As Byte 'also used by Blit subs
Public Loca&, Loca1&, Loca2&
Public StepX&
Public Last_Blue_Byte&
Public Last_Red_Byte&
Public ViewPort_Right&
Public ViewPort_TopLeft&
Public ViewPort_Right_Blue&

'Variable specific to this module
Dim Looper&

Public StandardSpeed As Single





