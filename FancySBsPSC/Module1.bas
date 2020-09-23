Attribute VB_Name = "Module1"
'Module1.bas

' Fancy Scrollbars  by Robert Rayment June 2003

Option Explicit

' On Form put

' Form1  Horz ScrollBar HS(0), HS(1) etc
'        Container: picHS(0), picHS(1) etc (twips)
'        Controls:  Left,       Thumb,     Right
'                  cmdHSLEF(0), cmdHSTUB(0), cmdHSRIT(0)
'                  cmdHSLEF(1), cmdHSTUB(1), cmdHSRIT(1)
'                              etc
'        Output values in HSVal(0), HSVal(1) etc

' Form1  Vert ScrollBar VS(0), VS(1) etc
'        Container: picVS(0), picVS(1) etc (twips)
'        Controls:  Top,        Thumb,      Bottom
'                  cmdVSTOP(0), cmdVSTUB(0), cmdVSBOT(0)
'                  cmdVSTOP(1), cmdVSTUB(1), cmdVSBOT(1)
'                              etc
'        Output values in VSVal(0), VSVal(1) etc

' NB These same Control Names must be used

' --------------------------------------------------------------
' Shaping APIs   CAN BE OMITTED IF NOT WANTED
Public Declare Function CreateRoundRectRgn Lib "gdi32" _
(ByVal X1 As Long, ByVal Y1 As Long, ByVal _
 X2 As Long, ByVal Y2 As Long, _
 ByVal X3 As Long, ByVal Y3 As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" _
(ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Declare Function DeleteObject Lib "gdi32" _
(ByVal hObject As Long) As Long

' --------------------------------------------------------------
' Timing - if removed then scrollbar buttons will need repeated pressing

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' Use:-

'   thumb.Refresh
'   DoEvents
'   Sleep TLIM
'   DoEvents

' both DoEvents seem necessary to ensure thumb redrawn
' and to pick up _MouseUp event to exit Do Loop.

' --------------------------------------------------------------

Public HSVal() As Long  ' Output horizontal scrollbar values
Public zHSMax() As Single, zHSMin()  As Single
Public zHSSlope()  As Single, zHSCut() As Single
Public aHSSwapMinMax() As Boolean   ' Flag to show if HS Max < Min
'-----------------------------------------------------
Public VSVal() As Long  ' Output vertical scrollbar values
Public zVSMax() As Single, zVSMin()  As Single
Public zVSSlope()  As Single, zVSCut() As Single
Public aVSSwapMinMax() As Boolean   ' Flag to show if VS Max < Min

'-----------------------------------------------------
Public STX As Long  '= Screen.TwipsPerPixelX
Public STY As Long  '= Screen.TwipsPerPixelY
'-----------------------------------------------------

Private ThumbPosition As Long
Private TOPY As Long, LEFTX  As Long

Private aDone As Boolean   ' Loop exit test

Private TLIM As Long  ' Timing
Private T As Long  'Spare


'#### HORIZONTAL SCROLLBARS #####################################################

Public Sub HSLEF_MouseDown(HScr As Control, thumb As CommandButton, Index As Integer)

' LEFT BUTTON

' Called:-
'   HSLEF_MouseDown HS(Index), cmdHSTUB(Index), Index

aDone = False
TLIM = 150

'HScr == HS(Index).Value

Do
   
   If HScr > HScr.Min + HScr.SmallChange Then
      HScr = HScr - HScr.SmallChange
      thumb.Left = (HScr - zHSCut(Index)) / zHSSlope(Index)
   Else
      HScr = HScr.Min
      thumb.Left = zHSMin(Index)
   End If
   
   ' Timer  ' NB Can be left out but Command button will need re-pressing
   thumb.Refresh
   DoEvents
   Sleep TLIM
   DoEvents

Loop Until aDone

End Sub


Public Sub HSRIT_MouseDown _
 (HScr As Control, thumb As CommandButton, Index As Integer)

' RIGHT BUTTON

' Called:-
'   HSRIT_MouseDown HS(Index), cmdHSTUB(Index), Index

'HScr == HS(Index).Value

aDone = False
TLIM = 150

Do
   
   If HScr < HScr.Max - HScr.SmallChange Then
      HScr = HScr + HScr.SmallChange
      thumb.Left = (HScr - zHSCut(Index)) / zHSSlope(Index)
   Else
      HScr = HScr.Max
      thumb.Left = zHSMax(Index)
   End If

   ' Timer  ' NB Can be left out but Command button will need re-pressing
   thumb.Refresh
   DoEvents
   Sleep TLIM
   DoEvents
   
   ' Wierd - optional
   'thumb.Top = 60 * (Rnd - 0.5)
   
Loop Until aDone
   
   'thumb.Top = 0

End Sub

Public Sub Cancel_Loop()
   aDone = True
End Sub

Public Sub HSThumb_MouseDown(Button As Integer, X As Single, Y As Single)

' Called:-
'   HSThumb_MouseDown Button, X, Y

   If Button <> vbLeftButton Then Exit Sub
   TOPY = Y
   LEFTX = X

End Sub

Public Sub HSThumb_MouseMove(HScr As Control, thumb As CommandButton, _
 Index As Integer, Button As Integer, X As Single, Y As Single)

' Called:-
'  HSThumb_MouseMove HS(Index), cmdHSTUB(Index), Index, Button, X, Y

If Button = vbLeftButton Then
      
   thumb.Top = 0
   
   ' Check position of thumb
   ThumbPosition = thumb.Left + X - LEFTX
   Select Case ThumbPosition
   Case Is > zHSMax(Index)
      thumb.Left = zHSMax(Index)
      HSVal(Index) = HScr.Max
   Case Is < zHSMin(Index)
      thumb.Left = zHSMin(Index)
      HSVal(Index) = HScr.Min
   Case Else
      thumb.Left = thumb.Left + X - LEFTX
      HSVal(Index) = thumb.Left * zHSSlope(Index) + zHSCut(Index)
   End Select
   
   HScr = HSVal(Index)
   
End If

End Sub

Public Sub HSPic_MouseDown(HScr As Control, thumb As CommandButton, _
 Index As Integer, X As Single, Y As Single)

' Called:-
' HSPic_MouseDown HS(Index), cmdHSTUB(Index), Index, X, Y
  
If X > (thumb.Left + thumb.Width) Then
   
   If X < (zHSMax(Index) + thumb.Width) Then
      
      thumb.Left = X
      HSVal(Index) = X * zHSSlope(Index) + zHSCut(Index)
      
      If HSVal(Index) <= HScr.Max Then
         HScr = HSVal(Index)
      Else
         HScr = HScr.Max  ' Also sets HSVal(Index)
         thumb.Left = zHSMax(Index)
      End If
      
   End If

ElseIf X < thumb.Left Then
   
   If X > zHSMin(Index) Then
      
      thumb.Left = X
      HSVal(Index) = X * zHSSlope(Index) + zHSCut(Index)
      
      If HSVal(Index) >= HScr.Min Then
         HScr = HSVal(Index)
      Else
         HScr = HScr.Min  ' Also sets HSVal(Index)
         thumb.Left = zHSMin(Index)
      End If
   
   End If

End If

End Sub


Public Sub SCALE_HorzScrollbars(frm As Form, HScr As Control, Index As Integer)
   
' Form1 Container: picHS(Index) (twips)
'       Controls:  Left,           Thumb,         Right
'                  cmdHSLEF(Index), cmdHSTUB(Index, cmdHSRIT(INdex)
   
   
   If frm.ScaleMode = vbPixels Then
      STX = Screen.TwipsPerPixelX
      STY = Screen.TwipsPerPixelY
   ElseIf frm.ScaleMode = vbTwips Then
      STX = 1
      STY = 1
   Else
      MsgBox "Form ScaleMode not Pixels or Twips", vbCritical, "Fancy Scrollbars"
      Unload frm
      End
   End If
   '-------------------------------------------------------
With frm
   
   ' Scale picHS(Index) & command buttons to HS size
   ' Container
   .picHS(Index).Top = .HS(Index).Top
   .picHS(Index).Left = .HS(Index).Left
   .picHS(Index).Width = .HS(Index).Width
   .picHS(Index).Height = .HS(Index).Height

   ' Left button
   .cmdHSLEF(Index).Top = 0
   .cmdHSLEF(Index).Left = 0
   .cmdHSLEF(Index).Width = 240   ' Fixed width of end buttons
   .cmdHSLEF(Index).Height = .picHS(Index).Height * STY

   ' Right button
   .cmdHSRIT(Index).Top = 0
   .cmdHSRIT(Index).Left = .picHS(Index).Width * STX - 255   ' Tweak
   .cmdHSRIT(Index).Width = 240
   .cmdHSRIT(Index).Height = .picHS(Index).Height * STY

   ' Horz Thumb
   .cmdHSTUB(Index).Top = 0
   .cmdHSTUB(Index).Left = .cmdHSLEF(Index).Width   ' Thumb @ Left pos
   .cmdHSTUB(Index).Height = .picHS(Index).Height * STY
   '------------------------------------------------------
   ' Hide HS - still operative
   .HS(Index).Visible = False
   
   ' Calc Thumb's (cmdHSTUB(Index)'s) max & min location on picHS(Index)
   ' & zHSSlope, zHSCut to HS max & min
   
   aHSSwapMinMax(Index) = False ' Not swapped
   ' Check if Min > Max
   If HScr.Min > HScr.Max Then
      aHSSwapMinMax(Index) = True
      T = HScr.Min
      HScr.Min = HScr.Max
      HScr.Max = T
   End If
   
   zHSMin(Index) = .cmdHSLEF(Index).Width
   zHSMax(Index) = .cmdHSRIT(Index).Left - .cmdHSTUB(Index).Width
   zHSSlope(Index) = (.HS(Index).Max - .HS(Index).Min) / (zHSMax(Index) - zHSMin(Index))
   zHSCut(Index) = .HS(Index).Max - zHSSlope(Index) * zHSMax(Index)

   HScr = HScr.Min

End With

End Sub


'#### VERTICAL SCROLLBARS #####################################################

Public Sub VSTOP_MouseDown(VScr As Control, thumb As CommandButton, Index As Integer)

' TOP BUTTON

' Called:-
'   VSTOP_MouseDown VS(Index), cmdVSTUB(Index), Index

aDone = False
TLIM = 150

' VScr == VS(Index).Value

Do
   
   If VScr > VScr.Min + VScr.SmallChange Then
      VScr = VScr - VScr.SmallChange
      thumb.Top = (VScr - zVSCut(Index)) / zVSSlope(Index)
   Else
      VScr = VScr.Min
      thumb.Top = zVSMin(Index)
   End If
   
   ' Timer  ' NB Can be left out but Command button will need re-pressing
   thumb.Refresh
   DoEvents
   Sleep TLIM
   DoEvents

Loop Until aDone

End Sub


Public Sub VSBOT_MouseDown(VScr As Control, thumb As CommandButton, Index As Integer)

' BOTTOM BUTTON

' Called:-
'   VSBOT_MouseDown VS(Index), cmdVSTUB(Index), Index

'VScr == VS(Index).Value

aDone = False
TLIM = 150

Do
   
   If VScr < VScr.Max - VScr.SmallChange Then
      VScr = VScr + VScr.SmallChange
      thumb.Top = (VScr - zVSCut(Index)) / zVSSlope(Index)
   Else
      VScr = VScr.Max
      thumb.Top = zVSMax(Index)
   End If

     ' Timer  ' NB Can be left out but Command button will need re-pressing
   thumb.Refresh
   DoEvents
   Sleep TLIM
   DoEvents

Loop Until aDone

End Sub

Public Sub VSThumb_MouseDown(Button As Integer, X As Single, Y As Single)

' Called:-
'   VSThumb_MouseDown Button, X, Y
   
   If Button <> 1 Then Exit Sub
   TOPY = Y
   LEFTX = X

End Sub

Public Sub VSThumb_MouseMove(VScr As Control, thumb As CommandButton, _
 Index As Integer, Button As Integer, X As Single, Y As Single)

' Called:-
'   VSThumb_MouseMove VS(Index), cmdVSTUB(Index), Index, Button, X, Y

If Button = vbLeftButton Then
      
   thumb.Left = 0
   
   ' Check position of thumb
   ThumbPosition = thumb.Top + Y - TOPY
   Select Case ThumbPosition
   Case Is > zVSMax(Index)
      thumb.Top = zVSMax(Index)
      VSVal(Index) = VScr.Max
   Case Is < zVSMin(Index)
      thumb.Top = zVSMin(Index)
      VSVal(Index) = VScr.Min
   Case Else
      thumb.Top = thumb.Top + Y - TOPY
      VSVal(Index) = thumb.Top * zVSSlope(Index) + zVSCut(Index)
   End Select
   
   VScr = VSVal(Index)

End If

End Sub

Public Sub VSPic_MouseDown(VScr As Control, thumb As CommandButton, _
 Index As Integer, X As Single, Y As Single)

' Called:-
'   VSPic_MouseDown VS(Index), cmdVSTUB(Index), Index, X, Y
 
If Y > (thumb.Top + thumb.Height) Then
   
   If Y < (zVSMax(Index) + thumb.Height) Then
      
      thumb.Top = Y
      VSVal(Index) = Y * zVSSlope(Index) + zVSCut(Index)
      
      If VSVal(Index) <= VScr.Max Then
         VScr = VSVal(Index)
      Else
         VScr = VScr.Max  ' Also sets VSVal(Index)
         thumb.Top = zVSMax(Index)
      End If
      
   End If

ElseIf Y < thumb.Top Then
   
   If Y > zVSMin(Index) Then
      
      thumb.Top = Y
      VSVal(Index) = Y * zVSSlope(Index) + zVSCut(Index)
      
      If VSVal(Index) >= VScr.Min Then
         VScr = VSVal(Index)
      Else
         VScr = VScr.Min  ' Also sets VSVal(Index)
         thumb.Top = zVSMin(Index)
      End If
   
   End If

End If

End Sub

Public Sub SCALE_VertScrollbars(frm As Form, VScr As Control, Index As Integer)
   
' Form1 Container: picVS(Index) (twips)
'       Controls:  Top,             Thumb,          Bottom
'                  cmdVSTOP(Index), cmdVSTUB(Index), cmdVSBOT(Index)
   
   If frm.ScaleMode = vbPixels Then
      STX = Screen.TwipsPerPixelX
      STY = Screen.TwipsPerPixelY
   ElseIf frm.ScaleMode = vbTwips Then
      STX = 1
      STY = 1
   Else
      MsgBox "Form ScaleMode not Pixels or Twips", vbCritical, "Fancy Scrollbars"
      Unload frm
      End
   End If
   '-------------------------------------------------------
With frm
   
   ' Scale picVS(Index) & command buttons to VS size
   ' Container
   .picVS(Index).Top = .VS(Index).Top
   .picVS(Index).Left = .VS(Index).Left
   .picVS(Index).Width = .VS(Index).Width
   .picVS(Index).Height = .VS(Index).Height

   ' Top button
   .cmdVSTOP(Index).Top = 0
   .cmdVSTOP(Index).Left = 0
   .cmdVSTOP(Index).Width = .picVS(Index).Width * STX
   .cmdVSTOP(Index).Height = 240

   ' Bottom button
   .cmdVSBOT(Index).Top = .picVS(Index).Height * STY - 280 ' Tweak
   .cmdVSBOT(Index).Left = 0
   .cmdVSBOT(Index).Width = .picVS(Index).Width * STY
   .cmdVSBOT(Index).Height = 240

   ' Vert Thumb
   .cmdVSTUB(Index).Top = .cmdVSTOP(Index).Height   ' Thumb top pos
   .cmdVSTUB(Index).Left = 0
   .cmdVSTUB(Index).Width = .picVS(Index).Width * STY
   '------------------------------------------------------
   ' Hide VS - still operative
   .VS(Index).Visible = False
   
   ' Calc Thumb's (cmdVSTUB(Index)'s) max & min location on picVS(Index)
   ' & zVSSlope, zVSCut to VS max & min
   
   aVSSwapMinMax(Index) = False ' Not swapped
   ' Check if Min > Max
   If VScr.Min > VScr.Max Then
      aVSSwapMinMax(Index) = True
      T = VScr.Min
      VScr.Min = VScr.Max
      VScr.Max = T
   End If
   
   zVSMin(Index) = .cmdVSTOP(Index).Height
   zVSMax(Index) = .cmdVSBOT(Index).Top - .cmdVSTUB(Index).Height
   zVSSlope(Index) = (.VS(Index).Max - .VS(Index).Min) / (zVSMax(Index) - zVSMin(Index))
   zVSCut(Index) = .VS(Index).Max - zVSSlope(Index) * zVSMax(Index)

   VScr = VScr.Min

End With

End Sub

'Trebor Tnemyar

