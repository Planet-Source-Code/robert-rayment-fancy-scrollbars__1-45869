VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   6270
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHS 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   2
      Left            =   2325
      ScaleHeight     =   420
      ScaleWidth      =   2430
      TabIndex        =   25
      Top             =   2265
      Width           =   2430
      Begin VB.CommandButton cmdHSTUB 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   60
         Width           =   450
      End
      Begin VB.CommandButton cmdHSRIT 
         BackColor       =   &H00FFC0C0&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   60
         Width           =   240
      End
      Begin VB.CommandButton cmdHSLEF 
         BackColor       =   &H00FFC0C0&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   45
         Width           =   225
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   375
      Index           =   2
      LargeChange     =   10
      Left            =   1800
      Max             =   10
      Min             =   -10
      TabIndex        =   24
      Top             =   1575
      Width           =   2565
   End
   Begin VB.PictureBox picVS 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   2085
      Index           =   1
      Left            =   5205
      ScaleHeight     =   2025
      ScaleWidth      =   390
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   915
      Width           =   450
      Begin VB.CommandButton cmdVSTOP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "±"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   6
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   15
         Width           =   285
      End
      Begin VB.CommandButton cmdVSBOT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   6
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1680
         Width           =   225
      End
      Begin VB.CommandButton cmdVSTUB 
         BackColor       =   &H80000000&
         Height          =   420
         Index           =   1
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   720
         Width           =   165
      End
   End
   Begin VB.VScrollBar VS 
      Height          =   3795
      Index           =   1
      LargeChange     =   10
      Left            =   4920
      Max             =   -10
      Min             =   100
      SmallChange     =   5
      TabIndex        =   18
      Top             =   150
      Value           =   -10
      Width           =   330
   End
   Begin VB.PictureBox picVS 
      BackColor       =   &H00000000&
      Height          =   2085
      Index           =   0
      Left            =   1125
      ScaleHeight     =   2025
      ScaleWidth      =   810
      TabIndex        =   13
      Top             =   2010
      Width           =   870
      Begin VB.CommandButton cmdVSTUB 
         DownPicture     =   "Main.frx":0000
         Height          =   420
         Index           =   0
         Left            =   135
         Picture         =   "Main.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   720
         Width           =   360
      End
      Begin VB.CommandButton cmdVSBOT 
         BackColor       =   &H00FFC0FF&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1695
         Width           =   345
      End
      Begin VB.CommandButton cmdVSTOP 
         BackColor       =   &H00FFC0FF&
         Caption         =   "±"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   45
         Width           =   375
      End
   End
   Begin VB.VScrollBar VS 
      Height          =   2025
      Index           =   0
      Left            =   660
      Max             =   10
      TabIndex        =   12
      Top             =   1425
      Width           =   525
   End
   Begin VB.PictureBox picHS 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   1
      Left            =   1665
      ScaleHeight     =   420
      ScaleWidth      =   2430
      TabIndex        =   7
      Top             =   1110
      Width           =   2430
      Begin VB.CommandButton cmdHSLEF 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Î"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   45
         Width           =   225
      End
      Begin VB.CommandButton cmdHSRIT 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ï"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   60
         Width           =   240
      End
      Begin VB.CommandButton cmdHSTUB 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1335
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   60
         Width           =   225
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   525
      Index           =   1
      LargeChange     =   10
      Left            =   1470
      Max             =   10
      Min             =   -10
      TabIndex        =   6
      Top             =   690
      Width           =   1980
   End
   Begin VB.PictureBox picHS 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   255
      ScaleHeight     =   315
      ScaleWidth      =   2430
      TabIndex        =   1
      Top             =   330
      Width           =   2430
      Begin VB.CommandButton cmdHSTUB 
         BackColor       =   &H00C0E0FF&
         Height          =   225
         Index           =   0
         Left            =   1335
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   60
         Width           =   225
      End
      Begin VB.CommandButton cmdHSRIT 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1980
         Picture         =   "Main.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
      End
      Begin VB.CommandButton cmdHSLEF 
         BackColor       =   &H00FFFF80&
         Height          =   240
         Index           =   0
         Left            =   30
         Picture         =   "Main.frx":12DE
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   45
         Width           =   225
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   180
      Index           =   0
      LargeChange     =   10
      Left            =   645
      Max             =   100
      SmallChange     =   2
      TabIndex        =   0
      Top             =   195
      Width           =   2400
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   285
      Index           =   2
      Left            =   2715
      TabIndex        =   29
      Top             =   2025
      Width           =   690
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   23
      Top             =   3660
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      Height          =   285
      Index           =   0
      Left            =   570
      TabIndex        =   17
      Top             =   3540
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   285
      Index           =   1
      Left            =   3495
      TabIndex        =   11
      Top             =   825
      Width           =   690
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   285
      Index           =   0
      Left            =   3150
      TabIndex        =   4
      Top             =   135
      Width           =   630
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form1  (Main.frm)

' Fancy Scrollbars by Robert Rayment June 2003

Option Explicit

' With optional timing API GetTickCount
' and optional shaping API

' The method used is to put a VB ScrollBar on the form fixing
' final position, Max, Min & SmallChange.  LargeChange is not
' used.  Then a picture box containing 3 command buttons. The
' prog will then overlay the VB ScrollBar matching the size
' and position. 6 arrays are set up for each scrollbar for
' retrieving info from the VB ScrollBar and setting scale
' factors. The VB ScrollBars and the controls must be indexed.

' NB This assumes the Form Scalemode = Pixels or Twips

' On Form put

' Form1  Horz VB ScrollBar HS(0), HS(1) etc
'        Container: picHS(0), picHS(1) etc (twips)
'        Controls:  Left,       Thumb,     Right
'        in picHS  cmdHSLEF(0), cmdHSTUB(0), cmdHSRIT(0)
'                  cmdHSLEF(1), cmdHSTUB(1), cmdHSRIT(1)
'                              etc
'        Output values in HSVal(0), HSVal(1) etc

' Form1  Vert VB ScrollBar VS(0), VS(1) etc
'        Container: picVS(0), picVS(1) etc (twips)
'        Controls:  Top,        Thumb,      Bottom
'        in picVS  cmdVSTOP(0), cmdVSTUB(0), cmdVSBOT(0)
'                  cmdVSTOP(1), cmdVSTUB(1), cmdVSBOT(1)
'                              etc
'        Output values in VSVal(0), VSVal(1) etc

' NB These same Control Names must be used.  Also even
'    for one scrollbar it must have an index (ie 0).

' Things that could be added: include LargeChange &
' adjust thumb size according to range of values.

' A full user control could be made (though I'm not
' practised at that). It would enable some properties
' to be more easily set but I think not so flexible
' for the optional items used here, eg like shaping
' APIs.  There is a scrollbar OCX on PSC at CodeID =
' 22788 but is more complicated than the method used
' here.



Private Sub Form_Load()

Dim Spare As Long
Dim j As Long
   
   Me.Caption = " Fancy Scrollbars  by Robert Rayment"
   
   '-----------------------------------------------------
   ' FOR THREE HORIZONTAL SCROLLBARS  HS(0), HS(1) & HS(2)
   
   ' REQUIRED
   
   ReDim zHSMin(2), zHSMax(2)
   ReDim zHSSlope(2), zHSCut(2)
   ReDim HSVal(2)
   ReDim aHSSwapMinMax(2)
   
   ' SCALE also sets STX,STY TwipsPerPixel
   SCALE_HorzScrollbars Me, HS(0), 0
   SCALE_HorzScrollbars Me, HS(1), 1
   SCALE_HorzScrollbars Me, HS(2), 2
   
   ' OPTIONAL STUFF
   
   ' Show start values as Max < or > Min - Labels - optional
   Label2(0).Caption = HS(0).Min
   If aHSSwapMinMax(0) Then Label2(0).Caption = HS(0).Max    ' ie Swapped
   Label2(1).Caption = HS(1).Min
   If aHSSwapMinMax(1) Then Label2(1).Caption = HS(1).Max    ' ie Swapped
   Label2(2).Caption = HS(2).Min
   If aHSSwapMinMax(1) Then Label2(2).Caption = HS(2).Max    ' ie Swapped
   
   ' Shape thumb on 2nd horizontal scrollbar - optional
   Spare = CreateRoundRectRgn(0, 2, 13, 34, 8, 10)
   SetWindowRgn cmdHSTUB(1).hWnd, Spare, True
   DeleteObject Spare
   
   ' Wingdings arrows - optional
   cmdHSLEF(1).Caption = Chr$(215)
   cmdHSRIT(1).Caption = Chr$(216)
   
   ' Draw bars on picHS(2) - optional
   picHS(2).AutoRedraw = True
   Spare = 255
   For j = 30 To 150
      picHS(2).Line (0, j)-(picHS(2).Width, j), RGB(Spare, Spare, Spare)
      picHS(2).Line (0, j + 180)-(picHS(2).Width, j + 180), RGB(Spare, Spare, Spare)
      Spare = Spare - 1
   Next j
   
   '-------------------------------------------
   ' FOR TWO VERICAL SCROLLBARS  VS(0) & VS(1)
   
   ' REQUIRED
   
   ReDim zVSMin(1), zVSMax(1)
   ReDim zVSSlope(1), zVSCut(1)
   ReDim VSVal(1)
   ReDim aVSSwapMinMax(1)
   
   SCALE_VertScrollbars Me, VS(0), 0
   SCALE_VertScrollbars Me, VS(1), 1
   
   ' OPTIONAL STUFF
   
   ' Show start values as Max < or > Min - Labels - optional
   Label3(0).Caption = VS(0).Min
   If aVSSwapMinMax(0) Then Label3(0).Caption = VS(0).Max    ' ie Swapped
   Label3(1).Caption = VS(1).Min
   If aVSSwapMinMax(1) Then Label3(1).Caption = VS(1).Max    ' ie Swapped
   
   ' Wingdings arrows - optional
   cmdVSTOP(0).Caption = Chr$(241)
   cmdVSBOT(0).Caption = Chr$(242)
   
   cmdVSTOP(1).Caption = Chr$(217)
   cmdVSBOT(1).Caption = Chr$(218)
   
   ' Put lines on 2nd vertical scrollbar - optional
   picVS(1).AutoRedraw = True
   For j = picVS(1).Top To picVS(1).Height * STY Step 60
      picVS(1).Line (0, j)-(picVS(1).Width * STX, j), 0
   Next j
   '-------------------------------------------

   Me.Show

   cmdHSTUB(0).SetFocus

End Sub

'#### HORIZONTAL SCROLLBARS ### All these needed for horizontal scrollbars ###################################
'#### though there are some options in Private Sub HS_Change                ##################################

Private Sub HS_Change(Index As Integer)
   HSVal(Index) = HS(Index).Value
   
   If aHSSwapMinMax(Index) Then     ' ie Swapped
      HSVal(Index) = -HSVal(Index) + (HS(Index).Max + HS(Index).Min)
   End If
   
   Label2(Index).Caption = HSVal(Index) ' - Labeling optional to show values
   
   ' Special thumb numbering - optional
   If Index = 2 Then cmdHSTUB(2).Caption = Str$(HSVal(Index))

End Sub

Private Sub cmdHSLEF_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   HSLEF_MouseDown HS(Index), cmdHSTUB(Index), Index
End Sub
Private Sub cmdHSLEF_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Cancel_Loop
   picHS(Index).SetFocus
End Sub

Private Sub cmdHSRIT_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   HSRIT_MouseDown HS(Index), cmdHSTUB(Index), Index
End Sub
Private Sub cmdHSRIT_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Cancel_Loop
   picHS(Index).SetFocus
End Sub

Private Sub cmdHSTUB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   HSThumb_MouseDown Button, X, Y
End Sub
Private Sub cmdHSTUB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   HSThumb_MouseMove HS(Index), cmdHSTUB(Index), Index, Button, X, Y
End Sub
Private Sub cmdHSTUB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   picHS(Index).SetFocus
End Sub

Private Sub picHS_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   HSPic_MouseDown HS(Index), cmdHSTUB(Index), Index, X, Y
End Sub


'#### VERTICAL SCROLLBARS ### All these needed for vertical scroll bars ######################################
'#### though there is an option in Private Sub VS_Change                ######################################

Private Sub VS_Change(Index As Integer)
   VSVal(Index) = VS(Index).Value
   
   If aVSSwapMinMax(Index) Then     ' ie Swapped
      VSVal(Index) = -VSVal(Index) + (VS(Index).Max + VS(Index).Min)
   End If
   
   Label3(Index).Caption = VSVal(Index) ' - Labeling optional to show values
End Sub


Private Sub cmdVSTOP_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   VSTOP_MouseDown VS(Index), cmdVSTUB(Index), Index
End Sub
Private Sub cmdVSTOP_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Cancel_Loop
   picVS(Index).SetFocus
End Sub

Private Sub cmdVSBOT_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   VSBOT_MouseDown VS(Index), cmdVSTUB(Index), Index
End Sub
Private Sub cmdVSBOT_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Cancel_Loop
   picVS(Index).SetFocus
End Sub

Private Sub cmdVSTUB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   VSThumb_MouseDown Button, X, Y
End Sub
Private Sub cmdVSTUB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   VSThumb_MouseMove VS(Index), cmdVSTUB(Index), Index, Button, X, Y
End Sub
Private Sub cmdVSTUB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   picVS(Index).SetFocus
End Sub

Private Sub picVS_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   VSPic_MouseDown VS(Index), cmdVSTUB(Index), Index, X, Y
End Sub

'###################################################################################
'Trebor Tnemyar


