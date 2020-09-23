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
   Begin VB.PictureBox picVS 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   825
      Index           =   2
      Left            =   3240
      ScaleHeight     =   765
      ScaleWidth      =   390
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2805
      Width           =   450
      Begin VB.PictureBox picVSTUB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   2
         Left            =   0
         ScaleHeight     =   120
         ScaleWidth      =   405
         TabIndex        =   35
         Top             =   330
         Width           =   435
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
         Index           =   2
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   555
         Width           =   225
      End
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
         Index           =   2
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   15
         Width           =   285
      End
   End
   Begin VB.VScrollBar VS 
      Height          =   510
      Index           =   2
      LargeChange     =   2
      Left            =   2550
      Max             =   0
      Min             =   10
      TabIndex        =   25
      Top             =   2955
      Width           =   495
   End
   Begin VB.PictureBox picHS 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   2
      Left            =   2325
      ScaleHeight     =   420
      ScaleWidth      =   2430
      TabIndex        =   21
      Top             =   2265
      Width           =   2430
      Begin VB.PictureBox picHSTUB 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   840
         ScaleHeight     =   195
         ScaleWidth      =   270
         TabIndex        =   32
         Top             =   45
         Width           =   330
      End
      Begin VB.CommandButton cmdHSRIT 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   60
         Width           =   240
      End
      Begin VB.CommandButton cmdHSLEF 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   45
         Width           =   225
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   270
      Index           =   2
      LargeChange     =   10
      Left            =   1800
      Max             =   10
      Min             =   -10
      TabIndex        =   20
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   915
      Width           =   450
      Begin VB.PictureBox picVSTUB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   1
         Left            =   45
         ScaleHeight     =   390
         ScaleWidth      =   300
         TabIndex        =   34
         Top             =   915
         Width           =   330
      End
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   1680
         Width           =   225
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
      TabIndex        =   15
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
      TabIndex        =   11
      Top             =   2010
      Width           =   870
      Begin VB.PictureBox picVSTUB 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   0
         Left            =   195
         ScaleHeight     =   390
         ScaleWidth      =   405
         TabIndex        =   33
         Top             =   645
         Width           =   435
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   45
         Width           =   375
      End
   End
   Begin VB.VScrollBar VS 
      Height          =   2025
      Index           =   0
      Left            =   660
      Max             =   10
      TabIndex        =   10
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
      TabIndex        =   6
      Top             =   1110
      Width           =   2430
      Begin VB.PictureBox picHSTUB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   990
         ScaleHeight     =   225
         ScaleWidth      =   300
         TabIndex        =   31
         Top             =   75
         Width           =   330
      End
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
         TabIndex        =   8
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   510
      Index           =   1
      LargeChange     =   10
      Left            =   1485
      Max             =   10
      Min             =   -10
      TabIndex        =   5
      Top             =   675
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
      Begin VB.PictureBox picHSTUB 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   930
         ScaleHeight     =   225
         ScaleWidth      =   300
         TabIndex        =   30
         Top             =   30
         Width           =   330
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
         Picture         =   "Main.frx":0000
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
         Picture         =   "Main.frx":014A
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      Height          =   285
      Index           =   2
      Left            =   2445
      TabIndex        =   29
      Top             =   3645
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   285
      Index           =   2
      Left            =   2715
      TabIndex        =   24
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
      TabIndex        =   19
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
      TabIndex        =   14
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
      TabIndex        =   9
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

' Fancy Scrollbars (pic Thumb) by Robert Rayment June 2003

' Alternative using a picture box for the scrollbar thumb
' instead of a command button.
' Can't easily use up/down icons on thumb, but better for
' up/down control appearance.

Option Explicit

' With optional timing API GetTickCount
' and optional shaping API

' The method used is to put a VB ScrollBar on the form fixing
' final position, Max, Min & SmallChange.  LargeChange is not
' used.  Then a picture box containing 2 command buttons & a
' picturebox thumb. The prog will then overlay the VB ScrollBar
' matching the size and position. 6 arrays are set up for each
' scrollbar for retrieving info from the VB ScrollBar and setting
' scale factors. The VB ScrollBars and the controls must be indexed.

' NB This assumes the Form Scalemode = Pixels or Twips

' On Form put

' Form1  Horz VB ScrollBar HS(0), HS(1) etc
'        Container: picHS(0), picHS(1) etc (twips)
'        Controls:  Left,       Thumb,     Right
'        in picHS  cmdHSLEF(0), picHSTUB(0), cmdHSRIT(0)
'                  cmdHSLEF(1), picHSTUB(1), cmdHSRIT(1)
'                              etc
'        Output values in HSVal(0), HSVal(1) etc

' Form1  Vert VB ScrollBar VS(0), VS(1) etc
'        Container: picVS(0), picVS(1) etc (twips)
'        Controls:  Top,        Thumb,      Bottom
'        in picVS  cmdVSTOP(0), picVSTUB(0), cmdVSBOT(0)
'                  cmdVSTOP(1), picVSTUB(1), cmdVSBOT(1)
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
   
   Me.Caption = " Fancy Scrollbars (pic Thumb) by Robert Rayment"
   
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
   Spare = CreateRoundRectRgn(0, 0, 23, 34, 22, 28)
   SetWindowRgn picHSTUB(1).hWnd, Spare, True
   DeleteObject Spare
   
   ' Wingdings arrows - optional
   cmdHSLEF(1).Caption = Chr$(215)
   cmdHSRIT(1).Caption = Chr$(216)
   
   picHSTUB(2).Cls
   picHSTUB(2).Print Str$(HSVal(2))
   '-------------------------------------------
   
   ' FOR THREE VERICAL SCROLLBARS  VS(0) & VS(1)
   
   ' REQUIRED
   
   ReDim zVSMin(2), zVSMax(2)
   ReDim zVSSlope(2), zVSCut(2)
   ReDim VSVal(2)
   ReDim aVSSwapMinMax(2)
   
   SCALE_VertScrollbars Me, VS(0), 0
   SCALE_VertScrollbars Me, VS(1), 1
   picVSTUB(2).Height = 4     ' Special to give Up/Dn appearance
   SCALE_VertScrollbars Me, VS(2), 2
   
   ' OPTIONAL STUFF
   
   ' Show start values as Max < or > Min - Labels - optional
   Label3(0).Caption = VS(0).Min
   If aVSSwapMinMax(0) Then Label3(0).Caption = VS(0).Max    ' ie Swapped
   Label3(1).Caption = VS(1).Min
   If aVSSwapMinMax(1) Then Label3(1).Caption = VS(1).Max    ' ie Swapped
   Label3(2).Caption = VS(2).Min
   If aVSSwapMinMax(2) Then Label3(2).Caption = VS(2).Max    ' ie Swapped
   
   ' Wingdings arrows - optional
   cmdVSTOP(0).Caption = Chr$(241)
   cmdVSBOT(0).Caption = Chr$(242)
   
   cmdVSTOP(1).Caption = Chr$(217)
   cmdVSBOT(1).Caption = Chr$(218)
   
   cmdVSTOP(2).Caption = Chr$(217)
   cmdVSBOT(2).Caption = Chr$(218)
   
   ' Put lines on 2nd vertical scrollbar - optional
   For j = picVS(1).Top To picVS(1).Height * STY Step 60
      picVS(1).Line (0, j)-(picVS(1).Width * STX, j), 0
   Next j
   ' NB for this picVS(1) must be AutoRedraw = True
   '-------------------------------------------

   Me.Show

   picHSTUB(0).SetFocus

End Sub

'#### HORIZONTAL SCROLLBARS ### All these needed for horizontal scrollbars ###################################
'#### though there are some options in Private Sub HS_Change                ##################################

Private Sub cmdHSLEF_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   HSLEF_MouseDown HS(Index), picHSTUB(Index), Index
End Sub
Private Sub cmdHSLEF_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Cancel_Loop
   picHS(Index).SetFocus
End Sub

Private Sub cmdHSRIT_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   HSRIT_MouseDown HS(Index), picHSTUB(Index), Index
End Sub
Private Sub cmdHSRIT_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Cancel_Loop
   picHS(Index).SetFocus
End Sub

Private Sub picHSTUB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   HSThumb_MouseDown Button, X, Y
End Sub
Private Sub picHSTUB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   HSThumb_MouseMove HS(Index), picHSTUB(Index), Index, Button, X, Y
End Sub
Private Sub picHSTUB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   picHS(Index).SetFocus
End Sub

Private Sub picHS_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   HSPic_MouseDown HS(Index), picHSTUB(Index), Index, X, Y
End Sub

Private Sub HS_Change(Index As Integer)
   HSVal(Index) = HS(Index).Value
   
   If aHSSwapMinMax(Index) Then     ' ie Swapped
      HSVal(Index) = -HSVal(Index) + (HS(Index).Max + HS(Index).Min)
   End If
   
   Label2(Index).Caption = HSVal(Index) ' - Labeling optional to show values
   
   ' Special thumb numbering - optional
   If Index = 2 Then
      picHSTUB(2).Cls
      picHSTUB(2).Print Str$(HSVal(Index))
   End If

End Sub

'#### VERTICAL SCROLLBARS ### All these needed for vertical scroll bars ######################################
'#### though there is an option in Private Sub VS_Change                ######################################

Private Sub cmdVSTOP_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   VSTOP_MouseDown VS(Index), picVSTUB(Index), Index
End Sub
Private Sub cmdVSTOP_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Cancel_Loop
   picVS(Index).SetFocus
End Sub

Private Sub cmdVSBOT_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   VSBOT_MouseDown VS(Index), picVSTUB(Index), Index
End Sub
Private Sub cmdVSBOT_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Cancel_Loop
   picVS(Index).SetFocus
End Sub

Private Sub picVSTUB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   VSThumb_MouseDown Button, X, Y
End Sub
Private Sub picVSTUB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   VSThumb_MouseMove VS(Index), picVSTUB(Index), Index, Button, X, Y
End Sub
Private Sub picVSTUB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   picVS(Index).SetFocus
End Sub

Private Sub picVS_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   VSPic_MouseDown VS(Index), picVSTUB(Index), Index, X, Y
End Sub

Private Sub VS_Change(Index As Integer)
   VSVal(Index) = VS(Index).Value
   
   If aVSSwapMinMax(Index) Then     ' ie Swapped
      VSVal(Index) = -VSVal(Index) + (VS(Index).Max + VS(Index).Min)
   End If
   
   Label3(Index).Caption = VSVal(Index) ' - Labeling optional to show values
End Sub

'###################################################################################
'Trebor Tnemyar


