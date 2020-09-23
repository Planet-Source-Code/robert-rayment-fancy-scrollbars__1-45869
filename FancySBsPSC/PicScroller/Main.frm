VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   5310
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   ScaleHeight     =   328
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   354
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      Height          =   3810
      Left            =   855
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   10
      Top             =   255
      Width           =   4125
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3660
         Left            =   15
         ScaleHeight     =   244
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   237
         TabIndex        =   11
         Top             =   15
         Width           =   3555
      End
   End
   Begin VB.PictureBox picVS 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      Height          =   2085
      Index           =   0
      Left            =   -30
      ScaleHeight     =   2025
      ScaleWidth      =   390
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1170
      Width           =   450
      Begin VB.CommandButton cmdVSTOP 
         BackColor       =   &H0000C000&
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
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   15
         Width           =   285
      End
      Begin VB.CommandButton cmdVSBOT 
         BackColor       =   &H0000C000&
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
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1680
         Width           =   225
      End
      Begin VB.CommandButton cmdVSTUB 
         BackColor       =   &H0000C000&
         Height          =   225
         Index           =   0
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   225
      End
   End
   Begin VB.VScrollBar VS 
      Height          =   3795
      Index           =   0
      LargeChange     =   10
      Left            =   285
      Max             =   -10
      Min             =   100
      TabIndex        =   5
      Top             =   330
      Value           =   -10
      Width           =   330
   End
   Begin VB.PictureBox picHS 
      BackColor       =   &H00008000&
      Height          =   315
      Index           =   0
      Left            =   1590
      ScaleHeight     =   255
      ScaleWidth      =   2370
      TabIndex        =   1
      Top             =   4380
      Width           =   2430
      Begin VB.CommandButton cmdHSTUB 
         BackColor       =   &H0000C000&
         Height          =   225
         Index           =   0
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   60
         Width           =   225
      End
      Begin VB.CommandButton cmdHSRIT 
         BackColor       =   &H0000C000&
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
         Index           =   0
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   -15
         Width           =   240
      End
      Begin VB.CommandButton cmdHSLEF 
         BackColor       =   &H0000C000&
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
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   45
         Width           =   225
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   345
      Index           =   0
      LargeChange     =   10
      Left            =   990
      Max             =   100
      TabIndex        =   0
      Top             =   4215
      Width           =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form1  (Main.frm)

' PicScroller using
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


Dim j As Long

Private Sub Form_Load()
   
   Me.Caption = " Pic Scroller  by Robert Rayment"
   
   ' Picture1 contains Picture2
   ' Make Picture2 > Picture1
   Picture2.Width = 2 * Picture1.Width
   Picture2.Height = 2 * Picture1.Height
   Picture2.Top = 0
   Picture2.Left = 0
   
   DrawSillyPicture
   
   ' Line up VB Scrollbars with Picture1
   VS(0).Top = Picture1.Top
   VS(0).Height = Picture1.Height
   VS(0).Left = Picture1.Left - VS(0).Width - 10
   ' Wingdings arrows
   cmdVSTOP(0).Caption = Chr$(217)
   cmdVSBOT(0).Caption = Chr$(218)
   
   HS(0).Left = Picture1.Left
   HS(0).Width = Picture1.Width
   HS(0).Top = Picture1.Top + Picture1.Height + 10
   HS(0).Height = VS(0).Width
   ' Wingdings arrows
   cmdHSLEF(0).Caption = Chr$(215)
   cmdHSRIT(0).Caption = Chr$(216)
   
   
   '-----------------------------------------------------
   ' FOR ONE HORIZONTAL SCROLLBARS  HS(0)
   
   ' REQUIRED
   
   ReDim zHSMin(1), zHSMax(1)
   ReDim zHSSlope(1), zHSCut(1)
   ReDim HSVal(1)
   ReDim aHSSwapMinMax(1)
   
   HS(0).Min = 0
   HS(0).SmallChange = 2
   HS(0).Max = Picture2.Width - (Picture1.Width - 4)
   ' -4 for Picture1 border
   
   ' SCALE also sets STX,STY TwipsPerPixel
   SCALE_HorzScrollbars Me, HS(0), 0
   
   
   '-------------------------------------------
   ' FOR ONE VERICAL SCROLLBARS  VS(0)
   
   ' REQUIRED
   
   ReDim zVSMin(1), zVSMax(1)
   ReDim zVSSlope(1), zVSCut(1)
   ReDim VSVal(1)
   ReDim aVSSwapMinMax(1)
   
   VS(0).Min = 0
   VS(0).SmallChange = 2
   VS(0).Max = Picture2.Height - (Picture1.Height - 4)
   ' -4 for Picture1 border
   
   SCALE_VertScrollbars Me, VS(0), 0
   
   '-------------------------------------------

   Me.Show

   cmdHSTUB(0).SetFocus

End Sub

Private Sub DrawSillyPicture()
   Dim cul1 As Long
   Dim cul2 As Long

   Randomize

   For j = 0 To Picture2.Height Step 2
      
      cul1 = RGB(Rnd * 255, 180, 200)
      cul2 = RGB(Rnd * 255, 200, 180)
      
      If (Rnd > 0.8) Then
         Picture2.Line (0, j)-(Picture2.Width, j), cul1
      Else
         Picture2.Line (0, j)-(Picture2.Width, j + 9), cul2
      End If
   
      If (Rnd > 0.8) Then
         Picture2.Line (j, 0)-(j, Picture2.Height), cul1
      Else
         Picture2.Line (j, 0)-(j + 9, Picture2.Height), cul2
      End If
   
   Next j
   
   ' Draw red boundary box to check out Scrollbar maxes
   Picture2.Line (0, 0)-(Picture2.Width - 1, Picture2.Height - 1), vbRed, B

End Sub


'#### HORIZONTAL SCROLLBARS ### All these needed for horizontal scrollbars ###################################
'#### though there are some options in Private Sub HS_Change                ##################################

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

Private Sub HS_Change(Index As Integer)
   HSVal(Index) = HS(Index).Value
   
   If aHSSwapMinMax(Index) Then     ' ie Swapped
      HSVal(Index) = -HSVal(Index) + (HS(Index).Max + HS(Index).Min)
   End If

   ' SPECIAL TO THIS PROG
   
   Picture2.Left = -HSVal(Index)

End Sub

'#### VERTICAL SCROLLBARS ### All these needed for vertical scroll bars ######################################
'#### though there is an option in Private Sub VS_Change                ######################################

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

Private Sub VS_Change(Index As Integer)
   VSVal(Index) = VS(Index).Value
   
   If aVSSwapMinMax(Index) Then     ' ie Swapped
      VSVal(Index) = -VSVal(Index) + (VS(Index).Max + VS(Index).Min)
   End If
   
   ' SPECIAL TO THIS PROG
   
   Picture2.Top = -VSVal(Index)

End Sub

'###################################################################################
'Trebor Tnemyar


