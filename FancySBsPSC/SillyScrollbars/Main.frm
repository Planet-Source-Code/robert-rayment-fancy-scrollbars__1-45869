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
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHS 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      Height          =   420
      Index           =   3
      Left            =   915
      ScaleHeight     =   360
      ScaleWidth      =   2370
      TabIndex        =   19
      Top             =   3465
      Width           =   2430
      Begin VB.CommandButton cmdHSLEF 
         BackColor       =   &H00FFC0FF&
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
         Index           =   3
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   45
         Width           =   225
      End
      Begin VB.CommandButton cmdHSRIT 
         BackColor       =   &H00FFC0FF&
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
         Index           =   3
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   60
         Width           =   240
      End
      Begin VB.CommandButton cmdHSTUB 
         BackColor       =   &H00FFFFC0&
         DownPicture     =   "Main.frx":0000
         Height          =   315
         Index           =   3
         Left            =   990
         Picture         =   "Main.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   60
         Width           =   570
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   525
      Index           =   3
      LargeChange     =   10
      Left            =   540
      Max             =   10
      Min             =   -10
      TabIndex        =   18
      Top             =   3180
      Width           =   3330
   End
   Begin VB.PictureBox picHS 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   2
      Left            =   585
      ScaleHeight     =   420
      ScaleWidth      =   2430
      TabIndex        =   13
      Top             =   2700
      Width           =   2430
      Begin VB.CommandButton cmdHSTUB 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   2
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   60
         Width           =   450
      End
      Begin VB.CommandButton cmdHSRIT 
         BackColor       =   &H00FFC0FF&
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   60
         Width           =   240
      End
      Begin VB.CommandButton cmdHSLEF 
         BackColor       =   &H0080C0FF&
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   45
         Width           =   225
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   375
      Index           =   2
      LargeChange     =   10
      Left            =   540
      Max             =   10
      Min             =   -10
      TabIndex        =   12
      Top             =   2205
      Width           =   2565
   End
   Begin VB.PictureBox picHS 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   1
      Left            =   495
      ScaleHeight     =   420
      ScaleWidth      =   2430
      TabIndex        =   7
      Top             =   1695
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
      Left            =   645
      Max             =   10
      Min             =   -10
      TabIndex        =   6
      Top             =   1080
      Width           =   1980
   End
   Begin VB.PictureBox picHS 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   600
      Picture         =   "Main.frx":074C
      ScaleHeight     =   315
      ScaleWidth      =   2430
      TabIndex        =   1
      Top             =   645
      Width           =   2430
      Begin VB.CommandButton cmdHSTUB 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   1335
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   60
         Width           =   345
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
         Picture         =   "Main.frx":0D46
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
         Picture         =   "Main.frx":0E90
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   45
         Width           =   225
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   375
      Index           =   0
      LargeChange     =   10
      Left            =   600
      Max             =   100
      SmallChange     =   2
      TabIndex        =   0
      Top             =   150
      Width           =   2400
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   285
      Index           =   3
      Left            =   1590
      TabIndex        =   23
      Top             =   3885
      Width           =   690
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   285
      Index           =   2
      Left            =   3285
      TabIndex        =   17
      Top             =   2295
      Width           =   690
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   285
      Index           =   1
      Left            =   2805
      TabIndex        =   11
      Top             =   1245
      Width           =   690
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   285
      Index           =   0
      Left            =   3120
      TabIndex        =   4
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form1  (Main.frm)

' Silly Horizontal Scrollbars by Robert Rayment June 2003

' NB This only uses the Horizontal routines in Main.frm & Module1.bas.

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


' NB These same Control Names must be used.  Also even
'    for one scrollbar it must have an index (ie 0).

' Things that could be added: include LargeChange &
' adjust thumb size according to range of values.



Private Sub Form_Load()
   
Dim Spare As Long
Dim j As Long
   
   Me.Caption = " Silly Scrollbars  by Robert Rayment"
   
   '-----------------------------------------------------
   ' FOR FOUR HORIZONTAL SCROLLBARS  HS(0), HS(1(, HS(2) & HS(3)
   
   ' REQUIRED
   
   ReDim zHSMin(3), zHSMax(3)
   ReDim zHSSlope(3), zHSCut(3)
   ReDim HSVal(3)
   ReDim aHSSwapMinMax(3)
   
   ' SCALE also sets STX,STY TwipsPerPixel
   SCALE_HorzScrollbars Me, HS(0), 0
   SCALE_HorzScrollbars Me, HS(1), 1
   SCALE_HorzScrollbars Me, HS(2), 2
   SCALE_HorzScrollbars Me, HS(3), 3
   
   ' OPTIONAL STUFF
   
   ' Show start values as Max < or > Min - Labels - optional
   Label2(0).Caption = HS(0).Min
   If aHSSwapMinMax(0) Then Label2(0).Caption = HS(0).Max    ' ie Swapped
   Label2(1).Caption = HS(1).Min
   If aHSSwapMinMax(1) Then Label2(1).Caption = HS(1).Max    ' ie Swapped
   Label2(2).Caption = HS(2).Min
   If aHSSwapMinMax(1) Then Label2(2).Caption = HS(2).Max    ' ie Swapped
   
   
   ' Shaping 1st horizontal scrollbar - optional
   Spare = CreateEllipticRgn(0, 0, cmdHSTUB(0).Width / 20, cmdHSTUB(0).Height / 15)
   SetWindowRgn cmdHSTUB(0).hWnd, Spare, True
   DeleteObject Spare
   
   Spare = CreateRoundRectRgn(0, 2, picHS(0).Width, picHS(0).Height, 40, 40)
   SetWindowRgn picHS(0).hWnd, Spare, True
   DeleteObject Spare
   
   ' Shaping 2nd horizontal scrollbar - optional
   Spare = CreateRoundRectRgn(0, 2, 13, 34, 8, 10)
   SetWindowRgn cmdHSTUB(1).hWnd, Spare, True
   DeleteObject Spare
   Spare = CreateRoundRectRgn(0, 2, 13, 34, 8, 10)
   SetWindowRgn cmdHSLEF(1).hWnd, Spare, True
   DeleteObject Spare
   Spare = CreateRoundRectRgn(0, 2, 13, 34, 8, 10)
   SetWindowRgn cmdHSRIT(1).hWnd, Spare, True
   DeleteObject Spare
   
   Spare = CreateRoundRectRgn(0, 2, picHS(1).Width, picHS(1).Height, 200, 200)
   SetWindowRgn picHS(1).hWnd, Spare, True
   DeleteObject Spare
   
   
   ' Shaping 3rd horizontal scrollbar - optional
   For j = 1 To picHS(2).Height * STY Step 45
      picHS(2).Line (0, j)-(picHS(2).Width * STX, j), RGB(255, 150, 110)
   Next j
   
   Spare = CreateEllipticRgn(0, 0, cmdHSTUB(2).Width / 15, cmdHSTUB(2).Height / 15)
   SetWindowRgn cmdHSTUB(2).hWnd, Spare, True
   DeleteObject Spare
   
   Spare = CreateEllipticRgn(0, 0, picHS(2).Width, picHS(2).Height)
   SetWindowRgn picHS(2).hWnd, Spare, True
   DeleteObject Spare
   
   ' Wingdings arrows - optional
   cmdHSLEF(1).Caption = Chr$(215)
   cmdHSRIT(1).Caption = Chr$(216)
   

   Me.Show

   cmdHSTUB(0).SetFocus

End Sub

'#### HORIZONTAL SCROLLBARS ### All these needed for horizontal scrollbars ###################################

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
   
   Label2(Index).Caption = HSVal(Index) ' - Labeling optional to show values
   
   ' Special thumb numbering - optional
   If Index = 2 Then cmdHSTUB(2).Caption = Str$(HSVal(Index))

End Sub

'###################################################################################
'Trebor Tnemyar


