VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Pause optimisation"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Long Parameter"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   3960
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Optimised2"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Optimised1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Unoptimised"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Read the comments in the Click event for this button for explanation of what is happening"
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Aussie so it is 's' not 'z'."
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   5
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   1215
      Left            =   3720
      TabIndex        =   4
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "The testrig is very primitive so do nothing else on your machine while a test is running as it will interfer"
      Height          =   735
      Index           =   0
      Left            =   3600
      TabIndex        =   3
      Top             =   1800
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'1 Don't apply Code optimisers like Code Fixer to this code or it destroys the point
'2 the timer testing in this prog is pretty primative don't do anything else on your machine while testing or it will interfer
'3 don't click more than one button at a time
'4 the purpose of using a Loop rather than a single call to the routine is to allow timing differences to accumulate
'  without the loop you would probably need atomic clock level detectors to see any real difference.
'  The point is to develop the habit of optimizing
'
Const FLoop As Long = 1000             'change these if you want longer/shorter tests
Const PLen As Double = 0.01            'the defaults FLoop=1000 PLen=0.01 test each loop for 10 seconds
'                                      'If you set PLen to a fraction above 1 (EG 1.5) then the Long demo will not work as expected

Private Sub Command1_Click()

  Dim I As Long

  Dim T As Double
  Command1.Caption = "working"
  T = Timer
  For I = 1 To FLoop
    PauseWeak PLen
  Next I
  Command1.Caption = "Unoptimised  " & Timer - T

End Sub

Private Sub Command2_Click()

  Dim I As Long

  Dim T As Double
  Command2.Caption = "working"
  T = Timer
  For I = 1 To FLoop
    PauseOptimised1 PLen
  Next I
  Command2.Caption = "Optimised1  " & Timer - T

End Sub

Private Sub Command3_Click()

  Dim I As Long

  Dim T As Double
  Command3.Caption = "working"
  T = Timer
  For I = 1 To FLoop
    PauseOptimised2 PLen
  Next I
  Command3.Caption = "Optimised2  " & Timer - T

End Sub

Private Sub Command4_Click()
'This is an interesting optimization
'It actually much runs faster than the doubles
'mainly because it makes less calls to the routine
'but in fact it is a bit too fast; it is consistently under time
'This is because the comparision in the Do line
'coerces the Timer value to Long so as soon as the timer return value
'passes X.5 it jumps to the next whole value, the test succeeds and the loop exits
'
' the stuff at the top is to convert fractional pause(from the Constant) to Longs by
'reducing the Loops and increasing the pause by a factor of 10
'It will not work if you use fractional pauses above 1

Dim T As Double
  Dim I As Long
  Dim lngFloop As Long  'local For maximum value (no tmp requred because FLoop is also Long)
  Dim dblPLen As Double 'temp to move fractional to whole number
  Dim lngPLen As Long   'variable for passing to routine
  dblPLen = PLen
  lngFloop = FLoop
  Do While dblPLen < 1
    lngFloop = lngFloop / 10
    dblPLen = dblPLen * 10
  Loop
lngPLen = dblPLen
  
  Command4.Caption = "working"
  T = Timer
  For I = 1 To lngFloop
    PauseLong lngPLen
  Next I
  Command4.Caption = "Long Parameter  " & Timer - T

End Sub

Private Sub Form_Load()
Label2.Caption = "Requested Pause = " & FLoop * PLen & " secs" & vbNewLine & _
                 "Calls to Pause = " & FLoop & vbNewLine & _
                 "Pause Len = " & PLen
End Sub

Private Sub Form_Unload(Cancel As Integer)
 End
End Sub
