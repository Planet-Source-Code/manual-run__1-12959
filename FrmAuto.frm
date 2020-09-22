VERSION 5.00
Begin VB.Form FrmAuto 
   BorderStyle     =   0  'None
   Caption         =   "Manual Run"
   ClientHeight    =   1590
   ClientLeft      =   5835
   ClientTop       =   4800
   ClientWidth     =   3375
   Icon            =   "FrmAuto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Acid Technology® All Rights Reserved"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Compiled: November 20th 2000 At 1:00am"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PDL: Visual Basics 6 EE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By Ashley Partington"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Run"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "FrmAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------'
'This code demonstrates a quite in-proper way of simulating the modern cd drives '
'autorun feature, i made this simply becoz i hate autorun so have disabled it in '
'system settings, but autorun is also a shortcut to loading cd's, so i converted '
'it to "Manual Run", doesnt load automatically, yet still does the same action   '
'This code also shows a trick i use in quite a few of my programs i make, its the'
'Keydown feature, i often use it to hold secret parts in a program etc. If a key '
'i choose is pushed upon load, the prog loads a dif way, pretty sneaky way of    '
'secret areas in appz me thinks, anyway, this code is free to use as you please  '
'i dont care what you do with it, rename it and say u coded it for all i care    '
'(Just dont forget the truth, and rate this code on PSC "ash_acid@hotmail.com")  '
'--------------------------------------------------------------------------------'

Private Declare Function SetWindowPos Lib "user32" (ByVal h%, ByVal hb%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer 'Delares a function from a file
Const HWND_TOPMOST = -1 'Declares a constant
Dim CdInfo 'Holds decision on whether the user wants to view info or use autoplay

Private Sub Form_Load() 'Very start of prog
StayOnTop Me 'Does the StayOnTop function
Me.Width = "0" 'Resizes form width
Me.Height = "0" 'Resizes form height
'Form size set to 0 x 0 so that it is hiden, keydown doesnt work if form invisable
End Sub 'Exits/Finish form load

Private Sub Form_KeyDown(key As Integer, shift As Integer) 'Start of keydown
If key = 17 Then 'If the key thats being held down is the "Ctrl" key then
    Timer1.Enabled = False 'Disables timer
    ShowCdInfo 'Do this function
    CdInfo = True 'Sets variable to "True", telling program not to run autoplay
End If 'End of IF Statment
End Sub 'Exits keydown

Private Sub Form_Click() 'Start of form click
End 'If form is clicked then ends this prog
End Sub 'End of form click

Private Function ShowCdInfo() 'Start of ShowCdInfo function
Me.Width = "3375" 'Resizes form width
Me.Height = "1605" 'Resizes form height
'Does this so the form can be seen (remember it was set to 0 x 0 at form load)
End Function 'Ends the function

Private Function Autorun() 'Start of Autorun function
Dim Value, File 'Declaring variables
On Error GoTo erro 'If theres an error it skips function and goes to end
Open "D:\Autorun.inf" For Input As #1 'Opens the autoplay info on the cd
Do While Not EOF(1) 'Do the below until the end of file has been reached
Input #1, Value 'The files content is now held in the variable "Value"
If Left(Value, 5) = "OPEN=" Then 'Looks for this text in the file
    File = Right(Value, Len(Value) - 5) 'takes all text to the right of that text
ElseIf Left(Value, 5) = "open=" Then 'If previous text not found, looks for this
    File = Right(Value, Len(Value) - 5) 'takes all text to the right of that text
End If 'Ends  IF statment
Loop 'Loops until end of file
Shell "D:\" & File 'Loads autoplay
End 'Ends this prog
erro: 'If theres an error, the function skips the above and comes straight to here
MsgBox "CD Doesn't Support Autorun", vbCritical, "Error" 'Shows a error message
End 'Then ends this prog
End Function 'Ends the function

Private Function Pause(Length) 'Starts the Pause function
Current = Timer 'Sets the current variable to the same as the Timer varaible
Do While Timer - Current < Val(Length) 'Do While Timer - Current = less than pause
    DoEvents 'Increments the Timer variable (increasing its value)
Loop 'Loop the process
End Function 'Ends the function

Private Function StayOnTop(frm As Form) 'Starts the StayOnTop function
On Error GoTo skip 'If theres an error, skips the function
Dim OnTop 'Declares variable
OnTop = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS) 'Sets form ontop
skip: 'Where the functions continues from after an error
End Function 'Ends the function

Private Sub Label1_Click(Index As Integer) 'Start of label click
End 'If label is clicked then ends this prog
End Sub 'End of label click

Private Sub Timer1_Timer() 'Start of timer
Pause 1 'Pause for 1 second to allow for keydown to be received
If CdInfo = True Then GoTo NoAuto 'Finding out whether user wants autoplay
Autorun 'Does autorun function (if keydown is Ctrl, autorun will skip)
NoAuto: 'Where it carries on from is CdInfo = True
Exit Sub 'Exits the timer
End Sub 'Exits the timer







