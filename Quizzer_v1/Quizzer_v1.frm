VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Quizzer  v1.0  -  NeoProgrammics"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Quizzer_v1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Next_Question_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "?"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6345
      TabIndex        =   31
      ToolTipText     =   " This Button Selects The Next Quiz Question "
      Top             =   2205
      Width           =   1905
   End
   Begin VB.TextBox Status_Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   855
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5040
      Width           =   7395
   End
   Begin VB.Frame Select_Quiz_Frame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Select Quiz From List Below"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3525
      Left            =   8280
      TabIndex        =   20
      ToolTipText     =   " This Is a Listing Of All Available Quiz Files "
      Top             =   0
      Width           =   3570
      Begin VB.CommandButton Load_Quiz_Button 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Load Quiz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2385
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   " This Button Activates The Quiz Selection List "
         Top             =   3105
         Width           =   1095
      End
      Begin VB.FileListBox Quiz_List 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2820
         Left            =   90
         Pattern         =   "*.quiz"
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   225
         Width           =   3390
      End
   End
   Begin VB.TextBox Status_Text 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   855
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2205
      Width           =   5460
   End
   Begin VB.TextBox MultChoice 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   495
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4635
      Width           =   7755
   End
   Begin VB.TextBox MultChoice 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   495
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4230
      Width           =   7755
   End
   Begin VB.TextBox MultChoice 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   495
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3825
      Width           =   7755
   End
   Begin VB.TextBox MultChoice 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   495
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3420
      Width           =   7755
   End
   Begin VB.TextBox MultChoice 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   495
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3015
      Width           =   7755
   End
   Begin VB.TextBox MultChoice 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   495
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2610
      Width           =   7755
   End
   Begin VB.TextBox Current_QNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1125
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   " This Is the Current Question Number "
      Top             =   405
      Width           =   690
   End
   Begin VB.CommandButton Choice 
      BackColor       =   &H00C0C0C0&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   45
      TabIndex        =   6
      Top             =   4635
      Width           =   420
   End
   Begin VB.CommandButton Choice 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   45
      TabIndex        =   5
      Top             =   4230
      Width           =   420
   End
   Begin VB.CommandButton Choice 
      BackColor       =   &H00C0C0C0&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   45
      TabIndex        =   4
      Top             =   3825
      Width           =   420
   End
   Begin VB.CommandButton Choice 
      BackColor       =   &H00C0C0C0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   45
      TabIndex        =   3
      Top             =   3420
      Width           =   420
   End
   Begin VB.CommandButton Choice 
      BackColor       =   &H00C0C0C0&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   45
      TabIndex        =   2
      Top             =   3015
      Width           =   420
   End
   Begin VB.CommandButton Choice 
      BackColor       =   &H00C0C0C0&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   2610
      Width           =   420
   End
   Begin VB.TextBox Work 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1410
      Left            =   45
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   765
      Width           =   8205
   End
   Begin VB.Frame Score_Frame 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1950
      Left            =   8280
      TabIndex        =   24
      ToolTipText     =   " These Are The Current Quiz Statistics "
      Top             =   3465
      Width           =   3570
      Begin VB.TextBox Final_Score 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   855
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Correct_Answers 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   135
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   450
         Width           =   1545
      End
      Begin VB.TextBox Incorrect_Answers 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   1890
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   450
         Width           =   1545
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Final Score"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   855
         TabIndex        =   30
         Top             =   1665
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Correct  Answers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   135
         TabIndex        =   28
         Top             =   225
         Width           =   1545
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incorrect  Answers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1845
         TabIndex        =   27
         Top             =   225
         Width           =   1635
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      TabIndex        =   18
      ToolTipText     =   " This Text Field Is For Status And Error Messages "
      Top             =   5085
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1890
      TabIndex        =   17
      Top             =   450
      Width           =   1905
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Answer:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      TabIndex        =   16
      ToolTipText     =   " Below Is a List of the Multiple Choice Answers Available "
      Top             =   2295
      Width           =   870
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Question  #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   90
      TabIndex        =   15
      Top             =   450
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Title_Label 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   45
      TabIndex        =   8
      Top             =   90
      Width           =   8205
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit

' Quizzer v1.0 - Visual BASIC 6
'
' Written by Jay Tanner - NeoProgrammics.com
'
' This is a simple quiz generating program that can be easily expanded upon.
'
' Given a list of quiz questions with multiple choice answers, the program
' will ask all the questions in the list in a random sequence each time.  The
' multiple-choice answers are also shuffled into a random arrangement each
' time also, so the same quiz taken a second time will occur in a different
' random arrangement of questions and answers each time.
'
' Each quiz question may have from 2 to 6 possible multiple-choice answers.
'
' Many quizzes can be easily written simply using a plain text editor and an
' easy to learn and remember quiz file format.  A quiz can contain as many
' questions as desired.
'
' The docs included with this source code describe how to make the quiz data
' files as simple text files.  The quizzes the programmer can invent are only
' limited by the imagination.
'
'
' As designed, the program looks in the same folder as the application for the
' quiz data files, but this can be easily changed if required.
'
' All quiz files found in the application folder are displayed in a list when
' the program starts and can be selected by double clicking on the file name.
'
' This is a relatively easy to use and understand program and it can easily
' have new features added by any enterprising programmer.
'
  Dim NumQ             As Long    ' Total number of questions in quiz
  Dim Question()       As String  ' Array to hold quiz questions
  Dim Answer()         As String  ' Array to hold answers to current question

  Dim AnsFlag          As Boolean ' Flag to show if question was answered
  Dim QuizFlag         As Boolean ' Flag to show if quiz is loaded and ready
  Dim QPendingFlag     As Boolean ' Flag to show if question pending answer
  Dim Score            As Double  ' Score as percent of correct answers
  Dim AnsCorrect       As Long    ' Number of correct answers
  Dim AnsIncorrect     As Long    ' Number of incorrect answers
  Dim QuizOverFlag     As Boolean ' Flag to indicate if quiz is over

' ==============================================================================
  Private Sub Form_Load()
' What to do when this program starts up

' Set width to 12000 for 800 x 600 screen resolution
  Form1.Width = 12000

' Set default quiz directory to same folder as the program
  ChDrive App.Path
  ChDir App.Path
  RESET_QUIZ
  PRINT_STARTUP_INFO
  End Sub

' ==============================================================================
' This code controls the button to fetch the next quiz question

  Private Sub Next_Question_Button_Click()

' Change caption on button when starting quiz
  If InStr(Next_Question_Button.Caption, "Begin") > 0 Then
     Next_Question_Button.Caption = "Next Question"
     Beep
  End If

' Check if quiz is finished - All questions have been answered.
  If QuizOverFlag = True Then QuizFlag = False: QPendingFlag = False

' Check if current question is still pending an answer
  If QPendingFlag = True Then
  Status_Text2.BackColor = RGB(255, 0, 0)
  Status_Text2.ForeColor = RGB(255, 255, 255)
  Status_Text2 = "The current question has not been answered yet."
  Beep
  Exit Sub
  End If

' Check if a quiz has been loaded and is ready to proceed
  If QuizFlag = False Then
  Status_Text2.BackColor = RGB(255, 0, 0)
  Status_Text2.ForeColor = RGB(255, 255, 255)
  Status_Text2 = "Select a quiz from the yellow list."
  Beep
  Exit Sub
  End If

' Display next available quiz question
  CLEAR
  Current_QNum = Val(Current_QNum) + 1
  DISPLAY_QUESTION (Current_QNum)
  QPendingFlag = True

  End Sub

' ==============================================================================
' This controls the loading of a quiz from the displayed yellow quiz listing.

  Private Sub Load_Quiz_Button_Click()
  
  Dim Tx As String

  Close 1

' Display warning with cancel option if a quiz is already in progress
  If QuizFlag = True Then
     Tx = "Are you sure ?  This will end any quiz currently in progress "
     Tx = Tx & "and no score will recorded."
     Tx = MsgBox(Tx, vbOKCancel, " Start a New Quiz")
     If Tx <> 1 Then Exit Sub ' Exit and resume current quiz if canceled
  End If

' Proceed to activate quiz listing for selection process.
  RESET_QUIZ
  Quiz_List.Enabled = True

  Current_QNum = ""
  CLEAR
  PRT "Load a quiz by double-clicking on its file name in the yellow listing."
  Work.BackColor = RGB(0, 255, 255)
  Work.ForeColor = RGB(0, 0, 0)

  End Sub


' ==============================================================================
' Load in selected quiz file & shuffle the questions into a random sequence.
' This code will load the quiz whose file name is double-clicked on.

  Public Sub Quiz_List_DblClick()

  LOAD_QUIZ
  Current_QNum = ""
  Quiz_List.Enabled = False
  Quiz_List.ListIndex = -1
  QuizFlag = True
  Status_Text.BackColor = RGB(192, 192, 192)
  Status_Text = ""
  Label2.Caption = "of  " & NumQ & "  questions."
  Next_Question_Button.Caption = "Begin Selected Quiz"
  CLEAR
  PRT Title_Label.Caption
  PRT " has been loaded and initialized."
  BLIN
  PRT " Click on the [Begin Selected Quiz] button to start quiz."
  Work.BackColor = RGB(0, 255, 255)
  Work.ForeColor = RGB(0, 0, 0)

  End Sub

' ==============================================================================

' Enable multiple choice buttons from 1 to N where N is the number
' of multiple choice answers from 2 to 6 available for the question.
' If n=0, then all choices are disabled.

  Private Sub ENABLE_CHOICES_1_To(N)

  Dim i   As Integer
   
' First disable all choices
  For i = Choice.LBound To 5
  Choice.Item(i).Enabled = False
  Next i

  If Val(N) = 0 Then Exit Sub

' Activate choices up to N for all other cases
  For i = Choice.LBound To Val(N) - 1
  Choice.Item(i).Enabled = True
  Next i
  
  End Sub
  
' ==============================================================================
' Read quiz file and count rhe number of questions.  Then load the quiz into
' a working array.

  Public Sub LOAD_QUIZ()

  Dim QuizFileName    As String
  Dim QuizTitle       As String
  Dim L               As String ' Line read from the quiz file
  Dim i               As Long   ' Counter index and loop control
  Dim j               As Long   ' Internal string pointer
  Dim R1              As Long   ' Random variable #1
  Dim R2              As Long   ' Random variable #2
  Dim W               As String ' Work string

' Fetch file name of selected quiz
  QuizFileName = Quiz_List.FileName

' Prepare for randomization
  Randomize Timer

' Count questions in quiz
  i = 0
  Open QuizFileName For Input As #1
  While Not EOF(1)
  Line Input #1, L
  If InStr(UCase(L), "<Q>") > 0 Then i = i + 1
  Wend
  Close 1
  NumQ = i

' Create an array to hold the quiz questions
  i = 1
  ReDim Question(1 To NumQ) As String
  Open QuizFileName For Input As #1

' Read in the selected quiz file and store it in array
  While Not EOF(1)
     Line Input #1, L
  If InStr(UCase(L), "<TITLE>") > 0 Then _
     Title_Label = " " & Mid(L, 8, Len(L))
  If InStr(UCase(L), "<Q>") > 0 Then
     Question(i) = L
     Line Input #1, L
     Question(i) = Question(i) & Mid(L, InStr(UCase(L), "<A>"), Len(L))
     i = i + 1
  End If
     
  Wend

  i = i - 1

' Done reading the file
  Close 1

' Shuffle the quiz questions into a random sequence
  For i = 1 To 1000
  R1 = 1 + Int(Rnd * NumQ): R2 = 1 + Int(Rnd * NumQ)
  W = Question(R1): Question(R1) = Question(R2): Question(R2) = W
  Next i

  End Sub

' ==============================================================================
' Display selected question number QN and activate multiple choices
    
  Public Sub DISPLAY_QUESTION(QN)

  Dim Q     As String  ' Quiz question text
  Dim A     As String  ' Answers for question (Q)
  Dim NumA  As Long    ' Number of multiple-choice answers for question (Q)
  Dim i     As Long    ' Loop control index and internal string pointer
  Dim R1    As Long    ' Random variable #1
  Dim R2    As Long    ' Random variable #2
  Dim W     As String  ' Work vriable

  Dim NQ As Integer    ' Index number of currently active question
      NQ = Val(QN)

  AnsFlag = False

' Reset and clear status display text fields
  Status_Text.BackColor = RGB(192, 192, 192)
  Status_Text.ForeColor = RGB(0, 0, 0)
  Status_Text = ""
  Status_Text2.BackColor = RGB(192, 192, 192)
  Status_Text2 = ""

' Disable all multiple-choice selection buttons
  ENABLE_CHOICES_1_To (0)

' Clear any previous answer array
  Answer() = Split(" | | | | | ", "|")

' Clear all answer slots and reset color to white
  CLEAR_ANSWERS

' Check if quiz is finished - All questions answered & tally final score
  If NQ > NumQ Then
  Work.BackColor = RGB(0, 0, 192)
  Work.ForeColor = RGB(255, 255, 255)
  PRT " Quiz Is Finished"
  Beep
  BLIN
  Current_QNum = Current_QNum - 1
  PRT " Total Correct Answers :  " & AnsCorrect
  PRT " Total Incorrect Answers :  " & AnsIncorrect
  W = Format(100 * (AnsCorrect / (NumQ)), "#0.#0") & " %"
  PRT " Final Score Is :  " & W
  Final_Score = W
  QuizOverFlag = True
  Exit Sub
  End If

' Get text of current quiz question and attached answers
  Q = Question(NQ)
  i = InStr(UCase(Q), "<A>")

' Separate multiple-choice answers from question part
  A = Mid(Q, i + 3, Len(Q))

' Separate question text from multiple-choice answers part
  Q = Left(Q, i - 1): Q = Mid(Q, 4, Len(Q))

' Put answers to question into answers array
  Answer() = Split(A, "|")
  NumA = UBound(Answer) + 1

' Shuffle the answer array into a random sequence
  For i = 1 To 1000
  R1 = Int(Rnd * NumA): R2 = Int(Rnd * NumA)
  W = Answer(R1): Answer(R1) = Answer(R2): Answer(R2) = W
  Next i

' Copy answers into answer slots
  For i = 0 To NumA - 1
  W = Answer(i): If Left(W, 1) = "*" Then W = Mid(W, 2, Len(W))
  MultChoice.Item(i) = " " & W
  Next i

' Display question text
  PRT Q

' Activate appropriate multiple choices
  ENABLE_CHOICES_1_To (NumA)

  End Sub

' ==============================================================================
' Check if (Index) refers to the correct answer for the current question.

  Private Sub Choice_Click(Index As Integer)

  Dim W As String

  If AnsFlag = True Then
  Status_Text2.BackColor = RGB(255, 0, 0)
  Status_Text2.ForeColor = RGB(255, 255, 255)
  Status_Text2 = "You have already answered the question !"
  Beep
  Exit Sub
  End If

  MultChoice.Item(Index).BackColor = RGB(0, 255, 255)
  If InStr(Answer(Index), "*") > 0 Then GoTo CORRECT Else GoTo INCORRECT
  Index = Index

  Exit Sub

' Correct answer given
CORRECT:
  Status_Text.BackColor = RGB(128, 255, 0)
  Status_Text.ForeColor = RGB(0, 0, 0)

' Randomize responses
  W = "Correct."
  If Rnd > 0.9 Then W = "Correct answer.": GoTo OK1
  If Rnd > 0.85 Then W = "Good. That is the correct answer.": GoTo OK1
  If Rnd > 0.8 Then W = "That is correct.": GoTo OK1
  If Rnd > 0.75 Then W = "Yes.  You are right.": GoTo OK1
  If Rnd > 0.7 Then W = "You got it.": GoTo OK1
  If Rnd > 0.65 Then W = "Correct answer.": GoTo OK1
  If Rnd > 0.6 Then W = "That's it.": GoTo OK1
  If Rnd > 0.55 Then W = "Good. That's the right answer.": GoTo OK1
  If Rnd > 0.5 Then W = "You got it right.": GoTo OK1
  If Rnd > 0.45 Then W = "Excellent.  You are correct.": GoTo OK1
  If Rnd > 0.4 Then W = "Right.": GoTo OK1
  If Rnd > 0.35 Then W = "OK.  That is correct.": GoTo OK1
  If Rnd > 0.3 Then W = "Yes. That's right.": GoTo OK1
  If Rnd > 0.25 Then W = "Excellent!  You're right.": GoTo OK1
  If Rnd > 0.2 Then W = "You are correct.": GoTo OK1
  If Rnd > 0.15 Then W = "Absolutely correct!": GoTo OK1
  If Rnd > 0.1 Then W = "You are right.": GoTo OK1
OK1:
  Status_Text = W
  Beep
  AnsFlag = True
  AnsCorrect = AnsCorrect + 1
  Correct_Answers = AnsCorrect
  QPendingFlag = False
  Next_Question_Button.SetFocus
  Exit Sub

' Incorrect answer given
INCORRECT:
  Status_Text.BackColor = RGB(255, 0, 0)
  Status_Text.ForeColor = RGB(255, 255, 255)

' Randomize responses
  W = "Incorrect."
  If Rnd > 0.9 Then W = "Sorry, wrong answer.": GoTo OK2
  If Rnd > 0.85 Then W = "That's the wrong answer.": GoTo OK2
  If Rnd > 0.8 Then W = "Wrong.": GoTo OK2
  If Rnd > 0.75 Then W = "Sorry. That's not it.": GoTo OK2
  If Rnd > 0.7 Then W = "No. That's not it.": GoTo OK2
  If Rnd > 0.65 Then W = "Sorry.  That answer is incorrect.": GoTo OK2
  If Rnd > 0.6 Then W = "Incorrect answer.": GoTo OK2
  If Rnd > 0.55 Then W = "Your answer is incorrect.": GoTo OK2
  If Rnd > 0.5 Then W = "Incorrect.": GoTo OK2
  If Rnd > 0.45 Then W = "No. That's wrong.": GoTo OK2
  If Rnd > 0.4 Then W = "You missed it.": GoTo OK2
  If Rnd > 0.35 Then W = "Sorry.  You guessed wrong.": GoTo OK2
  If Rnd > 0.3 Then W = "That is wrong.": GoTo OK2
  If Rnd > 0.25 Then W = "Your answer is wrong.": GoTo OK2
  If Rnd > 0.2 Then W = "Wrong answer.": GoTo OK2
  If Rnd > 0.15 Then W = "That's not the correct answer.": GoTo OK2
  If Rnd > 0.1 Then W = "You got it wrong.": GoTo OK2
OK2:
  Status_Text = W
  Beep
  AnsFlag = True
  AnsIncorrect = AnsIncorrect + 1
  Incorrect_Answers = AnsIncorrect
  QPendingFlag = False
  Next_Question_Button.SetFocus
  End Sub

' ==============================================================================
' Clear all answer slots and reset colors to white
  Private Sub CLEAR_ANSWERS()

  Dim i As Integer

  For i = 0 To 5
  MultChoice.Item(i) = ""
  MultChoice.Item(i).BackColor = RGB(255, 255, 255)
  Next i

  End Sub

' ==============================================================================
' Print quiz startup information

  Private Sub PRINT_STARTUP_INFO()
  
  Dim Tx As String

  Tx = "To take a quiz, click on the [Load Quiz] button on the right and then"
  Tx = Tx & " double-click on the desired quiz file name in the yellow listing"
  Tx = Tx & " to load it in."
  CLEAR
  PRT Tx

  End Sub

' ==============================================================================
' Reset quiz for start or restart
  
  Private Sub RESET_QUIZ()

  QuizFlag = False
  QPendingFlag = False
  CLEAR_ANSWERS
  Status_Text.BackColor = RGB(192, 192, 192)
  Status_Text = ""
  Status_Text2 = ""
  Status_Text2.BackColor = RGB(192, 192, 192)
  ENABLE_CHOICES_1_To (0)
  Next_Question_Button.Caption = "?"
  Quiz_List.Enabled = False
  AnsCorrect = 0
  Correct_Answers = ""
  AnsIncorrect = 0
  Incorrect_Answers = ""
  Final_Score = ""
  QuizOverFlag = False
  End Sub
