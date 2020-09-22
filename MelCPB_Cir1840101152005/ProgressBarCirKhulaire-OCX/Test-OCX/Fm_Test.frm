VERSION 5.00
Object = "{749556F6-C635-497B-9D1A-A9552F5CC2C2}#1.0#0"; "MelCPB.ocx"
Begin VB.Form Fm_Test 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MelCPB (OCX test)  : Graphical circular progress bar ..."
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjMelCPB.MelCPB MelCPB3 
      Height          =   555
      Left            =   8160
      TabIndex        =   16
      Top             =   1680
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   979
      cpbOuterCircleMax=   4
      cpbPicInitialState=   "Fm_Test.frx":0000
      cpbPicFinalState=   "Fm_Test.frx":0E8A
      cpbPicBorder    =   -1  'True
   End
   Begin prjMelCPB.MelCPB MelCPB2 
      Height          =   1905
      Left            =   4200
      TabIndex        =   15
      Top             =   1680
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   3360
      cpbInnerCircleMax=   500
      cpbPicInitialState=   "Fm_Test.frx":1D14
      cpbPicFinalState=   "Fm_Test.frx":6168
   End
   Begin prjMelCPB.MelCPB MelCPB1 
      Height          =   2895
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5106
      cpbDrawLine     =   -1  'True
      cpbPicInitialState=   "Fm_Test.frx":121BC
      cpbPicFinalState=   "Fm_Test.frx":177B6
   End
   Begin VB.CommandButton Cmd_Test2 
      Caption         =   "start DrawInnerCircle and DrawOuterCircle test"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   6480
      Width           =   3255
   End
   Begin VB.CommandButton Cmd_Test1 
      Caption         =   "Start DrawBothCircles test"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   6480
      Width           =   3255
   End
   Begin VB.CommandButton Cmd_Test3 
      Caption         =   "Start Real test :-)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   3
      Top             =   6480
      Width           =   3255
   End
   Begin VB.CommandButton Cmd_Quit 
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   7320
      Width           =   9975
   End
   Begin VB.OptionButton Opt_Circle 
      Caption         =   "Only the inner circle bar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   3720
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.OptionButton Opt_Circle 
      Caption         =   "Only the outer circle bar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   0
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Lbl_Title 
      BackStyle       =   0  'Transparent
      Caption         =   "This tutorial will show you how to create such a circular progress bar. The image can be changed."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Top             =   480
      Width           =   7215
   End
   Begin VB.Label Lbl_Title 
      BackStyle       =   0  'Transparent
      Caption         =   "Create a graphical circular progress bar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   12
      Top             =   120
      Width           =   7215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   750
      Left            =   240
      Picture         =   "Fm_Test.frx":1DCB2
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DrawBothCircles test"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Shape Shape3 
      Height          =   4575
      Left            =   120
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Fm_Test.frx":1E9A2
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DrawInnerCircle  and DrawOuterCircle test"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      Height          =   4575
      Left            =   3480
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This test allows you to draw the inner or the outer circle independently."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   3480
      TabIndex        =   8
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Real  test"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6840
      TabIndex        =   7
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Shape Shape4 
      Height          =   4575
      Left            =   6840
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Fm_Test.frx":1EA53
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   6840
      TabIndex        =   6
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   10080
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "Fm_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'___________________________________________________________________________
' Program name      : MelCPB OCX.
' Description       : A simple graphical circular progress bar OCX.
' Company           : MELANTECH
' Authors           : Weitten Pascal
'___________________________________________________________________________
'
' Date              : (c) 2005.01.15
' Version N°        : V0.1
' Customer          : Internal stuff.
'
' Last Modification : 2005.01.15
'___________________________________________________________________________
' TODO :
'       - Fix the CDbl cast problem (see comments below).
'       -
'___________________________________________________________________________
'
'cpbDrawLine        : If set to true a small line is drawn before the angle drawing.
'cpbInnerCircleMax  : Max value for the inner circle.
'cpbOuterCircleMax  : Max value for the outer circle.
'cpbPicInitialState : Picture of the initial state (grey level or whatever).
'cpbPicFinalState   : Picture of the final state (what it should look like once the progress is finished).
'cpbPicBorder       : is there a borderline around the picture.
'
' Don't forget to register the OCX (projetc,components, browse ...)
'

Private Sub Cmd_Test1_Click()
    Dim i As Double
    
    Call EnableButtons(False)
    MelCPB1.Init_CircularBar
    'The next info can be set here or directly in the properties window.
    MelCPB1.cpbOuterCircleMax = 1000
    
    For i = 1 To MelCPB1.cpbOuterCircleMax
        Call MelCPB1.DrawBothCircles(CDbl(i))   'If you don't put the CDbl cast, then there's a weirdo problem.
    Next i
    
    'If you want to clear the bar then just enable the following line.
    'Call MelCPB1.DrawBothCircles(0)
    Call EnableButtons(True)
End Sub

Private Sub Cmd_Quit_Click()
    End
End Sub

Private Sub Cmd_Test2_Click()
    Dim i As Double
    
    Call EnableButtons(False)
    MelCPB2.Init_CircularBar
    'The next info can be set here or directly in the properties window.
    'MelCPB2.cpbInnerCircleMax = 500
    'MelCPB2.cpbOuterCircleMax = 1000
    
    If Opt_Circle(0).Value = True Then
        For i = 1 To MelCPB2.cpbInnerCircleMax
            Call MelCPB2.DrawInnerCircle(CDbl(i))
        Next i
    Else
        For i = 1 To MelCPB2.cpbOuterCircleMax
            Call MelCPB2.DrawOuterCircle(CDbl(i))
        Next i
    End If
    'If you want to clear the bar then just enable the following line.
    'Call MelCPB2.DrawBothCircles(0)
    Call EnableButtons(True)
End Sub

Private Sub Cmd_Test3_Click()
    Dim i As Double, j As Double
    Dim strFileName As String
    Dim FileNumber As Integer

    On Error Resume Next
    Call EnableButtons(False)
    'The files will be in the TestFiles directory.

    MelCPB3.cpbOuterCircleMax = 4       'We'll create 4 files.
    MelCPB3.Init_CircularBar            'Needed :-)

    'Just in case.
    For i = 1 To MelCPB3.cpbOuterCircleMax
        If FileLen(App.Path + "\TestFiles\" + strFileName + Format(i, "00")) > 0 Then Kill App.Path + "\TestFiles\" + strFileName + Format(i, "00")
    Next i
    
    For i = 1 To MelCPB3.cpbOuterCircleMax
        FileNumber = FreeFile
        Open App.Path + "\TestFiles\" + strFileName + Format(i, "00") For Append As #FileNumber

            'For some fun we'll change the number of lines for the files.
            If i = 1 Then MelCPB3.cpbInnerCircleMax = 100
            If i = 2 Then MelCPB3.cpbInnerCircleMax = 2000
            If i = 3 Then MelCPB3.cpbInnerCircleMax = 8400
            If i = 4 Then MelCPB3.cpbInnerCircleMax = 523

            For j = 1 To MelCPB3.cpbInnerCircleMax
                Call MelCPB3.DrawInnerCircle(CDbl(j))
                Print #1, "Line" + Format(j, "0000")
            Next j
        Close #FileNumber
        Call MelCPB3.DrawOuterCircle(CDbl(i))
    Next i

    'If you want to clear the bar then just enable the following line.
    'Call MelCPB3.DrawBothCircles(0)
    Call EnableButtons(True)
End Sub

Sub EnableButtons(butEnabled As Boolean)
    Cmd_Test1.Enabled = butEnabled
    Cmd_Test2.Enabled = butEnabled
    Cmd_Test3.Enabled = butEnabled
    Cmd_Quit.Enabled = butEnabled
End Sub

