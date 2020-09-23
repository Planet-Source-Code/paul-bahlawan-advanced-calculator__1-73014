VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF00FF&
   Caption         =   "Advanced Calculator"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer15 
      Enabled         =   0   'False
      Interval        =   77
      Left            =   4080
      Top             =   6360
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Interval        =   77
      Left            =   4320
      Top             =   6360
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Interval        =   77
      Left            =   3960
      Top             =   6360
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   77
      Left            =   4200
      Top             =   6360
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   77
      Left            =   3840
      Top             =   6360
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   77
      Left            =   4440
      Top             =   6360
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   77
      Left            =   4080
      Top             =   6360
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   77
      Left            =   4200
      Top             =   6360
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   77
      Left            =   3720
      Top             =   6360
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   77
      Left            =   3120
      Top             =   6360
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   2760
      Top             =   6480
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   2280
      Top             =   6360
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   1800
      Top             =   6360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   1200
      Top             =   6480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   600
      Top             =   6360
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H0000C000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H000000FF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C000&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C000&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C000&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C000&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C000&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1575
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Caption = Label1.Caption & "1"
End Sub

Private Sub Command10_Click()
Label1 = Label1 & "10"
End Sub

Private Sub Command2_Click()


    Timer2.Enabled = True


End Sub

Private Sub Command3_Click()
Timer1.Enabled = True
End Sub

Private Sub Command4_Click()



Timer4.Enabled = True
End Sub

Private Sub Command5_Click()
Label1.Caption = Label1.Caption & "%"

End Sub

Private Sub Command6_Click()
Label1.Caption = Label1.Caption & "6"

End Sub

Private Sub Command7_Click()
Label2 = Label1
Label1 = ""
End Sub

Private Sub Command8_Click()
Timer15.Enabled = True
End Sub

Private Sub Command11_Click()
Label1 = Sin(Val(Label1)) + Cos(Val(Label2))
End Sub

Private Sub Command9_Click()
Label1.Caption = Label1.Caption + "9"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
MsgBox "You have pressed the invalid key"
MsgBox "Please do not pressd the invalid key"
End
End Sub

Private Sub Timer1_Timer()
Timer3.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer11_Timer()
Timer11.Enabled = False
Timer15.Enabled = False
End Sub

Private Sub Timer15_Timer()
Label1.Caption = Label1.Caption + "8"
Timer11.Enabled = True
End Sub

Private Sub Timer2_Timer()
Label1.Caption = Label1.Caption & "2"
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
Label1.Caption = 3
End Sub

Private Sub Timer4_Timer()
For a = 4 To 44
DoEvents
Next
Call four
End Sub

Private Function four()
Timer4.Enabled = False
Label1.Caption = "4" & Label1.Caption

End Function
