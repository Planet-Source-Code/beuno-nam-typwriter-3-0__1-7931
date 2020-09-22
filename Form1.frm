VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Type Writer v3  -  Beuno´s Re-Make"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Pause 
      BackColor       =   &H0000C000&
      Caption         =   "Pause"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Interval 
      BackColor       =   &H8000000A&
      Height          =   285
      Left            =   5760
      TabIndex        =   7
      Text            =   "100"
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Clear 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "Clear && Re-start"
      Height          =   375
      Left            =   3120
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox newd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Height          =   285
      Left            =   3075
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   2625
      Width           =   555
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   3600
   End
   Begin VB.CommandButton Start 
      BackColor       =   &H0000C000&
      Caption         =   "Scroll It"
      Height          =   345
      Left            =   120
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1005
   End
   Begin VB.TextBox ST 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1125
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0442
      Top             =   2280
      Width           =   6405
   End
   Begin VB.TextBox SH 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1710
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   270
      Width           =   6420
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the text to Type Write below..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   623
      TabIndex        =   5
      Top             =   2055
      Width           =   3435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The text will start typing on the box below"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   360
      TabIndex        =   4
      Top             =   45
      Width           =   4005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Added the pause/resume
'Added the possibility to change the typing speed
'Added the possibility to type as many things as you want

Private Sub Pause_Click()
Timer1.Enabled = False 'stop Timer
Start.Caption = "Resume"
End Sub

Private Sub Start_Click()
Timer1.Interval = Interval.Text 'Set the timers speed
Timer1.Enabled = True 'Turns on the timer
If Start.Caption = "Resume" Then
Start.Caption = "Scroll It"
End If
End Sub

Private Sub Clear_Click()
Timer1.Enabled = False 'stop Timer
SH.Text = "" 'Clear all text
newd.Text = "0" 'Set hidden box to 0
End Sub

Private Sub Timer1_Timer()
newd.Text = newd.Text + 1 'Sets the hidden text box value
SH.Text = Left(ST.Text, newd.Text) & "–" 'Scroll Here box equals the character next in the string
Timer1.Enabled = False 'Turn off the timer
Timer1.Enabled = True 'Restart the timer to enable code again
If Len(SH.Text) = (Len(ST.Text) + 1) Then 'If Scrolled equals the Scrolling text, then stop.
Timer1.Enabled = False 'stop Timer
Else
End If
End Sub
