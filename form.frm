VERSION 5.00
Begin VB.Form Text 
   Caption         =   "Text Animation Example"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   5190
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   1800
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3960
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "WWW.TENGAUGE.COM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
   Begin VB.Menu Exit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Text"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Text Animation Example
'John Heslop
'Tengauge Software Inc.
'http://www.Tengauge.com
'Pebbycat@tengauge.co.uk
'
'I don't care if you make millions out of it, e-mail me comments please, at least!
'
'I hope this helps you!
'
'Please feel free to e-mail me @ Pebbycat@tengauge.co.uk for questions. I get lonley sometimes!!
'
'This was created for newbies, it is very simple
'
'Just move the Label where you want and see the results
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub About_Click()

'Pop up box with a information image on, & vbCrLf &  that means NL (New Line)

MsgBox "Tengauge Software inc. - Please vote for me if you have time :)" & vbCrLf & "Visit our site @ http://www.tengauge.com - Next Generation Tech Portal" & vbCrLf & "Planet Source Code Rules!!!!!", vbQuestion, "About Text Animation"

End Sub

Private Sub Command1_Click()

'Exits the Program

Unload Me

End Sub

Private Sub Exit_Click()

'Exits the Program

Unload Me

End Sub

Private Sub Label1_Click()
    
'This is how i make a link to goto a url, not sure if it works on every system

End Sub

Private Sub Timer1_Timer()
With Label1
If .Font.Size < 20 Then .Font.Size = .Font.Size + 1
If .Font.Size > 20 Then Timer1.Enabled = False: Timer2.Enabled = True
End With
End Sub
Private Sub Timer2_Timer()
With Label1

'Change the numbers in the 2 lines below to see diffent effects
         'This number                        'this is speed
If .Font.Size > 2 Then .Font.Size = .Font.Size - 1
If .Font.Size < 2 Then Timer2.Enabled = False: Timer1.Enabled = True
End With
End Sub



