VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Begin VB.Form Form1 
   Caption         =   "Text-2-Speech"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS SpeechCtrl 
      Height          =   615
      Left            =   0
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   3240
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "Settings"
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Stop"
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton ctrlPause 
         Caption         =   "Pause"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Ctrlplay 
         Caption         =   "Play"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Txtsay 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   5220
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
SpeechCtrl.AudioReset
Form1.Visible = False
Form2.Visible = True
End Sub


Private Sub Command5_Click()
'stops voice
SpeechCtrl.AudioReset
End Sub








Private Sub ctrlPause_Click()
'Pause voice & Resume voice
Select Case ctrlPause.Caption
Case Is = "Pause"
SpeechCtrl.AudioPause
ctrlPause.Caption = "Resume"
Case Is = "Resume"
SpeechCtrl.AudioResume
ctrlPause.Caption = "Pause"
End Select
End Sub


Private Sub Ctrlplay_Click()

SpeechCtrl.Speak Txtsay 'speeks what is in txtsay

SpeechCtrl.Pitch = Form2.Pitch.Value 'Changes pitch of the voice

SpeechCtrl.Speed = Form2.Speed.Value 'changes speed of the voice

End Sub




Private Sub Form_Load()
LoadData App.Path & "\Settings.ini"
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub


