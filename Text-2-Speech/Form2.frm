VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Settins"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3405
   LinkTopic       =   "Form2"
   ScaleHeight     =   2685
   ScaleWidth      =   3405
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider Pitch 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "hm"
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Min             =   50
      Max             =   400
      SelStart        =   50
      TickStyle       =   3
      Value           =   50
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin MSComctlLib.Slider Speed 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Min             =   50
      Max             =   250
      SelStart        =   50
      TickStyle       =   3
      Value           =   50
   End
   Begin VB.Label Label2 
      Caption         =   "Speed:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Pitch:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'saves settings for user and go back to Form1
SaveData App.Path & "\Settings.ini"
Form2.Visible = False
Form1.Visible = True
End Sub

Private Sub Command2_Click()
'just tests the voice settings
Select Case Command2.Caption
Case Is = "Test"
With Form1
.SpeechCtrl.Speak "This program was Created by overload and He Would Like to thank you for downloading his Text to Voice application"
.SpeechCtrl.Pitch = Form2.Pitch.Value
.SpeechCtrl.Speed = Form2.Speed.Value
End With
Command2.Caption = "Stop"
Case Is = "Stop"
Form1.SpeechCtrl.AudioReset
Command2.Caption = "Test"
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveData App.Path & "\Settings.ini"
End Sub


Private Sub Pitch_Change()
Label1.Caption = "Pitch: " & Pitch.Value
End Sub


Private Sub Speed_Change()
Label2.Caption = "Speed: " & Speed.Value
End Sub


