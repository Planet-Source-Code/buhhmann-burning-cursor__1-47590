VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   272
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   4080
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   3825
      Left            =   120
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   120
      Width           =   3825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim XX As Integer
Dim YY As Integer
Dim xxx As Integer
Dim yyy As Integer

Private Sub Command1_Click()
On Error Resume Next

Do: DoEvents
For i = 0 To Menge: DoEvents
Picture1.PSet (P(i).Xpos, P(i).Ypos), Picture1.BackColor
P(i).Xpos = P(i).Xpos + P(i).Xmov
P(i).Ypos = P(i).Ypos + P(i).Ymov
P(i).Farbe = P(i).Farbe - P(i).Fmov
If P(i).Farbe < 5 Then
P(i).Farbe = 255
P(i).Xpos = XX
P(i).Ypos = YY
P(i).Fmov = 4 * Rnd + 2
End If
SetPixel Picture1.hdc, P(i).Xpos, P(i).Ypos, RGB(P(i).Farbe, P(i).Xpos, P(i).Ypos)
'SetPixel Picture1.hdc, P(i).Xpos, P(i).Ypos, RGB(0, P(i).Farbe, 0)
Next i
Loop

End Sub


Private Sub Form_Load()
On Error Resume Next

Randomize
XX = Picture1.Width / 2
YY = Picture1.Height / 2

For i = 0 To Menge
P(i).Xpos = XX
P(i).Ypos = YY
P(i).Farbe = 5
P(i).Xmov = 1 * Rnd - 0.5
P(i).Ymov = 1 * Rnd - 0.5
P(i).Fmov = 4 * Rnd + 2
Next i
xxx = Picture1.Width * Rnd
yyy = Picture1.Height * Rnd
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
XX = x
YY = y
End Sub

Private Sub Timer1_Timer()
XX = Bewegung(XX, YY, xxx, yyy, 1).Xpos
YY = Bewegung(XX, YY, xxx, yyy, 1).Ypos
If XX = xxx Then If YY = yyy Then xxx = Picture1.Width * Rnd: yyy = Picture1.Height * Rnd
End Sub
