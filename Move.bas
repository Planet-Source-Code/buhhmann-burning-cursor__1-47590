Attribute VB_Name = "Module1"
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Const Menge As Long = 320
Public P(Menge) As Koordinate

Public Type Koordinate
Xpos As Single
Ypos As Single
Farbe As Byte
Fmov As Single
Xmov As Single
Ymov As Single
End Type

Public Function Bewegung(X1, Y1, X2, Y2, Geschwindigkeit) As Koordinate
On Error Resume Next
aa = (X1 - X2)
bb = (Y1 - Y2)
cc = Sqr((aa ^ 2) + (bb ^ 2))
aa = aa / (cc / Geschwindigkeit)
bb = bb / (cc / Geschwindigkeit)
Bewegung.Xpos = X1 - aa
Bewegung.Ypos = Y1 - bb
End Function

Public Sub Pause(Dauer)
start = Timer
Do While Timer < start + Dauer
DoEvents
Loop
End Sub

