Attribute VB_Name = "Module1"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const FadeDelay = 5
Const FadeMax = 255
 Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" (ByVal hWnd As Long, _
   ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long

 Global Const GWL_EXSTYLE = (-20)
 Global Const WS_EX_LAYERED = &H80000
 Global Const LWA_ALPHA = &H2

 Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" (ByVal hWnd As Long, _
  ByVal nIndex As Long) As Long


 Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hWnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
 Global LastAlpha As Long

Sub FadeIn(f As Form)
Dim c
Dim ne As Integer, en(32767) As Boolean
For Each c In f.Controls
 ne = ne + 1
 en(ne) = c.Enabled
 c.Enabled = False
Next
   TransForm f.hWnd, 0
    f.Show
    Dim i As Long
    For i = 0 To FadeMax Step 3
        TransForm f.hWnd, CByte(i)
        DoEvents
        Call Sleep(FadeDelay)
    Next
    TransForm f.hWnd, FadeMax
    i = 0
For Each c In f.Controls
 i = i + 1
 c.Enabled = en(i)
Next
End Sub


Sub FadeOut(f As Form)
On Local Error Resume Next
Dim c
For Each c In f.Controls
 c.Enabled = False
Next
Dim i As Long
    For i = LastAlpha To 0 Step -3
        TransForm f.hWnd, CByte(i)
        DoEvents
        Call Sleep(FadeDelay)
    Next

End Sub

Public Function TransForm(fhWnd As Long, Alpha As Byte) As Boolean
'Set alpha between 0-255
' 0 = Invisible , 128 = 50% transparent , 255 = Opaque
    SetWindowLong fhWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes fhWnd, 0, Alpha, LWA_ALPHA
    LastAlpha = Alpha
End Function


