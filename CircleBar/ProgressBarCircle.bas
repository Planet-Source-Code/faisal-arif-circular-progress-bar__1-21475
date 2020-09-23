Attribute VB_Name = "ProgressBarCircle"
Option Explicit

Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Dim X, Y As Long
Dim i As Long
Public C, S As Single
Public SL, ST, Size As Integer
Public Num As Integer
Public MeterValue As Single
Dim LastNum As Integer
Dim Col1, Col2, Col3 As Long

Public Function LoadMeter(CurForm As Form, BarColor As Long, PercentColor As Long)
On Error GoTo Error
'Basic Numbers
SL = CurForm.MeterShape.Left + CurForm.MeterShape.Width / 2
ST = CurForm.MeterShape.Top + CurForm.MeterShape.Height / 2
Size = CurForm.MeterShape.Width / 2
CurForm.MeterBox.Circle (SL, ST), CurForm.MeterShape.Width / 2 + 1
CurForm.MeterBox.Picture = CurForm.MeterBox.Image
'MeterPos
CurForm.MeterPos.Caption = "0%"
CurForm.MeterPos.Left = 0
CurForm.MeterPos.Width = CurForm.MeterBox.ScaleWidth
CurForm.MeterPos.Top = CurForm.MeterShape.Top + CurForm.MeterShape.Height / 2 - CurForm.MeterPos.Height / 2
CurForm.MeterPos.ForeColor = PercentColor
'Meter
CurForm.MeterBox.ForeColor = BarColor
Col1 = RGB(0, 0, 0)
Col2 = BarColor
Col3 = CurForm.MeterBox.BackColor
SetMeter 0, Form1
Exit Function
Error:
MsgBox Err.Description
End Function

Public Function SetMeter(SetNum As Single, CurForm As Form)
On Error GoTo Error
Dim MeterNum As Single
'Checks if SetNum for Certain Values
SetNum = Round(SetNum)
Select Case SetNum
Case Is <= -3
    SetNum = 0
Case -1
    'Go Up One
    SetNum = LastNum + 1
    If SetNum >= 100 Then
        SetNum = 100
        End If
Case -2
    'Go Down One
    SetNum = LastNum - 1
    If SetNum < 0 Then
        SetNum = 0
        End If
Case Is > 100
    Exit Function
End Select

'Sets the Circle Bar
CurForm.MeterBox.Cls
For i = 0 To Round(SetNum * 3.6, 0)
    C = i
    S = i
    C = Cos(C * (3.14159 / 180))
    S = Sin(S * (3.14159 / 180))
    CurForm.MeterBox.Line (SL, ST)-(S * Size + SL, -C * Size + ST)
Next i
CurForm.MeterBox.Refresh
CurForm.MeterPos.Caption = SetNum & "%"
LastNum = SetNum
Exit Function
Error:
MsgBox Err.Description
End Function

Public Function GetMeter(CurForm As Form)
MeterValue = LastNum
End Function
