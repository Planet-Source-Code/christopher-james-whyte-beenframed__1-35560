Attribute VB_Name = "Module1"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Function OpenURL(ByVal URL As String) As Long
OpenURL = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function
Public Function timedPause(secs As Long)
    Dim secStart As Variant
    Dim secNow As Variant
    Dim secDiff As Variant
    Dim Temp%
    exitPause = False 'this is our early way out out of the pause
    secStart = Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") 'get the starting seconds
    Do While secDiff < secs
        If exitPause = True Then Exit Do
        secNow = Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") 'this is the current time and Date at any itteration of the Loop
        secDiff = DateDiff("s", secStart, secNow) 'this compares the start time With the current time
        Temp% = DoEvents
    Loop
End Function
Public Sub Main()
On Error Resume Next
ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
frmSplash.Show
timedPause 3
Unload frmSplash
If ShowAtStartup = 1 Then
Frame.Show
frmTip.Show 1, Frame
FrameText.Hide
ElseIf ShowAtStartup = 0 Then
Frame.Show
FrameText.Hide
End If
End Sub
