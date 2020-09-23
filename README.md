<div align="center">

## CStopWatch


</div>

### Description

Want to know how long it takes to execute some piece of logic in your code? Use this StopWatch class to find out. It does everything you'd expect a stopwatch to do: - Start - Stop - Reset - Get elapsed time - Get lap time. This class is so simple to use because you already know how a stopwatch works.
 
### More Info
 
Create a new class module and paste the text into it. Name the class CStopWatch.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Janofsky](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-janofsky.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-janofsky-cstopwatch__1-8546/archive/master.zip)





### Source Code

```
Option Explicit
Option Compare Text
'
'-- Copyright Matthew Janofsky 2000
'
'-- Use the class to implement a stopwatch whenever
' you want to time how many milliseconds it takes
' to perform some action.
'
' Example usage:
'
' Public Sub MySub()
' Dim SW As CStopWatch
' Dim X As Long
'
' Set SW = New CStopWatch
'
' '-- Start the timer.
' SW.StartTimer
' For X = 1 To 100000
'  '-- Do something.
'  If X Mod 10000 = 0 Then
'  '-- Show the lap time.
'  Debug.Print " Laptime: " & SW.LapTime _
'    & " Elapsed: " & SW.ElapsedMilliseconds
'  End If
' Next X
' SW.StopTimer
' Debug.Print "Loop Time: " & SW.ElapsedMilliseconds
'
' Set SW = Nothing
' End Sub
'
' Debug output:
' Laptime: 0 Elapsed: 0
' Laptime: 6 Elapsed: 6
' Laptime: 5 Elapsed: 11
' Laptime: 4 Elapsed: 15
' Laptime: 5 Elapsed: 20
' Laptime: 5 Elapsed: 25
' Laptime: 5 Elapsed: 30
' Laptime: 0 Elapsed: 30
' Laptime: 5 Elapsed: 35
' Laptime: 5 Elapsed: 40
' Loop Time: 40
'-- Local Declares
Private Declare Function GetTickCount Lib "kernel32" () As Long
'-- Local private variables
Private m_lStartTime As Long
Private m_lEndTime As Long
Private m_lLastLapTime As Long
Public Sub StopTimer()
 On Error GoTo StopTimer_Error
 m_lEndTime = GetTickCount()
 '-- Exit the procedure.
 GoTo StopTimer_Exit
StopTimer_Error:
 Err.Raise Err.Number, "CStopWatch::StopTimer()", _
 Err.Description, Err.HelpFile, Err.HelpContext
 Resume StopTimer_Exit
 Resume 'For debugging purposes
StopTimer_Exit:
End Sub
Public Sub ResetTimer()
 On Error GoTo ResetTimer_Error
 m_lStartTime = 0
 m_lEndTime = 0
 m_lLastLapTime = 0
 '-- Exit the procedure.
 GoTo ResetTimer_Exit
ResetTimer_Error:
 Err.Raise Err.Number, "CStopWatch::ResetTimer()", _
 Err.Description, Err.HelpFile, Err.HelpContext
 Resume ResetTimer_Exit
 Resume 'For debugging purposes
ResetTimer_Exit:
End Sub
Public Sub StartTimer()
 On Error GoTo StartTimer_Error
 Dim lStoppedTime As Long
 '-- If there is an endtime, we need to calculate how much time
 ' has elapsed since it was stopped and adjust the start time
 ' and last lap time accordingly. We don't want to
 ' include time that passed while the watch was stopped.
 If m_lEndTime > 0 Then
 '-- How long were we stopped?
 lStoppedTime = GetTickCount() - m_lEndTime
 '-- Adjust the start time.
 m_lStartTime = m_lStartTime + lStoppedTime
 '-- Adjust the LapTime.
 m_lLastLapTime = m_lLastLapTime + lStoppedTime
 Else
 '-- First time we've started. Just capture the start time.
 m_lStartTime = GetTickCount()
 End If
 '-- Clear the endtime.
 m_lEndTime = 0
 '-- Exit the procedure.
 GoTo StartTimer_Exit
StartTimer_Error:
 Err.Raise Err.Number, "CStopWatch::StartTimer()", _
 Err.Description, Err.HelpFile, Err.HelpContext
 Resume StartTimer_Exit
 Resume 'For debugging purposes
StartTimer_Exit:
End Sub
Public Property Get ElapsedMilliseconds() As Long
 On Error GoTo ElapsedMilliseconds_Error
 If m_lStartTime = 0 Then
 '-- The timer hasn't started yet. Return 0.
 ElapsedMilliseconds = 0
 GoTo ElapsedMilliseconds_Exit
 End If
 If m_lEndTime = 0 Then
 '-- The user has not clicked stop yet. Give an elapsed time.
 ElapsedMilliseconds = GetTickCount() - m_lStartTime
 Else
 '-- There is a stop time. Just calculate the difference.
 ElapsedMilliseconds = m_lEndTime - m_lStartTime
 End If
 '-- Exit the procedure.
 GoTo ElapsedMilliseconds_Exit
ElapsedMilliseconds_Error:
 Err.Raise Err.Number, "CStopWatch::ElapsedMilliseconds()", _
 Err.Description, Err.HelpFile, Err.HelpContext
 Resume ElapsedMilliseconds_Exit
 Resume 'For debugging purposes
ElapsedMilliseconds_Exit:
End Property
Public Property Get Laptime() As Long
 '-- Return the number of seconds since the last LapTime.
 On Error GoTo Laptime_Error
 Dim lCurrentLapTime As Long
 Dim lRetVal As Long
 lCurrentLapTime = Me.ElapsedMilliseconds
 If m_lLastLapTime = 0 Then
 '-- First Lap. Just return the Elapsed Milliseconds.
 lRetVal = lCurrentLapTime
 Else
 lRetVal = lCurrentLapTime - m_lLastLapTime
 End If
 '-- Save the last lap time.
 m_lLastLapTime = lCurrentLapTime
 '-- Return the lap time.
 Laptime = lRetVal
 '-- Exit the procedure.
 GoTo Laptime_Exit
Laptime_Error:
 Err.Raise Err.Number, "CStopWatch::Laptime()", _
 Err.Description, Err.HelpFile, Err.HelpContext
 Resume Laptime_Exit
 Resume 'For debugging purposes
Laptime_Exit:
End Property
```

