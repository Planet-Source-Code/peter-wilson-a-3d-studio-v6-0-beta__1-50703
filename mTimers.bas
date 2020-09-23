Attribute VB_Name = "mTimers"
Option Explicit

' The Sleep function suspends the execution of the current thread for a specified interval.
' (This is like a Pause function for the EXE, ie. It will slow the whole EXE down.)
' (1000 Milliseconds = 1 Second)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

