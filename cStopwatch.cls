VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************************
'  Copyright ©1995-2002, Karl E. Peterson, All Rights Reserved.
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************

' Win32 API declarations.
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeGetDevCaps Lib "winmm.dll" (lpTimeCaps As TIMECAPS, ByVal uSize As Long) As Long

' API Structure definitions.
Private Type TIMECAPS
    wPeriodMin As Long
    wPeriodMax As Long
End Type

' Set aside storage for private member variables.
Private m_StartTime As Long
Private m_PeriodMin As Long
Private m_PeriodMax As Long

' Desired timer resolution.
Private Const MINRES As Long = 1

Private Sub Class_Initialize()

Dim tc As TIMECAPS

    ' Retrieve system timer resolution.
    timeGetDevCaps tc, Len(tc)
    
    m_PeriodMin = tc.wPeriodMin
    m_PeriodMax = tc.wPeriodMax

    ' Reprogram timer chip for 1ms resolution.
    timeBeginPeriod MINRES

    ' Store starting time.
    Me.Reset
    
End Sub

Private Sub Class_Terminate()

    ' Balance initial call, letting system know
    ' we no longer need higher-resolution.
    'Debug.Print timeEndPeriod(MINRES)
    
End Sub

Public Property Get Elapsed() As Long
    
    ' Read-Only: return elapsed time in milliseconds since stopwatch was reset.
    Elapsed = timeGetTime() - m_StartTime

End Property

' ** redbird77 modified 2003-10-24 **
Public Property Get ElapsedStr() As String
    
    ElapsedStr = CStr((timeGetTime() - m_StartTime) / 1) & " ms"

End Property

Public Property Get MinimumResolution() As Long

    ' Read-Only: return minimum number of milliseconds timer is capable of resolving.
    MinimumResolution = m_PeriodMin
    
End Property

Public Property Get MaximumResolution() As Long
    
    ' Read-Only: return maximum number of milliseconds timer is capable of resolving.
    MaximumResolution = m_PeriodMax
    
End Property

Public Sub Reset()

    ' Wait until next tick before storing start time.
    m_StartTime = timeGetTime()
    
    Do
    '
    Loop Until m_StartTime <> timeGetTime()
    
    ' Initialize starting time.
    m_StartTime = timeGetTime()
    
End Sub
