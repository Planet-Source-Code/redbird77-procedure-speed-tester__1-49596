Attribute VB_Name = "mGroup_LNGtoRGB"
' mGroup_LNGtoRGB

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub RunTests(ByRef G As tGroup)

Dim i As Long, lRepeats As Long, T As cStopwatch

    Set T = New cStopwatch
    lRepeats = G.RepeatCount
    
' Declare variables here that are unique to the procedures to be tested.
Dim bRGB(2) As Byte

    ' -PROCEDURE 1---------------------------------------------------
    If G.ProcedureCount > 0 Then
        T.Reset
        For i = 1 To lRepeats
           vbLNGtoRGB vbYellow, bRGB()
        Next
        G.Procedures(0).Speed = T.Elapsed
    End If
    
    ' -PROCEDURE 2---------------------------------------------------
    If G.ProcedureCount > 1 Then
        T.Reset
        For i = 1 To lRepeats
           LNGtoRGB_wAPI vbYellow, bRGB()
        Next
        G.Procedures(1).Speed = T.Elapsed
    End If

    ' -PROCEDURE 3---------------------------------------------------

    ' Empty.

    ' -PROCEDURE 4---------------------------------------------------
    
    ' Empty.
    
    ' -PROCEDURE 5---------------------------------------------------
    
    ' Empty.
    
    ' ---------------------------------------------------------------
    
    Set T = Nothing

End Sub

' #PROCEDURE1#
Private Sub vbLNGtoRGB(ByVal lColor As Long, ByRef bRGB() As Byte)

' Use bit-shifting and bit-masking.

    bRGB(0) = lColor And &HFF
    bRGB(1) = (lColor \ &H100) And &HFF
    bRGB(2) = (lColor \ &H10000) And &HFF

End Sub

' #PROCEDURE2#
Private Sub LNGtoRGB_wAPI(ByVal lColor As Long, ByRef bRGB() As Byte)

' Use the RtlMoveMemory API.
    
    CopyMemory bRGB(0), lColor, 3

End Sub


