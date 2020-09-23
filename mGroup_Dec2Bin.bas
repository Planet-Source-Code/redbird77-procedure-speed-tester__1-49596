Attribute VB_Name = "mGroup_Dec2Bin"
' mGroup_Dec2Bin.bas

Option Explicit

Public Sub RunTests(ByRef G As tGroup)

Dim i As Long, lRepeats As Long, T As cStopwatch

    Set T = New cStopwatch
    lRepeats = G.RepeatCount
    
' Declare variables here that are unique to the procedures to be tested.
    Dim lDec As Long, sBin As String
    
    lDec = 235454
    
    ' -PROCEDURE 1---------------------------------------------------
    If G.ProcedureCount > 0 Then
        T.Reset
        For i = 1 To lRepeats
           sBin = Dec2Bin1(lDec)
        Next
        G.Procedures(0).Speed = T.Elapsed
    End If
    
    ' -PROCEDURE 2---------------------------------------------------
    If G.ProcedureCount > 1 Then
        T.Reset
        For i = 1 To lRepeats
           sBin = Dec2Bin2(lDec)
        Next
        G.Procedures(1).Speed = T.Elapsed
    End If

    ' -PROCEDURE 3---------------------------------------------------
    If G.ProcedureCount > 2 Then
        T.Reset
        For i = 1 To lRepeats
           sBin = Dec2Bin3(lDec)
        Next
        G.Procedures(2).Speed = T.Elapsed
    End If
    
    ' -PROCEDURE 4---------------------------------------------------
    If G.ProcedureCount > 3 Then
        T.Reset
        For i = 1 To lRepeats
           sBin = Dec2Bin4(lDec)
        Next
        G.Procedures(3).Speed = T.Elapsed
    End If
    
    ' -PROCEDURE 5---------------------------------------------------
    If G.ProcedureCount > 4 Then
        T.Reset
        For i = 1 To lRepeats
           sBin = Dec2Bin5(lDec)
        Next
        G.Procedures(4).Speed = T.Elapsed
    End If
    ' ---------------------------------------------------------------
    
    Set T = Nothing

End Sub

Private Function Dec2Bin1(ByVal d As Long) As String

' This version accomidates doubles.

    Dim b As String

    Do
        b = (CInt((d / 2) = Int(d / 2)) + 1) & b
        d = Int(d * 0.5)
    Loop While d

    Dec2Bin1 = b
    
End Function

Public Function Dec2Bin2(ByVal d As Long) As String

' This version only works with longs.

    Dim b As String
    
    Do
        b = (d And 1) & b
        d = d \ 2
    Loop Until d = 0
    
    Dec2Bin2 = b
    
End Function

Public Function Dec2Bin3(ByVal d As Long) As String

' Hmm.. I wrote this function thinking it would be slow just to show
' the difference in speeds, but - aha! - it seems to be the second fastest.

' Any ideas why?  Maybe the implicit conversion of (d And 1) to a string
' in Dec2Bin2 is slow?
    
    Dim b As String
    
    Do
        If d And 1 Then b = "1" & b Else b = "0" & b
        d = d \ 2
    Loop Until d = 0
    
    Dec2Bin3 = b
    
End Function

Public Function Dec2Bin4(ByVal d As Long) As String

' OK, this function was supposed to be the slowest, but now it's in first place.

    Dim b As String, i As Integer
    
    i = 32
    b = String$(i, "$")
     
    Do
        If d And 1 Then Mid(b, i, 1) = "1" Else Mid(b, i, 1) = "0"
        d = d \ 2
        i = i - 1
    Loop Until d = 0
    
    Dec2Bin4 = Right$(b, 32 - i)
    
End Function

Public Function Dec2Bin5(ByVal d As Long) As String

' Yeees, an actual slow version!  Proving how not to use "IIf" in time-
' critical code.
    
' Warning: Inefficient code ahead, for testing purposes only :)

    Dim b As String
  
    Do
        b = IIf(d And 1, "1", "0") & b
        d = d \ 2
    Loop While d
    
    Dec2Bin5 = b
    
End Function
