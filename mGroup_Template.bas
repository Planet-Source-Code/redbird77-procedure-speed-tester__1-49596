Attribute VB_Name = "mGroup_Template"
' mGroup_Template.bas

Option Explicit

Public Sub RunTests(ByRef G As tGroup)

Dim i As Long, lRepeats As Long, T As cStopwatch

    Set T = New cStopwatch
    lRepeats = G.RepeatCount
    
' Declare variables here that are unique to the procedures to be tested.
    Dim bRet As Boolean
    Dim a As Integer, b As Integer, c As Integer
    
    ' -PROCEDURE 1---------------------------------------------------
    If G.ProcedureCount > 0 Then
        T.Reset
        For i = 1 To lRepeats
           bRet = Function1(a, b, c)
        Next
        G.Procedures(0).Speed = T.Elapsed
    End If
    
    ' -PROCEDURE 2---------------------------------------------------
    If G.ProcedureCount > 1 Then
    T.Reset
    For i = 1 To lRepeats
       bRet = Function2(a, b, c)
    Next
    G.Procedures(1).Speed = T.Elapsed
    End If

    ' -PROCEDURE 3---------------------------------------------------
    If G.ProcedureCount > 2 Then
    T.Reset
    For i = 1 To lRepeats
       bRet = Function3(a, b, c)
    Next
    G.Procedures(2).Speed = T.Elapsed
    End If

    ' -PROCEDURE 4---------------------------------------------------
    
    ' Empty.
    
    ' -PROCEDURE 5---------------------------------------------------
    
    ' Empty.
    
    ' ---------------------------------------------------------------
    
    Set T = Nothing

End Sub

' #PROCEDURE1#
Private Function Function1(a As Integer, b As Integer, c As Integer) As Boolean
    Function1 = True
End Function

' #PROCEDURE2#
Private Function Function2(a As Integer, b As Integer, c As Integer) As Boolean
    Function2 = True
End Function

' #PROCEDURE3#
Private Function Function3(a As Integer, b As Integer, c As Integer) As Boolean
    Function3 = True
End Function
