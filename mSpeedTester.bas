Attribute VB_Name = "mSpeedTester"
' file    : mSpeedTester.bas
' revised : 2003-10-31
' author  : redbird77
' email   : redbird77@earthlink.net
' www     : http://home.earthlink.net/~redbird77

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const MAX_FUNC As Integer = 5

Public Type tProcedure
    Name      As String
    Speed     As Long
    SpeedNorm As Single
End Type

Public Type tGroup
    Name           As String
    ProcedureCount As Integer
    RepeatCount    As Long
    BestIndex      As Integer
    Procedures()   As tProcedure
End Type

Public Sub SpeedTester_Run(ByRef G As tGroup)

' This is where the speed testing begins.  All the user needs to do
' is edit the lines that begin as "EDIT:"
    
    ' EDIT: Module where procedures are located.
    'mGroup_LNGtoRGB.RunTests G
    mGroup_Dec2Bin.RunTests G
    
    pNormalizeSpeeds G
    
End Sub

Public Sub SpeedTester_Init(ByRef G As tGroup, ByVal sGroupName As String, _
                            ByVal iProcCount As Integer, ByVal lRepeats As Long)

Dim i As Integer

    G.Name = sGroupName
    G.RepeatCount = lRepeats
    G.ProcedureCount = iProcCount
    
    ReDim G.Procedures(iProcCount - 1)
    
    For i = 0 To UBound(G.Procedures)
        G.Procedures(i).Name = "Procedure #" & (i + 1)
    Next
    
End Sub

Public Function SpeedTester_ToHTML(ByRef G As tGroup, ByVal blnToClip As Boolean) As String

' This function can be completely user-edited.  The way I have done things
' is purely arbitrary.  As long as this function returns a string (or maybe
' writes to a file), everything is OK.

Dim i As Integer, sHTM As String, sRow As String, sTmp As String

    ' My default row.
    sRow = "<tr><td align=""right"">#PROC_NAME#</td>" & _
           "<td><span style=""margin: 0px; width: #BAR_WIDTH#px;" & _
           "background: #BAR_COLOR#;"">#BAR_CAPTION#</span></td>" & _
           "<td>#TIME_CAPTION#</td></tr>"
    
    sHTM = sHTM & "<h2>" & G.Name & " Results</h2>" & vbCrLf & "<table>" & vbCrLf
    
    For i = 0 To G.ProcedureCount - 1
    
        sTmp = sRow
    
        sTmp = Replace(sTmp, "#PROC_NAME#", G.Procedures(i).Name)
        sTmp = Replace(sTmp, "#BAR_WIDTH#", Int(IIf(i = G.BestIndex, 1, G.Procedures(i).SpeedNorm) * 150))
        sTmp = Replace(sTmp, "#BAR_COLOR#", IIf(i = G.BestIndex, "#ffffff", LNGtoHEX(HUEtoLNG(2 - G.Procedures(i).SpeedNorm * 2))))
        sTmp = Replace(sTmp, "#BAR_CAPTION#", IIf(i = G.BestIndex, "FASTEST!", ""))
        sTmp = Replace(sTmp, "#TIME_CAPTION#", G.Procedures(i).Speed & " (" & Format$(G.Procedures(i).Speed / G.RepeatCount, ".000") & ")")
        
        sHTM = sHTM & sTmp & vbCrLf
    Next
    
    sHTM = sHTM & "</table>"
    
    If blnToClip Then
        Clipboard.Clear
        Clipboard.SetText sHTM
    End If
        
    SpeedTester_ToHTML = sHTM
    
End Function

Public Sub SpeedTester_Graph(ByRef G As tGroup, vl As Variant, vt As Variant)

Dim i As Integer

    ' Set each label's properties based on procedure's time.
    For i = 0 To G.ProcedureCount - 1
    
        With vl(i)
            .ForeColor = vbButtonText
            .BackColor = IIf(i = G.BestIndex, vbButtonFace, HUEtoLNG(2 - G.Procedures(i).SpeedNorm * 2))
            .BorderStyle = IIf(i = G.BestIndex, 0, 1)
            .Caption = IIf(i = G.BestIndex, "FASTEST!", "")
            .Width = IIf(i = G.BestIndex, 1, G.Procedures(i).SpeedNorm) * 150 * Screen.TwipsPerPixelX
        End With
        
        vt(i).Caption = G.Procedures(i).Speed & " (" & Format$(G.Procedures(i).Speed / G.RepeatCount, ".000") & ")"
        vt(i).ForeColor = vbButtonText
        
    Next
    
    ' Set each unused label to a disabled look.
    For i = G.ProcedureCount To MAX_FUNC - 1
        With vl(i)
            .BackColor = vbButtonFace
            .ForeColor = vbGrayText
            .BorderStyle = 0
            .Caption = "N/A"
        End With
        vt(i).Caption = "N/A"
        vt(i).ForeColor = vbGrayText
    Next
    
End Sub

Private Sub pNormalizeSpeeds(ByRef G As tGroup)

' Normalize and store speeds.
Dim i As Long, lo As Long, hi As Long, hi_idx As Integer

    ' Find fastest procedure time.
    G.BestIndex = 0
    lo = G.Procedures(0).Speed: hi = lo
 
    For i = 1 To G.ProcedureCount - 1
    
        If G.Procedures(i).Speed < lo Then
            lo = G.Procedures(i).Speed
            G.BestIndex = i
        End If
        
        If G.Procedures(i).Speed > hi Then
            hi = G.Procedures(i).Speed
            hi_idx = i
        End If
        
    Next
    
    ' Normalize procedure times to be between 0 and 1.
    
    ' If only testing one procedure then just use unnormalized speed.
    If G.ProcedureCount = 1 Then
        G.Procedures(0).SpeedNorm = G.Procedures(0).Speed
    Else
        For i = 0 To G.ProcedureCount - 1
            G.Procedures(i).SpeedNorm = Normalize(G.Procedures(i).Speed, lo, hi, 0, 1)
        Next
    End If
    
    'Debug.Print "Lo: #" & G.BestIndex & " @ " & G.Procedures(G.BestIndex).Speed & " ms."
    'Debug.Print "Hi: #" & hi_idx & " @ " & G.Procedures(hi_idx).Speed & " ms."
    
End Sub

' -------------------------------------------------------------------
' Helper Functions
' -------------------------------------------------------------------

Public Function LNGtoHEX(ByVal c As Long) As String

    Dim b(2) As Byte
    
    CopyMemory b(0), c, 3
    
    LNGtoHEX = "#" & Right$("00000" & LCase$(Hex$(RGB(b(2), b(1), b(0)))), 6)
    
End Function

Private Function HUEtoLNG(ByVal h As Single) As Long

' ABOUT: This is a nifty function I came up with after studying the output
' of a rainbow gradient.  It looks suspiciously like the beginnings of a
' common HSBtoLNG function.  Aha, it is all starting to make sense!

Dim f As Single: f = (h - Int(h)) * 255
    
    Select Case Int(h) Mod 6
        Case 0: HUEtoLNG = RGB(255, f, 0)
        Case 1: HUEtoLNG = RGB(255 - f, 255, 0)
        Case 2: HUEtoLNG = RGB(0, 255, f)
        Case 3: HUEtoLNG = RGB(0, 255 - f, 255)
        Case 4: HUEtoLNG = RGB(255, 0, 255)
        Case 5: HUEtoLNG = RGB(255, 0, 255 - f)
    End Select

End Function

Public Function Normalize(ByVal uval As Long, _
                          ByVal ulo As Long, ByVal uhi As Long, _
                          ByVal nlo As Long, ByVal nhi As Long) As Single
                          
' ABOUT: Normalize takes a value and using it's position in a given input range,
' returns a value in a given output range.

' Thanx to Dr. Math for help with this function.
    
    Normalize = nlo + (uval - ulo) * (nhi - nlo) / (uhi - ulo)

End Function
