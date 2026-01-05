Option Explicit
' =====================================================
' Delimiter-based LEFT / RIGHT / MID (Add-in Edition)
'
' Author : Akshay Solanki 
' Date   : 2026-01-05
' Version: 1.0
'
' Rules:
' - Delimiter-based (not char-based)
' - Count always from LEFT (positive only)
' - Space-safe delimiters (no Trim on delimiter)
' - Fast (InStr-based)
' - Clear boundary behavior (returns "")
' =====================================================

' ---------------- LEFT ----------------
Public Function TextLeft(ByVal txt As String, ByVal delim As String, ByVal n As Long) As String
    Dim p As Long
    If n <= 0 Or Len(delim) = 0 Then Exit Function

    p = NthDelimPos(txt, delim, n)
    If p = 0 Then Exit Function

    TextLeft = Trim$(Left$(txt, p - 1))
End Function

' ---------------- RIGHT ----------------
Public Function TextRight(ByVal txt As String, ByVal delim As String, ByVal n As Long) As String
    Dim p As Long
    If n <= 0 Or Len(delim) = 0 Then Exit Function

    p = NthDelimPos(txt, delim, n)
    If p = 0 Then Exit Function

    TextRight = Trim$(Mid$(txt, p + Len(delim)))
End Function

' ---------------- MID ----------------
' Between N1th and N2th delimiter (example: 2,3)
Public Function TextMid(ByVal txt As String, ByVal delim As String, ByVal n1 As Long, ByVal n2 As Long) As String
    Dim p1 As Long, p2 As Long

    If n1 <= 0 Or n2 <= n1 Or Len(delim) = 0 Then Exit Function

    p1 = NthDelimPos(txt, delim, n1)
    p2 = NthDelimPos(txt, delim, n2)

    If p1 = 0 Or p2 = 0 Or p2 <= p1 Then Exit Function

    TextMid = Trim$(Mid$(txt, p1 + Len(delim), p2 - p1 - Len(delim)))
End Function

' -----------------------------------------------------
' Fast Nth delimiter position finder (InStr-based)
' -----------------------------------------------------
Private Function NthDelimPos(ByVal txt As String, ByVal delim As String, ByVal n As Long) As Long
    Dim i As Long, pos As Long

    pos = 0
    For i = 1 To n
        pos = InStr(pos + 1, txt, delim, vbBinaryCompare)
        If pos = 0 Then Exit For
    Next i

    NthDelimPos = pos
End Function
