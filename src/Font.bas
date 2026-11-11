Attribute VB_Name = "Font"
Option Explicit

' 3x5 pixel font (A-Z, 0-9, a few symbols). Text is drawn with a 1px dark shadow.

Private mGlyph(0 To 127) As String       ' 15 chars of "0"/"1", row-major 3 wide x 5 tall
Private mReady As Boolean

Private Sub Build()
    Dim i As Long
    For i = 0 To 127
        mGlyph(i) = "000000000000000"
    Next i
    G "A", "010101111101101": G "B", "110101110101110": G "C", "011100100100011"
    G "D", "110101101101110": G "E", "111100110100111": G "F", "111100110100100"
    G "G", "011100101101011": G "H", "101101111101101": G "I", "111010010010111"
    G "J", "001001001101010": G "K", "101101110101101": G "L", "100100100100111"
    G "M", "101111111101101": G "N", "110101101101101": G "O", "010101101101010"
    G "P", "110101110100100": G "Q", "010101101110011": G "R", "110101110101101"
    G "S", "011100010001110": G "T", "111010010010010": G "U", "101101101101111"
    G "V", "101101101101010": G "W", "101101111111101": G "X", "101101010101101"
    G "Y", "101101010010010": G "Z", "111001010100111"
    G "0", "111101101101111": G "1", "010110010010111": G "2", "110001010100111"
    G "3", "110001010001110": G "4", "101101111001001": G "5", "111100110001110"
    G "6", "011100111101111": G "7", "111001010010010": G "8", "111101111101111"
    G "9", "111101111001110"
    G "!", "010010010000010": G "<", "001010100010001": G ">", "100010001010100"
    G "-", "000000111000000": G ".", "000000000000010": G ":", "000010000010000"
    G "/", "001001010100100"
    G "=", "000111000111000"
    mReady = True
End Sub

Private Sub G(ByVal ch As String, ByVal bits As String)
    mGlyph(Asc(ch)) = bits
End Sub

Public Function TextWidth(ByVal s As String, ByVal scale As Long) As Long
    If Len(s) = 0 Then Exit Function
    TextWidth = (Len(s) * 4 - 1) * scale
End Function

' Draw text with its top-left at (x, y). align: 0 left, 1 center, 2 right.
Public Sub DrawText(ByVal s As String, ByVal x As Long, ByVal y As Long, ByVal color As Long, _
                    Optional ByVal scale As Long = 1, Optional ByVal align As Long = 0)
    Dim k As Long, ch As Long, bits As String, r As Long, c As Long, ox As Long, pass As Long, col As Long
    If Not mReady Then Build
    s = UCase$(s)
    If align = 1 Then x = x - TextWidth(s, scale) \ 2
    If align = 2 Then x = x - TextWidth(s, scale)
    For pass = 0 To 1                       ' 0 = shadow, 1 = text
        If pass = 0 Then col = RGB(12, 10, 18) Else col = color
        For k = 1 To Len(s)
            ch = Asc(Mid$(s, k, 1))
            If ch < 128 Then
                bits = mGlyph(ch)
                ox = x + (k - 1) * 4 * scale
                For r = 0 To 4
                    For c = 0 To 2
                        If Mid$(bits, r * 3 + c + 1, 1) = "1" Then
                            If pass = 0 Then
                                Renderer.FillRect ox + c * scale + 1, y + r * scale + 1, scale, scale, col
                            Else
                                Renderer.FillRect ox + c * scale, y + r * scale, scale, scale, col
                            End If
                        End If
                    Next c
                Next r
            End If
        Next k
    Next pass
End Sub
