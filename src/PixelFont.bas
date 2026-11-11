Attribute VB_Name = "PixelFont"
Option Explicit

' 3x5 pixel font (A-Z, 0-9, a few symbols). Text is drawn with a 1px dark shadow.

Private mGlyph(0 To 127) As String       ' 15 chars of "0"/"1", row-major 3 wide x 5 tall
Private mReady As Boolean

Private Sub Build()
    Dim i As Long
    For i = 0 To 127
        mGlyph(i) = "000000000000000"
    Next i
    SetGlyph "A", "010101111101101"
    SetGlyph "B", "110101110101110"
    SetGlyph "C", "011100100100011"
    SetGlyph "D", "110101101101110"
    SetGlyph "E", "111100110100111"
    SetGlyph "F", "111100110100100"
    SetGlyph "G", "011100101101011"
    SetGlyph "H", "101101111101101"
    SetGlyph "I", "111010010010111"
    SetGlyph "J", "001001001101010"
    SetGlyph "K", "101101110101101"
    SetGlyph "L", "100100100100111"
    SetGlyph "M", "101111111101101"
    SetGlyph "N", "110101101101101"
    SetGlyph "O", "010101101101010"
    SetGlyph "P", "110101110100100"
    SetGlyph "Q", "010101101110011"
    SetGlyph "R", "110101110101101"
    SetGlyph "S", "011100010001110"
    SetGlyph "T", "111010010010010"
    SetGlyph "U", "101101101101111"
    SetGlyph "V", "101101101101010"
    SetGlyph "W", "101101111111101"
    SetGlyph "X", "101101010101101"
    SetGlyph "Y", "101101010010010"
    SetGlyph "Z", "111001010100111"
    SetGlyph "0", "111101101101111"
    SetGlyph "1", "010110010010111"
    SetGlyph "2", "110001010100111"
    SetGlyph "3", "110001010001110"
    SetGlyph "4", "101101111001001"
    SetGlyph "5", "111100110001110"
    SetGlyph "6", "011100111101111"
    SetGlyph "7", "111001010010010"
    SetGlyph "8", "111101111101111"
    SetGlyph "9", "111101111001110"
    SetGlyph "!", "010010010000010"
    SetGlyph "<", "001010100010001"
    SetGlyph ">", "100010001010100"
    SetGlyph "-", "000000111000000"
    SetGlyph ".", "000000000000010"
    SetGlyph ":", "000010000010000"
    SetGlyph "/", "001001010100100"
    SetGlyph "=", "000111000111000"
    mReady = True
End Sub

Private Sub SetGlyph(ByVal ch As String, ByVal bits As String)
    mGlyph(Asc(ch)) = bits
End Sub

Public Function TextWidth(ByVal s As String, ByVal px As Long) As Long
    If Len(s) = 0 Then Exit Function
    TextWidth = (Len(s) * 4 - 1) * px
End Function

' Draw text with its top-left at (x, y). alignMode: 0 left, 1 center, 2 right. px = pixel size.
Public Sub DrawText(ByVal s As String, ByVal x As Long, ByVal y As Long, ByVal clr As Long, _
                    Optional ByVal px As Long = 1, Optional ByVal alignMode As Long = 0)
    Dim k As Long, ch As Long, bits As String, r As Long, c As Long, ox As Long, layer As Long, col As Long
    If Not mReady Then Build
    s = UCase$(s)
    If alignMode = 1 Then x = x - TextWidth(s, px) \ 2
    If alignMode = 2 Then x = x - TextWidth(s, px)
    For layer = 0 To 1                       ' 0 = shadow, 1 = text
        If layer = 0 Then col = RGB(12, 10, 18) Else col = clr
        For k = 1 To Len(s)
            ch = Asc(Mid$(s, k, 1))
            If ch < 128 Then
                bits = mGlyph(ch)
                ox = x + (k - 1) * 4 * px
                For r = 0 To 4
                    For c = 0 To 2
                        If Mid$(bits, r * 3 + c + 1, 1) = "1" Then
                            If layer = 0 Then
                                Renderer.FillRect ox + c * px + 1, y + r * px + 1, px, px, col
                            Else
                                Renderer.FillRect ox + c * px, y + r * px, px, px, col
                            End If
                        End If
                    Next c
                Next r
            End If
        Next k
    Next layer
End Sub
