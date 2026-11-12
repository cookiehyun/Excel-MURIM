Attribute VB_Name = "PixelFont"
Option Explicit

' 3x5 pixel font (A-Z, 0-9, a few symbols). Text is drawn with a 1px dark shadow.

Private mGlyph(0 To 127) As String       ' 15 chars of "0"/"1", row-major 3 wide x 5 tall
Private mReady As Boolean
Private mBig(0 To 127) As String         ' 35 chars, 5 wide x 7 tall (title letters)

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
    For i = 0 To 127
        mBig(i) = ""
    Next i
    SetBig "E", "11111100001000011110100001000011111"
    SetBig "X", "10001100010101000100010101000110001"
    SetBig "C", "01111100001000010000100001000001111"
    SetBig "L", "10000100001000010000100001000011111"
    SetBig "M", "10001110111010110101100011000110001"
    SetBig "U", "10001100011000110001100011000101110"
    SetBig "H", "10001100011000111111100011000110001"
    SetBig "Y", "10001100010101000100001000010000100"
    SetBig "O", "01110100011000110001100011000101110"
    SetBig "P", "11110100011000111110100001000010000"
    SetBig "A", "01110100011000111111100011000110001"
    SetBig "T", "11111001000010000100001000010000100"
    SetBig "R", "11110100011000111110101001001010001"
    SetBig "N", "10001110011010110011100011000110001"
    SetBig "I", "01110001000010000100001000010001110"
    SetBig "G", "01111100001000010111100011000101111"
    SetBig "S", "01111100001000001110000010000111110"
    SetBig "V", "10001100011000110001100010101000100"
    SetBig "W", "10001100011000110101101011101110001"
    SetBig "K", "10001100101010011000101001001010001"
    SetBig "D", "11110100011000110001100011000111110"
    SetBig "F", "11111100001000011110100001000010000"
    SetBig "B", "11110100011000111110100011000111110"
    SetBig " ", "00000000000000000000000000000000000"
    mReady = True
End Sub

Private Sub SetGlyph(ByVal ch As String, ByVal bits As String)
    mGlyph(Asc(ch)) = bits
End Sub

Private Sub SetBig(ByVal ch As String, ByVal bits As String)
    mBig(Asc(ch)) = bits
End Sub

' Large 5x7 letters (titles). Letters without a 5x7 shape fall back to 3x5. Centered on x.
Public Sub DrawBig(ByVal s As String, ByVal x As Long, ByVal y As Long, ByVal clr As Long, ByVal px As Long, _
                   Optional ByVal shadow As Long = -1)
    Dim k As Long, ch As Long, bits As String, r As Long, c As Long, ox As Long, layer As Long, col As Long
    If Not mReady Then Build
    s = UCase$(s)
    x = x - ((Len(s) * 6 - 1) * px) \ 2
    If shadow < 0 Then shadow = RGB(12, 10, 18)
    For layer = 0 To 1
        If layer = 0 Then col = shadow Else col = clr
        For k = 1 To Len(s)
            ch = Asc(Mid$(s, k, 1))
            ox = x + (k - 1) * 6 * px + (1 - layer) * px \ 2 + (1 - layer)
            If ch < 128 Then
                bits = mBig(ch)
                If Len(bits) = 35 Then
                    For r = 0 To 6
                        For c = 0 To 4
                            If Mid$(bits, r * 5 + c + 1, 1) = "1" Then
                                Renderer.FillRect ox + c * px, y + r * px + (1 - layer) * (px \ 2 + 1), px + IIf(px >= 3, 1, 0), px, col
                            End If
                        Next c
                    Next r
                End If
            End If
        Next k
    Next layer
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
