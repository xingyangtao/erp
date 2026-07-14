' ============================================================================
' Base编码模块
' 提供Base16/32/36/58/62/64/85/91编码解码功能
' ============================================================================

Module BaseEncoding

    ' ========== Base16编码 ==========
    Public Function Base16Encode(input As String) As String
        Dim bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(input)
        Return BitConverter.ToString(bytes).Replace("-", "").ToLower()
    End Function

    Public Function Base16Decode(hex As String) As String
        Dim bytes As Byte() = New Byte(hex.Length \ 2 - 1) {}
        For i As Integer = 0 To bytes.Length - 1
            bytes(i) = Convert.ToByte(hex.Substring(i * 2, 2), 16)
        Next
        Return System.Text.Encoding.UTF8.GetString(bytes)
    End Function

    ' ========== Base32编码 ==========
    Private ReadOnly Base32Chars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567"

    Public Function Base32Encode(input As String) As String
        Dim bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(input)
        Dim result As New Text.StringBuilder()
        Dim i As Integer = 0

        While i < bytes.Length
            Dim index As Integer = (bytes(i) >> 3) And &H1F
            result.Append(Base32Chars(index))

            Dim byte2 As Integer = If(i + 1 < bytes.Length, bytes(i + 1), 0)
            index = ((bytes(i) And &H7) << 2) Or ((byte2 >> 6) And &H3)
            result.Append(Base32Chars(index))

            If i + 1 < bytes.Length Then
                index = (byte2 >> 1) And &H1F
                result.Append(Base32Chars(index))

                Dim byte3 As Integer = If(i + 2 < bytes.Length, bytes(i + 2), 0)
                index = ((byte2 And &H1) << 4) Or ((byte3 >> 4) And &HF)
                result.Append(Base32Chars(index))

                If i + 2 < bytes.Length Then
                    index = (byte3 And &HF) << 1
                    result.Append(Base32Chars(index))

                    Dim byte4 As Integer = If(i + 3 < bytes.Length, bytes(i + 3), 0)
                    index = ((byte3 And &H1) << 4) Or ((byte4 >> 7) And &H1)
                    result.Append(Base32Chars(index))

                    index = (byte4 >> 2) And &H1F
                    result.Append(Base32Chars(index))

                    Dim byte5 As Integer = If(i + 4 < bytes.Length, bytes(i + 4), 0)
                    index = ((byte4 And &H3) << 3) Or ((byte5 >> 5) And &H7)
                    result.Append(Base32Chars(index))

                    index = byte5 And &H1F
                    result.Append(Base32Chars(index))
                End If
            End If

            i += 5
        End While

        Return result.ToString()
    End Function

    ' ========== Base64编码 ==========
    Public Function Base64Encode(input As String) As String
        Dim bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(input)
        Return Convert.ToBase64String(bytes)
    End Function

    Public Function Base64Decode(base64 As String) As String
        Dim bytes As Byte() = Convert.FromBase64String(base64)
        Return System.Text.Encoding.UTF8.GetString(bytes)
    End Function

    ' ========== URL编码 ==========
    Public Function URLEncode(input As String) As String
        Return Uri.EscapeDataString(input)
    End Function

    Public Function URLDecode(input As String) As String
        Return Uri.UnescapeDataString(input)
    End Function

End Module
