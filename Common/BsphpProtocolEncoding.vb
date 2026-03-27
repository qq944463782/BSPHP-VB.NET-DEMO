Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text

Namespace Global.BSPHP.Protocol
    Public Module BsphpProtocolEncoding
        Public Function UrlEncodeCxx(value As String) As String
            If value Is Nothing Then Return ""
            Dim sb As New StringBuilder()
            Dim bytes As Byte() = Encoding.UTF8.GetBytes(value)
            For Each b As Byte In bytes
                Dim isUnreserved As Boolean =
                    (b >= AscW("0"c) AndAlso b <= AscW("9"c)) OrElse
                    (b >= AscW("A"c) AndAlso b <= AscW("Z"c)) OrElse
                    (b >= AscW("a"c) AndAlso b <= AscW("z"c)) OrElse
                    b = AscW("-"c) OrElse b = AscW("_"c) OrElse b = AscW("."c) OrElse b = AscW("~"c)
                If isUnreserved Then
                    sb.Append(ChrW(b))
                Else
                    sb.Append("%"c)
                    sb.Append(b.ToString("x2"))
                End If
            Next
            Return sb.ToString()
        End Function

        Public Function UrlDecodeCxx(value As String) As String
            If String.IsNullOrEmpty(value) Then Return ""
            Dim bytes As New List(Of Byte)()
            Dim i As Integer = 0
            While i < value.Length
                Dim c As Char = value(i)
                If c = "%"c AndAlso i + 2 < value.Length Then
                    Dim hex As String = value.Substring(i + 1, 2)
                    Dim n As Integer
                    If Integer.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, n) Then
                        bytes.Add(CByte(n And &HFF))
                        i += 3
                        Continue While
                    End If
                ElseIf c = "+"c Then
                    bytes.Add(CByte(AscW(" "c)))
                    i += 1
                    Continue While
                End If
                Dim chBytes As Byte() = Encoding.UTF8.GetBytes(New Char() {c})
                bytes.AddRange(chBytes)
                i += 1
            End While
            Return Encoding.UTF8.GetString(bytes.ToArray())
        End Function

        Public Function BuildPlainQueryString(values As Dictionary(Of String, String)) As String
            Dim keys As New List(Of String)(values.Keys)
            keys.Sort(StringComparer.Ordinal)
            Dim sb As New StringBuilder()
            For idx As Integer = 0 To keys.Count - 1
                Dim k As String = keys(idx)
                If idx > 0 Then sb.Append("&"c)
                sb.Append(UrlEncodeCxx(k))
                sb.Append("="c)
                sb.Append(UrlEncodeCxx(values(k)))
            Next
            Return sb.ToString()
        End Function

        Public Function TrySplitEncryptedResponse(raw As String, ByRef encPart As String, ByRef rsaPart As String) As Boolean
            encPart = ""
            rsaPart = ""
            If String.IsNullOrEmpty(raw) Then Return False
            Dim firstPipe As Integer = raw.IndexOf("|"c)
            If firstPipe < 0 Then Return False
            Dim secondPipe As Integer = raw.IndexOf("|"c, firstPipe + 1)
            Dim firstPart As String = raw.Substring(0, firstPipe)
            Dim hasOkPrefix As Boolean = firstPart.Length >= 2 AndAlso firstPart.Substring(0, 2).Equals("ok", StringComparison.OrdinalIgnoreCase)
            If hasOkPrefix AndAlso secondPipe > firstPipe Then
                encPart = raw.Substring(firstPipe + 1, secondPipe - firstPipe - 1)
                rsaPart = raw.Substring(secondPipe + 1)
            Else
                encPart = firstPart
                rsaPart = raw.Substring(firstPipe + 1)
            End If
            Return True
        End Function
    End Module
End Namespace
