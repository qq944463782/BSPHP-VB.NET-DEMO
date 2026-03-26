Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Net.Http
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading.Tasks
Imports System.Text.Json

Namespace BSPHP.LoginDemoVB
    ' Card demo local protocol client (SDK mode, full encrypted request/response).
    Public NotInheritable Class BSPHPConfig
        Public Property Url As String = ""
        Public Property MutualKey As String = ""
        Public Property ServerPrivateKey As String = ""
        Public Property ClientPublicKey As String = ""
        Public Property CodeUrlPrefix As String = "https://demo.bsphp.com/index.php?m=coode&sessl="
    End Class

    Public Class ApiResult
        Public Property Code As Integer?
        Public Property DataText As String = ""
    End Class

    Public Class BSPHPClient
        Private ReadOnly _http As New HttpClient()
        Private ReadOnly _cfg As BSPHPConfig
        Public Property Sessl As String = ""

        Public Sub New(cfg As BSPHPConfig)
            _cfg = cfg
        End Sub

        Public Shared Function MachineCode() As String
            Return Environment.MachineName & "_" & Environment.UserName
        End Function

        Private Shared Function Md5Hex(s As String) As String
            Dim bytes = Encoding.UTF8.GetBytes(s)
            Dim hash = MD5.HashData(bytes)
            Dim sb As New StringBuilder()
            For Each b In hash
                sb.Append(b.ToString("x2"))
            Next
            Return sb.ToString()
        End Function

        ' 计算当前运行程序文件的 MD5（用于协议中的 md5 字段）。
        Private Shared Function CurrentAppMd5Hex() As String
            Try
                Dim path = Environment.ProcessPath
                If String.IsNullOrWhiteSpace(path) OrElse Not File.Exists(path) Then Return ""
                Using fs = File.OpenRead(path)
                    Dim hash = MD5.HashData(fs)
                    Dim sb As New StringBuilder()
                    For Each b In hash
                        sb.Append(b.ToString("x2"))
                    Next
                    Return sb.ToString()
                End Using
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function AesEncryptBase64(plain As String, key16 As String) As String
            Using aes = Aes.Create()
                aes.Mode = CipherMode.CBC
                aes.Padding = PaddingMode.PKCS7
                aes.Key = Encoding.UTF8.GetBytes(key16)
                aes.IV = Encoding.UTF8.GetBytes(key16)
                Dim p = Encoding.UTF8.GetBytes(plain)
                Using enc = aes.CreateEncryptor()
                    Return Convert.ToBase64String(enc.TransformFinalBlock(p, 0, p.Length))
                End Using
            End Using
        End Function

        Private Shared Function AesDecryptBase64(cipherB64 As String, key16 As String) As String
            Using aes = Aes.Create()
                aes.Mode = CipherMode.CBC
                aes.Padding = PaddingMode.PKCS7
                aes.Key = Encoding.UTF8.GetBytes(key16)
                aes.IV = Encoding.UTF8.GetBytes(key16)
                Dim c = Convert.FromBase64String(cipherB64)
                Using dec = aes.CreateDecryptor()
                    Dim plain = dec.TransformFinalBlock(c, 0, c.Length)
                    Return Encoding.UTF8.GetString(plain)
                End Using
            End Using
        End Function

        Private Shared Function RsaEncryptPkcs1Base64(message As String, publicKeyB64Der As String) As String
            Using rsa = RSA.Create()
                rsa.ImportSubjectPublicKeyInfo(Convert.FromBase64String(publicKeyB64Der), Nothing)
                Dim enc = rsa.Encrypt(Encoding.UTF8.GetBytes(message), RSAEncryptionPadding.Pkcs1)
                Return Convert.ToBase64String(enc)
            End Using
        End Function

        Private Shared Function RsaDecryptPkcs1Base64(cipherB64 As String, privateKeyB64Der As String) As String
            Using rsa = RSA.Create()
                Dim keyBytes = Convert.FromBase64String(privateKeyB64Der)
                Try
                    rsa.ImportRSAPrivateKey(keyBytes, Nothing)
                Catch
                    rsa.ImportPkcs8PrivateKey(keyBytes, Nothing)
                End Try
                Dim plain = rsa.Decrypt(Convert.FromBase64String(cipherB64), RSAEncryptionPadding.Pkcs1)
                Return Encoding.UTF8.GetString(plain)
            End Using
        End Function

        Public Async Function BootstrapAsync() As Task(Of String)
            Dim c = Await SendAsync("internet.in")
            If c.DataText <> "1" Then Return "连接失败"
            Dim s = Await SendAsync("BSphpSeSsL.in")
            Sessl = s.DataText
            Return If(String.IsNullOrWhiteSpace(Sessl), "获取会话失败", "")
        End Function

        Public Function GetCodeImageUrl() As String
            If String.IsNullOrWhiteSpace(Sessl) Then Return ""
            Return _cfg.CodeUrlPrefix & Sessl
        End Function

        Public Async Function SendAsync(api As String, Optional extra As Dictionary(Of String, String) = Nothing) As Task(Of ApiResult)
            Dim now = DateTime.Now
            Dim appSafe = Md5Hex(now.ToString("yyyy-MM-dd HH:mm:ss"))
            Dim p As New Dictionary(Of String, String) From {
                {"api", api},
                {"BSphpSeSsL", Sessl},
                {"date", now.ToString("yyyy-MM-dd#HH:mm:ss")},
                {"md5", CurrentAppMd5Hex()},
                {"mutualkey", _cfg.MutualKey},
                {"appsafecode", appSafe}
            }
            If extra IsNot Nothing Then
                For Each kv In extra
                    p(kv.Key) = kv.Value
                Next
            End If

            Dim parts As New List(Of String)()
            For Each kv In p
                parts.Add(kv.Key & "=" & Uri.EscapeDataString(kv.Value))
            Next
            Dim dataStr = String.Join("&", parts)
            Dim aesKey = Md5Hex(_cfg.ServerPrivateKey & appSafe).Substring(0, 16)
            Dim encB64 = AesEncryptBase64(dataStr, aesKey)
            Dim sig = "0|AES-128-CBC|" & aesKey & "|" & Md5Hex(encB64) & "|json"
            Dim rsaB64 = RsaEncryptPkcs1Base64(sig, _cfg.ClientPublicKey)
            Dim payload = Uri.EscapeDataString(encB64 & "|" & rsaB64)
            Dim body = New StringContent("parameter=" & payload, Encoding.UTF8, "application/x-www-form-urlencoded")
            Dim resp = Await _http.PostAsync(_cfg.Url, body)
            Dim raw = Await resp.Content.ReadAsStringAsync()
            Dim u = Uri.UnescapeDataString(raw)
            Dim seg = u.Split("|"c)
            If seg.Length < 3 Then Return New ApiResult With {.DataText = "响应格式错误"}

            Dim sigPlain = RsaDecryptPkcs1Base64(seg(2).Trim(), _cfg.ServerPrivateKey)
            Dim sigSeg = sigPlain.Split("|"c)
            Dim respKey = sigSeg(2).Substring(0, 16)
            Dim jsonPlain = AesDecryptBase64(seg(1).Trim(), respKey)

            Using doc = JsonDocument.Parse(jsonPlain)
                Dim res = doc.RootElement.GetProperty("response")
                Dim r As New ApiResult()
                Dim cProp As JsonElement
                If res.TryGetProperty("code", cProp) Then
                    If cProp.ValueKind = JsonValueKind.Number Then r.Code = cProp.GetInt32()
                End If
                Dim d As JsonElement
                If res.TryGetProperty("data", d) Then
                    If d.ValueKind = JsonValueKind.String Then
                        r.DataText = d.GetString()
                    Else
                        r.DataText = d.ToString()
                    End If
                End If
                Dim ssl As JsonElement
                If res.TryGetProperty("SeSsL", ssl) AndAlso (r.Code = 1011 OrElse r.Code = 9908 OrElse r.Code = 1081) Then
                    Sessl = ssl.GetString()
                End If
                Return r
            End Using
        End Function
    End Class
End Namespace
