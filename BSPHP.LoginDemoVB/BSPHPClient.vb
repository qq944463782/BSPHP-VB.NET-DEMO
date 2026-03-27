Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Net.Http
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading.Tasks
Imports System.Xml.Linq
Imports System.Management

Namespace BSPHP.LoginDemoVB
    ' BSPHP 连接配置：接口地址、通信密钥、RSA 密钥、验证码地址。
    Public NotInheritable Class BSPHPConfig

    ' 初始化地址
        Public Property Url As String = "https://demo.bsphp.com/AppEn.php?appid=8888888&m=95e87faf2f6e41babddaef60273489e1&lang=0"
        
        ' 通信密钥
        Public Property MutualKey As String = "6600cfcd5ac01b9bb3f2460eb416daa8"
        ' 服务器私钥
        Public Property ServerPrivateKey As String = "MIIEqAIBADANBgkqhkiG9w0BAQEFAASCBJIwggSOAgEAAoH+DEr7H5BhMwRA9ZWXVftcCWHznBdl0gQBu5617qSe9in+uloF1sC64Ybdc8Q0JwQkGQANC5PnMqPLgXXIfnl7/LpnQ/BvghQI5cr/4DEezRKrmQaXgYfXHL3woVw7JIsLpPTGa7Ar9S6SEH8RcPIbZjlPVRZPwV3RgWgox2/4lkXsmopqD+mEtOI/ntvti147nEpK2c7cdtCU5M2hQSlIXsTWvri88RTYJ/CtopBOXarUkNBfpWGImiYGsmbZI+YZ6uU0wSYlq8huu+pkTseUUiymzmv8Rpg3coi7YU+pszvB9wnQ1Rz6Z/B6Z3WN7d6OP7f9w0Q0WvgrsKcEJhMCAwEAAQKB/gHa5t6yiRiL0cm902K0VgVMdNjfZww0cpZ/svDaguqfF8PDhhIMb6dNFOo9d6lTpKbpLQ7MOR2ZPkLBJYqAhsdy0dac2BcHMviKk+afQwirgp3LMt3nQ/0gZMnVA0/Wc+Fm1vK1WUzcxEodAuLKhnv8tg4fGdYSdGVU9KJ0MU1bKQZXv0CAIhJYWsiCa5y5bFO7K+ia+UIVBHcvITQLzlgEm+Z/X6ye5cws4pWbk8+spsBDvweb5jpelbkCYs5C5TRNIWXk7+QxTXTg1vrcsmZRcmpRJq7sOd3faZltNHTIlB3HhWnsf47Bz334j9RtU8iqonbuBmcnYbD3+bvBAn891RGdAl+rVU/sJ2kPXmV4eqJOwJfbi8o1WYDp4GcK0ThjrZ1pmaZMj2WTjb3QX1VUoi+7l3389KzzDn0VXLKXZvGxmLikA1FWuuLUmwfNTxyxtGTBVeZCEaQ2lEJuaDGsK0oLi4Bo8ELfQw6JFK7jlgtTlflcYcul99P9BThDAn8y5TpSQy8/07LCgMMZOgJomYzQUmd14Zn2VQLH1u1Z4v2CPlOzGanDt7mmGZCew7iMSO1P0TrwDIreKzYyERuVvZti/IFHH1+J1hAbvk9SJGmdt46W5lyIp3xjdR2QmiK+hSsc8HF9R+zPaSe9yGA8+FwxLRfo0snGP3MC3aXxAn4n2iyABgejZlkc3EnanfzIqkHygC9gUbkCqa1tEDVZw3+Uv1G1vlJxBftyHuk4ZDmbUu1w+zM41nqiLbRxEE4LR06AKO7Yx0qlm86XOVTN/y9/WcWW1saRzs0IYIZwordhQIV463DYMgLn41B7Cdmu1gZ22TLfWCjpz9HSQosCfwMJu9l9OSzOLjV+CidPVyV3RPiKcrKOrOoPWQMkyTY8XnWP0t82APQ121cW35Mai8GT+NZy3tnFZeStH6cNbmAZ2VSnTfA45zMLHBsL2SBGHCfV9ST8yzk9BifJreIb0UceG9y2XY/k4zXeSQkDFPuOt7IXxv2W14SF9Q+Ou4ECfzfRP1hXPwq2w4YJ8sLmqWJT+3aMDucei5MJEAJNifZWhdW0GIrlKRSbhIgLAunxq+KK+mAPqqWw7Prsa21JbXSe3gugusu5d6ESURvLENRKI+Pp9TgRESsydeLy8VcPKRJ5/Ct7/p6QB3A+7F/iPNE2GagGffG9i7e+OdcToYQ="

        ' 客户端公钥
        Public Property ClientPublicKey As String = "MIIBHjANBgkqhkiG9w0BAQEFAAOCAQsAMIIBBgKB/g26m2hYtESqcKW+95Lr+PfCd4bwHW2Z+mM0/vcKQ5j/ZGMigqkgl3QXCEcsCaw0KFSmqAPtLbrl6p5Sp+ZUSYEYQhSxAajE5qRCd3k0r/MIQQanBaOALkP71/u6U2SZhrTXd05n1wQo6ojMH/xVunBOFOa/Eon/Y5FVh6GiJpwwDkFzTlnecmff7Y+VDqRhZ7vu2CQjApOx23N6DiFEmVZYEb/efyASngoZ+3A/DSB5cwbaYVZ21EhPe/GNcwtUleFHn+d4vb0cvolO3Gyw6ObceOT/Q7E3k8ejIml6vPKzmRdtw0FXGOJTclx1CjShRDfXoUjFGyXHy3sZs9VLAgMBAAE="
        ' 验证码地址前缀
        Public Property CodeUrlPrefix As String = "https://demo.bsphp.com/index.php?m=coode&sessl="
    End Class

    ' 通用接口返回模型：code 为状态码，DataText 为服务器 data 字段文本化结果。
    Public Class ApiResult
        Public Property Code As Integer?
        Public Property DataText As String = ""
    End Class

    ' BSPHP 协议客户端：负责参数组装、加解密、接口调用与业务 API 封装。
    Public Class BSPHPClient
        Private ReadOnly _http As New HttpClient()
        Private ReadOnly _cfg As BSPHPConfig
        Public Property Sessl As String = ""

        ' 构造客户端实例，绑定一组 BSPHP 配置。
        Public Sub New(cfg As BSPHPConfig)
            _cfg = cfg
        End Sub

        ' 本机特征码：用于 login/register 等接口的 key/maxoror 参数。
        Public Shared Function MachineCode() As String
            Try
                Using mc As New ManagementClass("Win32_Processor")
                    For Each mo As ManagementObject In mc.GetInstances()
                        Dim pidObj = mo.Properties("ProcessorId").Value
                        If pidObj IsNot Nothing Then
                            Dim pid = pidObj.ToString().Trim()
                            If pid <> "" Then Return pid
                        End If
                    Next
                End Using
            Catch
            End Try
            Return Environment.MachineName & "_" & Environment.UserName
        End Function

        ' 计算 MD5 十六进制字符串（协议里用于 appsafecode 与摘要签名）。
        Private Shared Function Md5Hex(s As String) As String
            Dim bytes As Byte() = Encoding.UTF8.GetBytes(s)
            Dim hash As Byte()
            Using md5 As MD5 = MD5.Create()
                hash = md5.ComputeHash(bytes)
            End Using
            Dim sb As New StringBuilder()
            For Each b As Byte In hash
                sb.Append(b.ToString("x2"))
            Next
            Return sb.ToString()
        End Function

        ' 计算当前运行程序文件的 MD5（用于协议中的 md5 字段）。
        Private Shared Function CurrentAppMd5Hex() As String
            Try
                Dim path = Process.GetCurrentProcess().MainModule.FileName
                If String.IsNullOrWhiteSpace(path) OrElse Not File.Exists(path) Then Return ""
                Using fs = File.OpenRead(path)
                    Dim hash As Byte()
                    Using md5 As MD5 = MD5.Create()
                        hash = md5.ComputeHash(fs)
                    End Using
                    Dim sb As New StringBuilder()
                    For Each b As Byte In hash
                        sb.Append(b.ToString("x2"))
                    Next
                    Return sb.ToString()
                End Using
            Catch
                Return ""
            End Try
        End Function

        ' AES-128-CBC 加密并返回 Base64（key 与 iv 使用同一 16 字节字符串）。
        Private Shared Function AesEncryptBase64(plain As String, key16 As String) As String
            Using aes As Aes = Aes.Create()
                aes.Mode = CipherMode.CBC
                aes.Padding = PaddingMode.PKCS7
                aes.Key = Encoding.UTF8.GetBytes(key16)
                aes.IV = Encoding.UTF8.GetBytes(key16)
                Dim p As Byte() = Encoding.UTF8.GetBytes(plain)
                Using enc As ICryptoTransform = aes.CreateEncryptor()
                    Return Convert.ToBase64String(enc.TransformFinalBlock(p, 0, p.Length))
                End Using
            End Using
        End Function

        ' AES-128-CBC 解密 Base64 密文并返回 UTF-8 文本。
        Private Shared Function AesDecryptBase64(cipherB64 As String, key16 As String) As String
            Using aes As Aes = Aes.Create()
                aes.Mode = CipherMode.CBC
                aes.Padding = PaddingMode.PKCS7
                aes.Key = Encoding.UTF8.GetBytes(key16)
                aes.IV = Encoding.UTF8.GetBytes(key16)
                Dim c As Byte() = Convert.FromBase64String(cipherB64)
                Using dec As ICryptoTransform = aes.CreateDecryptor()
                    Dim plain = dec.TransformFinalBlock(c, 0, c.Length)
                    Return Encoding.UTF8.GetString(plain)
                End Using
            End Using
        End Function

        ' 使用客户端公钥做 RSA PKCS1 加密，返回 Base64。
        Private Shared Function RsaEncryptPkcs1Base64(message As String, publicKeyB64Der As String) As String
            Using rsa As New RSACryptoServiceProvider()
                ImportPublicKeyDer(rsa, Convert.FromBase64String(publicKeyB64Der))
                Dim enc As Byte() = rsa.Encrypt(Encoding.UTF8.GetBytes(message), False)
                Return Convert.ToBase64String(enc)
            End Using
        End Function

        ' 使用服务器私钥做 RSA PKCS1 解密，返回 UTF-8 文本。
        ' 兼容两类私钥格式：
        ' 1) PKCS#8（常见于 Java/OpenSSL 导出）
        ' 2) PKCS#1（纯 RSA 私钥序列）
        ' 并在解密失败时输出关键信息，便于和官方 SDK 对照排错。
        Private Shared Function RsaDecryptPkcs1Base64(cipherB64 As String, privateKeyB64Der As String) As String
            Try
                Dim privateDer As Byte() = Convert.FromBase64String(privateKeyB64Der)
                Dim rsaParams As RSAParameters = ParsePkcs8PrivateKey(privateDer)
                If rsaParams.Modulus Is Nothing OrElse rsaParams.D Is Nothing Then
                    rsaParams = ParsePkcs1PrivateKey(privateDer)
                End If
                If rsaParams.Modulus Is Nothing OrElse rsaParams.D Is Nothing Then
                    DebugApiLog("异常", "RSA私钥解析失败")
                    Return ""
                End If
                Using rsa As New RSACryptoServiceProvider()
                    rsa.PersistKeyInCsp = False
                    rsa.ImportParameters(rsaParams)
                    Dim cipherBytes As Byte() = Convert.FromBase64String(cipherB64)
                    If rsaParams.Modulus IsNot Nothing AndAlso cipherBytes.Length <> rsaParams.Modulus.Length Then
                        DebugApiLog("异常", "RSA密文长度不匹配，cipher=" & cipherBytes.Length & ", key=" & rsaParams.Modulus.Length)
                    End If
                    Dim decrypted As Byte() = rsa.Decrypt(cipherBytes, False)
                    Return Encoding.UTF8.GetString(decrypted)
                End Using
            Catch ex As Exception
                DebugApiLog("异常", "RSA解密异常: " & ex.GetType().Name & " - " & ex.Message)
                Return ""
            End Try
        End Function

        ' X.509 SubjectPublicKeyInfo (SPKI) DER：SEQUENCE { AlgorithmIdentifier, BIT STRING(RSAPublicKey) }
        Private Shared Sub ImportPublicKeyDer(rsa As RSACryptoServiceProvider, key As Byte())
            Using ms As New MemoryStream(key)
                Using br As New BinaryReader(ms)
                    If br.ReadByte() <> &H30 Then Throw New CryptographicException("Invalid public key")
                    ReadAsnLength(br)
                    If br.ReadByte() <> &H30 Then Throw New CryptographicException("Invalid public key")
                    Dim algIdLen = ReadAsnLength(br)
                    If algIdLen < 0 OrElse algIdLen > key.Length Then Throw New CryptographicException("Invalid public key")
                    br.ReadBytes(algIdLen)
                    If br.ReadByte() <> &H3 Then Throw New CryptographicException("Invalid public key")
                    ReadAsnLength(br)
                    Dim unusedBits = br.ReadByte()
                    If unusedBits <> 0 Then Throw New CryptographicException("Invalid public key")
                    If br.ReadByte() <> &H30 Then Throw New CryptographicException("Invalid public key")
                    ReadAsnLength(br)
                    Dim modulus = ReadAsnInteger(br)
                    Dim exponent = ReadAsnInteger(br)
                    rsa.ImportParameters(New RSAParameters With {.Modulus = modulus, .Exponent = exponent})
                End Using
            End Using
        End Sub

        Private Shared Sub ImportPrivateKeyDer(rsa As RSACryptoServiceProvider, key As Byte())
            Try
                ImportPkcs1PrivateKeyDer(rsa, key)
            Catch
                ImportPkcs8PrivateKeyDer(rsa, key)
            End Try
        End Sub

        Private Shared Sub ImportPkcs1PrivateKeyDer(rsa As RSACryptoServiceProvider, key As Byte())
            Using ms As New MemoryStream(key)
                Using br As New BinaryReader(ms)
                    If br.ReadByte() <> &H30 Then Throw New CryptographicException("Invalid private key")
                    ReadAsnLength(br)
                    If br.ReadByte() <> &H2 Then Throw New CryptographicException("Invalid private key")
                    Dim verLen = ReadAsnLength(br)
                    br.ReadBytes(verLen)
                    Dim p As RSAParameters
                    p.Modulus = ReadAsnInteger(br)
                    p.Exponent = ReadAsnInteger(br)
                    p.D = ReadAsnInteger(br)
                    p.P = ReadAsnInteger(br)
                    p.Q = ReadAsnInteger(br)
                    p.DP = ReadAsnInteger(br)
                    p.DQ = ReadAsnInteger(br)
                    p.InverseQ = ReadAsnInteger(br)
                    rsa.ImportParameters(p)
                End Using
            End Using
        End Sub

        ' PKCS#8: SEQUENCE { version, algorithmIdentifier, OCTET STRING(RSAPrivateKey) }
        Private Shared Sub ImportPkcs8PrivateKeyDer(rsa As RSACryptoServiceProvider, key As Byte())
            Using ms As New MemoryStream(key)
                Using br As New BinaryReader(ms)
                    If br.ReadByte() <> &H30 Then Throw New CryptographicException("Invalid private key")
                    ReadAsnLength(br)
                    If br.ReadByte() <> &H2 Then Throw New CryptographicException("Invalid private key")
                    br.ReadBytes(ReadAsnLength(br))
                    If br.ReadByte() <> &H30 Then Throw New CryptographicException("Invalid private key")
                    br.ReadBytes(ReadAsnLength(br))
                    If br.ReadByte() <> &H4 Then Throw New CryptographicException("Invalid private key")
                    Dim octLen = ReadAsnLength(br)
                    Dim pkcs1 = br.ReadBytes(octLen)
                    If pkcs1 Is Nothing OrElse pkcs1.Length = 0 Then Throw New CryptographicException("Invalid private key")
                    ImportPkcs1PrivateKeyDer(rsa, pkcs1)
                End Using
            End Using
        End Sub

        Private Shared Function ReadAsnInteger(br As BinaryReader) As Byte()
            If br.ReadByte() <> &H2 Then Throw New CryptographicException("Invalid integer")
            Dim len = ReadAsnLength(br)
            Dim v = br.ReadBytes(len)
            If v.Length > 0 AndAlso v(0) = 0 Then
                Dim trimmed(v.Length - 2) As Byte
                Buffer.BlockCopy(v, 1, trimmed, 0, trimmed.Length)
                Return trimmed
            End If
            Return v
        End Function

        Private Shared Function ReadAsnLength(br As BinaryReader) As Integer
            Dim b As Integer = br.ReadByte()
            If (b And &H80) = 0 Then Return b
            Dim bytesCount = b And &H7F
            Dim length As Integer = 0
            For i As Integer = 0 To bytesCount - 1
                length = (length << 8) Or br.ReadByte()
            Next
            Return length
        End Function

        ' 解析 PKCS#8 外层结构，再把内层 OCTET STRING 交给 PKCS#1 解析器。
        Private Shared Function ParsePkcs8PrivateKey(der As Byte()) As RSAParameters
            Dim idx As Integer = 0
            If Not TryReadAsn1Tag(der, idx, &H30) Then Return New RSAParameters()
            Dim topLen As Integer = 0
            If Not TryReadAsn1Length(der, idx, topLen) Then Return New RSAParameters()
            Dim version As Byte() = Nothing
            If Not TryReadAsn1Integer(der, idx, version) Then Return New RSAParameters()
            If version Is Nothing Then Return New RSAParameters()
            If Not TryReadAsn1Tag(der, idx, &H30) Then Return New RSAParameters()
            Dim algLen As Integer = 0
            If Not TryReadAsn1Length(der, idx, algLen) Then Return New RSAParameters()
            If idx + algLen > der.Length Then Return New RSAParameters()
            idx += algLen
            If Not TryReadAsn1Tag(der, idx, &H4) Then Return New RSAParameters()
            Dim keyOctLen As Integer = 0
            If Not TryReadAsn1Length(der, idx, keyOctLen) Then Return New RSAParameters()
            If keyOctLen <= 0 OrElse idx + keyOctLen > der.Length Then Return New RSAParameters()
            Dim keyOct(keyOctLen - 1) As Byte
            Buffer.BlockCopy(der, idx, keyOct, 0, keyOctLen)
            Return ParsePkcs1PrivateKey(keyOct)
        End Function

        ' 解析 PKCS#1 RSAParameters，并做左侧补零对齐，避免 ImportParameters 长度不匹配。
        Private Shared Function ParsePkcs1PrivateKey(der As Byte()) As RSAParameters
            Dim idx As Integer = 0
            If Not TryReadAsn1Tag(der, idx, &H30) Then Return New RSAParameters()
            Dim topLen As Integer = 0
            If Not TryReadAsn1Length(der, idx, topLen) Then Return New RSAParameters()
            Dim version As Byte() = Nothing
            If Not TryReadAsn1Integer(der, idx, version) Then Return New RSAParameters()
            If version Is Nothing Then Return New RSAParameters()
            Dim modulus As Byte() = Nothing
            Dim exponent As Byte() = Nothing
            Dim d As Byte() = Nothing
            Dim p As Byte() = Nothing
            Dim q As Byte() = Nothing
            Dim dp As Byte() = Nothing
            Dim dq As Byte() = Nothing
            Dim inverseQ As Byte() = Nothing
            If Not TryReadAsn1Integer(der, idx, modulus) Then Return New RSAParameters()
            If Not TryReadAsn1Integer(der, idx, exponent) Then Return New RSAParameters()
            If Not TryReadAsn1Integer(der, idx, d) Then Return New RSAParameters()
            If Not TryReadAsn1Integer(der, idx, p) Then Return New RSAParameters()
            If Not TryReadAsn1Integer(der, idx, q) Then Return New RSAParameters()
            If Not TryReadAsn1Integer(der, idx, dp) Then Return New RSAParameters()
            If Not TryReadAsn1Integer(der, idx, dq) Then Return New RSAParameters()
            If Not TryReadAsn1Integer(der, idx, inverseQ) Then Return New RSAParameters()
            Dim keyLen As Integer = If(modulus Is Nothing, 0, modulus.Length)
            If keyLen <= 0 Then Return New RSAParameters()
            Dim primeLen As Integer = keyLen \ 2
            If keyLen Mod 2 <> 0 Then primeLen += 1
            Return New RSAParameters With {
                .Modulus = LeftPadToLength(modulus, keyLen),
                .Exponent = exponent,
                .D = LeftPadToLength(d, keyLen),
                .P = LeftPadToLength(p, primeLen),
                .Q = LeftPadToLength(q, primeLen),
                .DP = LeftPadToLength(dp, primeLen),
                .DQ = LeftPadToLength(dq, primeLen),
                .InverseQ = LeftPadToLength(inverseQ, primeLen)
            }
        End Function

        ' ASN.1 整数字段可能带前导 00 或长度偏短，这里统一补齐/裁剪到目标长度。
        Private Shared Function LeftPadToLength(src As Byte(), targetLen As Integer) As Byte()
            If src Is Nothing Then Return Nothing
            If targetLen <= 0 Then Return src
            If src.Length = targetLen Then Return src
            If src.Length > targetLen Then
                Dim cut(targetLen - 1) As Byte
                Buffer.BlockCopy(src, src.Length - targetLen, cut, 0, targetLen)
                Return cut
            End If
            Dim padded(targetLen - 1) As Byte
            Buffer.BlockCopy(src, 0, padded, targetLen - src.Length, src.Length)
            Return padded
        End Function

        Private Shared Function TryReadAsn1Tag(data As Byte(), ByRef idx As Integer, expectedTag As Byte) As Boolean
            If data Is Nothing OrElse idx < 0 OrElse idx >= data.Length Then Return False
            If data(idx) <> expectedTag Then Return False
            idx += 1
            Return True
        End Function

        ' 通用 ASN.1 长度读取（短格式/长格式），同时做边界保护。
        Private Shared Function TryReadAsn1Length(data As Byte(), ByRef idx As Integer, ByRef length As Integer) As Boolean
            length = 0
            If data Is Nothing OrElse idx < 0 OrElse idx >= data.Length Then Return False
            Dim first As Integer = data(idx)
            idx += 1
            If (first And &H80) = 0 Then
                length = first
                Return idx + length <= data.Length
            End If
            Dim byteCount As Integer = first And &H7F
            If byteCount <= 0 OrElse byteCount > 4 Then Return False
            If idx + byteCount > data.Length Then Return False
            For i As Integer = 0 To byteCount - 1
                length = (length << 8) Or data(idx)
                idx += 1
            Next
            Return idx + length <= data.Length
        End Function

        Private Shared Function TryReadAsn1Integer(data As Byte(), ByRef idx As Integer, ByRef value As Byte()) As Boolean
            value = Nothing
            If Not TryReadAsn1Tag(data, idx, &H2) Then Return False
            Dim length As Integer = 0
            If Not TryReadAsn1Length(data, idx, length) Then Return False
            If length < 0 OrElse idx + length > data.Length Then Return False
            Dim start As Integer = idx
            idx += length
            If length = 0 Then
                value = New Byte() {}
                Return True
            End If
            While length > 1 AndAlso data(start) = 0
                start += 1
                length -= 1
            End While
            value = New Byte(length - 1) {}
            Buffer.BlockCopy(data, start, value, 0, length)
            Return True
        End Function

        ' 协议核心调用：
        ' 1) 组装参数 -> 2) AES 加密 -> 3) RSA 签名段加密 -> 4) POST
        ' 5) 解析返回并解密 -> 6) 抽取 code/data/SeSsL。
        ' BSPHP 核心请求管线：
        ' 明文参数 -> AES 加密 -> RSA 包装签名段 -> HTTP POST -> RSA/AES 解包 -> XML 解析。
        Public Async Function SendAsync(api As String, Optional extra As Dictionary(Of String, String) = Nothing) As Task(Of ApiResult)
            Try
                Dim now = DateTime.Now
                Dim appSafe = Md5Hex(now.ToString("yyyy-MM-dd HH:mm:ss"))
                Dim p As New Dictionary(Of String, String)(StringComparer.Ordinal) From {
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

                Dim dataStr = Global.BSPHP.Protocol.BsphpProtocolEncoding.BuildPlainQueryString(p)
                DebugApiLog("加密前", "api=" & api & " " & dataStr)
                Dim aesKey = Md5Hex(_cfg.ServerPrivateKey & appSafe).Substring(0, 16)
                Dim encB64 = AesEncryptBase64(dataStr, aesKey)
                Dim sig = "0|AES-128-CBC|" & aesKey & "|" & Md5Hex(encB64) & "|xml"
                Dim rsaB64 = RsaEncryptPkcs1Base64(sig, _cfg.ClientPublicKey)
                Dim payload = Global.BSPHP.Protocol.BsphpProtocolEncoding.UrlEncodeCxx(encB64 & "|" & rsaB64)
                DebugApiLog("加密后", "api=" & api & " " & payload)
                Dim body = New StringContent("parameter=" & payload, Encoding.UTF8, "application/x-www-form-urlencoded")
                Dim resp = Await _http.PostAsync(_cfg.Url, body)
                Dim rawBody = Await resp.Content.ReadAsStringAsync()
                Dim u = Global.BSPHP.Protocol.BsphpProtocolEncoding.UrlDecodeCxx(rawBody)
                Dim encPart As String = ""
                Dim rsaPart As String = ""
                If Not Global.BSPHP.Protocol.BsphpProtocolEncoding.TrySplitEncryptedResponse(u, encPart, rsaPart) Then Return New ApiResult With {.DataText = "响应格式错误"}
                DebugApiLog("解密前", "api=" & api & " " & encPart)

                Dim jsonPlain As String = ""
                Try
                    Dim sigPlain = RsaDecryptPkcs1Base64(rsaPart.Trim(), _cfg.ServerPrivateKey)
                    Dim sigSeg = sigPlain.Split("|"c)
                    If sigSeg.Length < 3 OrElse sigSeg(2).Length < 16 Then Return New ApiResult With {.DataText = "响应签名格式错误"}
                    Dim respKey = sigSeg(2).Substring(0, 16)
                    jsonPlain = AesDecryptBase64(encPart.Trim(), respKey)
                    DebugApiLog("解密后", "api=" & api & " " & jsonPlain)
                Catch ex As CryptographicException
                    Return New ApiResult With {.DataText = "加解密失败: " & ex.Message}
                End Try

                Dim r As New ApiResult()
                Try
                    Dim doc = XDocument.Parse(jsonPlain)
                    Dim respNode = doc.Root
                    If respNode IsNot Nothing AndAlso String.Equals(respNode.Name.LocalName, "response", StringComparison.OrdinalIgnoreCase) Then
                        Dim codeText = GetXmlChildValue(respNode, "code")
                        If codeText <> "" Then
                            Dim codeInt As Integer
                            If Integer.TryParse(codeText, codeInt) Then r.Code = codeInt
                        End If
                        r.DataText = GetXmlChildValue(respNode, "data")
                        Dim sslText = GetXmlChildValue(respNode, "SeSsL")
                        If sslText <> "" AndAlso (r.Code = 1011 OrElse r.Code = 9908 OrElse r.Code = 1081) Then
                            Sessl = sslText
                        End If
                        Return r
                    End If
                Catch
                End Try

                r.DataText = jsonPlain
                Return r
            Catch ex As Exception
                Return New ApiResult With {.DataText = "请求失败(" & api & "): " & ex.Message}
            End Try
        End Function

        Private Shared Function GetXmlChildValue(parent As XElement, childName As String) As String
            For Each e As XElement In parent.Elements()
                If String.Equals(e.Name.LocalName, childName, StringComparison.OrdinalIgnoreCase) Then
                    Return If(e.Value, "").Trim()
                End If
            Next
            Return ""
        End Function

        Private Shared Sub DebugApiLog(stage As String, content As String)
            Dim safeContent As String = If(content, "")
            Debug.WriteLine("[BSPHP][" & stage & "] " & TruncateForLog(safeContent, 1500))
        End Sub

        Private Shared Function TruncateForLog(text As String, maxLen As Integer) As String
            If text Is Nothing Then Return ""
            If text.Length <= maxLen Then Return text
            Return text.Substring(0, maxLen) & "...(truncated)"
        End Function

        ' 启动初始化流程：先连通测试 internet.in，再获取会话 BSphpSeSsL.in。
        Public Async Function BootstrapAsync() As Task(Of String)
            Dim c = Await SendAsync("internet.in")
            If c.DataText <> "1" Then Return "连接失败: " & c.DataText
            Dim s = Await SendAsync("BSphpSeSsL.in")
            Sessl = s.DataText
            Return If(String.IsNullOrWhiteSpace(Sessl), "获取会话失败: " & s.DataText, "")
        End Function

        ' ===== 与 mac 登录版 API 对齐的公共封装 =====
        ' 对应 internet.in，返回 True 表示网络与应用配置可用。
        Public Async Function ConnectAsync() As Task(Of Boolean)
            Dim r = Await SendAsync("internet.in")
            Return r.DataText = "1"
        End Function

        ' 取会话令牌，后续所有接口都依赖该值。
        Public Async Function GetSeSsLAsync() As Task(Of Boolean)
            Dim r = Await SendAsync("BSphpSeSsL.in")
            Sessl = If(r.DataText, "").Trim()
            Return Sessl <> ""
        End Function

        ' 注销后主动清空本地会话并重新取 SeSsL，保持客户端状态一致。
        Public Async Function LogoutAsync() As Task(Of ApiResult)
            Dim r = Await SendAsync("cancellation.lg")
            Sessl = ""
            Await GetSeSsLAsync()
            Return r
        End Function

        Public Async Function GetNoticeAsync() As Task(Of ApiResult)
            Return Await SendAsync("gg.in")
        End Function

        Public Async Function GetVersionAsync() As Task(Of ApiResult)
            Return Await SendAsync("v.in")
        End Function

        Public Async Function GetSoftInfoAsync() As Task(Of ApiResult)
            Return Await SendAsync("miao.in")
        End Function

        Public Async Function GetServerDateAsync(Optional dateFormatM As String = "") As Task(Of ApiResult)
            Dim p As Dictionary(Of String, String) = Nothing
            If dateFormatM <> "" Then p = New Dictionary(Of String, String) From {{"m", dateFormatM}}
            Return Await SendAsync("date.in", p)
        End Function

        Public Async Function GetPresetURLAsync() As Task(Of ApiResult)
            Return Await SendAsync("url.in")
        End Function

        Public Async Function GetWebURLAsync() As Task(Of ApiResult)
            Return Await SendAsync("weburl.in")
        End Function

        Public Async Function GetGlobalInfoAsync(Optional info As String = "") As Task(Of ApiResult)
            Dim p As Dictionary(Of String, String) = Nothing
            If info <> "" Then p = New Dictionary(Of String, String) From {{"info", info}}
            Return Await SendAsync("globalinfo.in", p)
        End Function

        Public Async Function GetAppCustomAsync(info As String) As Task(Of ApiResult)
            Return Await SendAsync("appcustom.in", New Dictionary(Of String, String) From {{"info", info}})
        End Function

        Public Async Function GetCodeEnabledAsync(Optional codeType As String = "") As Task(Of ApiResult)
            Dim p As Dictionary(Of String, String) = Nothing
            If codeType <> "" Then p = New Dictionary(Of String, String) From {{"type", codeType}}
            Return Await SendAsync("getsetimag.in", p)
        End Function

        Public Async Function GetCodeEnabledAsync(types As IEnumerable(Of String)) As Task(Of ApiResult)
            Dim joined = String.Join("|", types)
            Return Await GetCodeEnabledAsync(joined)
        End Function

        Public Async Function GetLogicAAsync() As Task(Of ApiResult)
            Return Await SendAsync("logica.in")
        End Function

        Public Async Function GetLogicBAsync() As Task(Of ApiResult)
            Return Await SendAsync("logicb.in")
        End Function

        Public Async Function GetLogicInfoAAsync() As Task(Of ApiResult)
            Return Await SendAsync("logicinfoa.in")
        End Function

        Public Async Function GetLogicInfoBAsync() As Task(Of ApiResult)
            Return Await SendAsync("logicinfob.in")
        End Function

        Public Async Function GetEndTimeAsync() As Task(Of ApiResult)
            Return Await SendAsync("vipdate.lg")
        End Function

        ' info 为空取默认字段，传逗号分隔字段时只返回指定项。
        Public Async Function GetUserInfoAsync(Optional info As String = "") As Task(Of ApiResult)
            Dim p As Dictionary(Of String, String) = Nothing
            If info <> "" Then p = New Dictionary(Of String, String) From {{"info", info}}
            Return Await SendAsync("getuserinfo.lg", p)
        End Function

        Public Async Function GetUserInfoAsync(fields As IEnumerable(Of String)) As Task(Of ApiResult)
            Return Await GetUserInfoAsync(String.Join(",", fields))
        End Function

        Public Async Function GetUserKeyAsync() As Task(Of ApiResult)
            Return Await SendAsync("userkey.lg")
        End Function

        Public Async Function HeartbeatAsync() As Task(Of ApiResult)
            Return Await SendAsync("timeout.lg")
        End Function

        ' 获取验证码图片 URL（依赖已拿到的 Sessl）。
        Public Function GetCodeImageUrl() As String
            Return If(String.IsNullOrWhiteSpace(Sessl), "", _cfg.CodeUrlPrefix & Sessl)
        End Function

        ' 账号密码登录（login.lg）。
        ' 账号密码登录；图形码为空时按官方行为传 "null"。
        Public Async Function LoginAsync(user As String, pwd As String, coode As String) As Task(Of ApiResult)
            Dim imageCode As String = If(String.IsNullOrWhiteSpace(coode), "null", coode)
            Return Await SendAsync("login.lg", New Dictionary(Of String, String) From {
                {"user", user},
                {"pwd", pwd},
                {"coode", imageCode},
                {"key", MachineCode()},
                {"maxoror", MachineCode()}
            })
        End Function

        ' 短信验证码登录（login_sms.lg）。
        Public Async Function LoginSmsAsync(mobile As String, area As String, smsCode As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("login_sms.lg", New Dictionary(Of String, String) From {
                {"mobile", mobile}, {"area", area}, {"sms_code", smsCode},
                {"key", MachineCode()}, {"maxoror", MachineCode()}, {"coode", coode}
            })
        End Function

        ' 邮箱验证码登录（login_email.lg）。
        Public Async Function LoginEmailAsync(email As String, emailCode As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("login_email.lg", New Dictionary(Of String, String) From {
                {"email", email}, {"email_code", emailCode},
                {"key", MachineCode()}, {"maxoror", MachineCode()}, {"coode", coode}
            })
        End Function

        ' mac 命名别名，便于一一对应迁移。
        Public Async Function SendSmsCodeAsync(scene As String, mobile As String, area As String, coode As String) As Task(Of ApiResult)
            Return Await SendSmsAsync(scene, mobile, area, coode)
        End Function

        Public Async Function SendEmailCodeAsync(scene As String, email As String, coode As String) As Task(Of ApiResult)
            Return Await SendEmailAsync(scene, email, coode)
        End Function

        Public Async Function RegisterSmsByMacNameAsync(user As String, mobile As String, area As String, smsCode As String, pwd As String, pwdb As String, key As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("register_sms.lg", New Dictionary(Of String, String) From {
                {"user", user}, {"mobile", mobile}, {"area", area}, {"sms_code", smsCode},
                {"pwd", pwd}, {"pwdb", pwdb}, {"key", key}, {"coode", coode}
            })
        End Function

        Public Async Function LoginSmsByMacNameAsync(mobile As String, area As String, smsCode As String, key As String, maxoror As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("login_sms.lg", New Dictionary(Of String, String) From {
                {"mobile", mobile}, {"area", area}, {"sms_code", smsCode},
                {"key", key}, {"maxoror", maxoror}, {"coode", coode}
            })
        End Function

        Public Async Function RegisterEmailByMacNameAsync(user As String, email As String, emailCode As String, pwd As String, pwdb As String, key As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("register_email.lg", New Dictionary(Of String, String) From {
                {"user", user}, {"email", email}, {"email_code", emailCode},
                {"pwd", pwd}, {"pwdb", pwdb}, {"key", key}, {"coode", coode}
            })
        End Function

        Public Async Function LoginEmailByMacNameAsync(email As String, emailCode As String, key As String, maxoror As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("login_email.lg", New Dictionary(Of String, String) From {
                {"email", email}, {"email_code", emailCode},
                {"key", key}, {"maxoror", maxoror}, {"coode", coode}
            })
        End Function

        ' 发送短信验证码（send_sms.lg），scene 示例：login/register/resetpwd。
        Public Async Function SendSmsAsync(scene As String, mobile As String, area As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("send_sms.lg", New Dictionary(Of String, String) From {
                {"scene", scene}, {"mobile", mobile}, {"area", area}, {"coode", coode}
            })
        End Function

        ' 发送邮箱验证码（send_email.lg），scene 示例：login/register/resetpwd。
        Public Async Function SendEmailAsync(scene As String, email As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("send_email.lg", New Dictionary(Of String, String) From {
                {"scene", scene}, {"email", email}, {"coode", coode}
            })
        End Function

        ' 账号注册（registration.lg）。
        Public Async Function RegisterAsync(user As String, pwd As String, pwdb As String, coode As String, mobile As String, question As String, answer As String, qq As String, mail As String, extensionCode As String) As Task(Of ApiResult)
            Return Await SendAsync("registration.lg", New Dictionary(Of String, String) From {
                {"user", user}, {"pwd", pwd}, {"pwdb", pwdb}, {"coode", coode},
                {"mobile", mobile}, {"mibao_wenti", question}, {"mibao_daan", answer},
                {"qq", qq}, {"mail", mail}, {"extension", extensionCode}
            })
        End Function

        ' 短信注册（register_sms.lg）。
        Public Async Function RegisterSmsAsync(user As String, mobile As String, area As String, smsCode As String, pwd As String, pwdb As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("register_sms.lg", New Dictionary(Of String, String) From {
                {"user", user}, {"mobile", mobile}, {"area", area}, {"sms_code", smsCode},
                {"pwd", pwd}, {"pwdb", pwdb}, {"key", MachineCode()}, {"coode", coode}
            })
        End Function

        ' 邮箱注册（register_email.lg）。
        Public Async Function RegisterEmailAsync(user As String, email As String, emailCode As String, pwd As String, pwdb As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("register_email.lg", New Dictionary(Of String, String) From {
                {"user", user}, {"email", email}, {"email_code", emailCode},
                {"pwd", pwd}, {"pwdb", pwdb}, {"key", MachineCode()}, {"coode", coode}
            })
        End Function

        ' 密保找回密码（backto.lg）。
        Public Async Function ResetPwdAsync(user As String, pwd As String, pwdb As String, question As String, answer As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("backto.lg", New Dictionary(Of String, String) From {
                {"user", user}, {"pwd", pwd}, {"pwdb", pwdb}, {"wenti", question}, {"daan", answer}, {"coode", coode}
            })
        End Function

        ' 短信找回密码（resetpwd_sms.lg）。
        Public Async Function ResetSmsPwdAsync(mobile As String, area As String, smsCode As String, pwd As String, pwdb As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("resetpwd_sms.lg", New Dictionary(Of String, String) From {
                {"mobile", mobile}, {"area", area}, {"sms_code", smsCode}, {"pwd", pwd}, {"pwdb", pwdb}, {"coode", coode}
            })
        End Function

        ' 邮箱找回密码（resetpwd_email.lg）。
        Public Async Function ResetEmailPwdAsync(email As String, emailCode As String, pwd As String, pwdb As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("resetpwd_email.lg", New Dictionary(Of String, String) From {
                {"email", email}, {"email_code", emailCode}, {"pwd", pwd}, {"pwdb", pwdb}, {"coode", coode}
            })
        End Function

        ' 解绑（jiekey.lg）。
        Public Async Function UnbindAsync(user As String, pwd As String) As Task(Of ApiResult)
            Return Await SendAsync("jiekey.lg", New Dictionary(Of String, String) From {{"user", user}, {"pwd", pwd}})
        End Function

        ' 账号充值（chong.lg），verifyPwd 对应 userset: 1=校验用户密码。
        Public Async Function RechargeAsync(user As String, userpwd As String, verifyPwd As Boolean, ka As String, pwd As String) As Task(Of ApiResult)
            Return Await SendAsync("chong.lg", New Dictionary(Of String, String) From {
                {"user", user}, {"userpwd", userpwd}, {"userset", If(verifyPwd, "1", "0")}, {"ka", ka}, {"pwd", pwd}
            })
        End Function

        ' 修改密码（password.lg）。
        Public Async Function ChangePwdAsync(user As String, oldPwd As String, newPwd As String, newPwd2 As String, img As String) As Task(Of ApiResult)
            Return Await SendAsync("password.lg", New Dictionary(Of String, String) From {
                {"user", user}, {"pwd", oldPwd}, {"pwda", newPwd}, {"pwdb", newPwd2}, {"img", img}
            })
        End Function

        ' 意见反馈（liuyan.in）。
        Public Async Function FeedbackAsync(user As String, pwd As String, title As String, qq As String, leix As String, text As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("liuyan.in", New Dictionary(Of String, String) From {
                {"user", user}, {"pwd", pwd}, {"table", title}, {"qq", qq}, {"leix", leix}, {"txt", text}, {"coode", coode}
            })
        End Function
    End Class
End Namespace
