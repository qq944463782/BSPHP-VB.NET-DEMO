Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Net.Http
Imports System.Diagnostics
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading.Tasks
Imports System.Xml.Linq
Imports System.Management

Namespace BSPHP.LoginDemoVB
    ' Card demo local protocol client (SDK mode, full encrypted request/response).
    Public Class BSPHPConfig
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

        Private Shared Function RsaEncryptPkcs1Base64(message As String, publicKeyB64Der As String) As String
            Using rsa As New RSACryptoServiceProvider()
                ImportPublicKeyDer(rsa, Convert.FromBase64String(publicKeyB64Der))
                Dim enc As Byte() = rsa.Encrypt(Encoding.UTF8.GetBytes(message), False)
                Return Convert.ToBase64String(enc)
            End Using
        End Function

        ' 与登录版保持同一私钥兼容策略，确保卡密模式与账号模式行为一致。
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

        ' PKCS#8 -> 取内层 PKCS#1。
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

        ' PKCS#1 解析并补齐参数长度，避免 RSA 导入时出现异常。
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

        Public Async Function BootstrapAsync() As Task(Of String)
            Dim c = Await SendAsync("internet.in")
            If c.DataText <> "1" Then Return "连接失败: " & c.DataText
            Dim s = Await SendAsync("BSphpSeSsL.in")
            Sessl = s.DataText
            Return If(String.IsNullOrWhiteSpace(Sessl), "获取会话失败: " & s.DataText, "")
        End Function

        ' ===== 与 mac 卡密版 API 对齐的封装 =====
        Public Async Function ConnectAsync() As Task(Of Boolean)
            Dim r = Await SendAsync("internet.in")
            Return r.DataText = "1"
        End Function

        Public Async Function GetSeSsLAsync() As Task(Of Boolean)
            Dim r = Await SendAsync("BSphpSeSsL.in")
            Sessl = If(r.DataText, "").Trim()
            Return Sessl <> ""
        End Function

        ' 注销卡密会话并重取 SeSsL，便于后续重新登录。
        Public Async Function LogoutAsync() As Task(Of ApiResult)
            Dim r = Await SendAsync("cancellation.ic")
            Sessl = ""
            Await GetSeSsLAsync()
            Return r
        End Function

        Private Async Function ApiResultAsync(api As String, Optional extra As Dictionary(Of String, String) = Nothing) As Task(Of ApiResult)
            Return Await SendAsync(api, extra)
        End Function

        Public Async Function AddCardFeaturesAsync(carid As String, key As String, maxoror As String) As Task(Of ApiResult)
            Return Await ApiResultAsync("AddCardFeatures.key.ic", New Dictionary(Of String, String) From {{"carid", carid}, {"key", key}, {"maxoror", maxoror}})
        End Function

        Public Async Function CallRemoteAsync(datas As String) As Task(Of ApiResult)
            Return Await ApiResultAsync("CallRemote.in", New Dictionary(Of String, String) From {{"datas", datas}})
        End Function

        Public Async Function GetMyDataAsync(keys As String) As Task(Of ApiResult)
            Return Await ApiResultAsync("GetMyData.in", New Dictionary(Of String, String) From {{"keys", keys}})
        End Function

        Public Async Function SetAppRemarksAsync(icid As String, icpwd As String, remarks As String) As Task(Of ApiResult)
            Return Await ApiResultAsync("SetAppRemarks.ic", New Dictionary(Of String, String) From {{"icid", icid}, {"icpwd", icpwd}, {"remarks", remarks}})
        End Function

        Public Async Function SetMyDataAsync(keys As String, datas As String) As Task(Of ApiResult)
            Return Await ApiResultAsync("SetMysData.in", New Dictionary(Of String, String) From {{"keys", keys}, {"datas", datas}})
        End Function

        Public Async Function AppBadPushAsync(table As String) As Task(Of ApiResult)
            Return Await ApiResultAsync("appbadpush.in", New Dictionary(Of String, String) From {{"table", table}})
        End Function

        Public Async Function GetAppCustomAsync(info As String, Optional getTypeValue As String = "", Optional user As String = "", Optional pwd As String = "", Optional icid As String = "", Optional icpwd As String = "") As Task(Of ApiResult)
            Dim p As New Dictionary(Of String, String) From {{"info", info}}
            If getTypeValue <> "" Then p("get_type") = getTypeValue
            If user <> "" Then p("user") = user
            If pwd <> "" Then p("pwd") = pwd
            If icid <> "" Then p("icid") = icid
            If icpwd <> "" Then p("icpwd") = icpwd
            Return Await ApiResultAsync("appcustom.in", p)
        End Function

        Public Async Function RechargeCardAsync(icid As String, ka As String, pwd As String) As Task(Of ApiResult)
            Return Await ApiResultAsync("chong.ic", New Dictionary(Of String, String) From {{"icid", icid}, {"ka", ka}, {"pwd", pwd}})
        End Function

        Public Async Function GetServerDateAsync(Optional dateFormatM As String = "") As Task(Of ApiResult)
            Dim p As Dictionary(Of String, String) = Nothing
            If dateFormatM <> "" Then p = New Dictionary(Of String, String) From {{"m", dateFormatM}}
            Return Await ApiResultAsync("date.in", p)
        End Function

        Public Async Function GetDataAsync(key As String) As Task(Of ApiResult)
            Return Await ApiResultAsync("getdata.ic", New Dictionary(Of String, String) From {{"key", key}})
        End Function

        Public Async Function GetDateICAsync() As Task(Of ApiResult)
            Return Await ApiResultAsync("getdate.ic")
        End Function

        Public Async Function GetCardInfoAsync(icCarid As String, icPwd As String, info As String, Optional typeValue As String = "") As Task(Of ApiResult)
            Dim p As New Dictionary(Of String, String) From {{"ic_carid", icCarid}, {"ic_pwd", icPwd}, {"info", info}}
            If typeValue <> "" Then p("type") = typeValue
            Return Await ApiResultAsync("getinfo.ic", p)
        End Function

        Public Async Function GetLoginInfoAsync() As Task(Of ApiResult)
            Return Await ApiResultAsync("getlkinfo.ic")
        End Function

        Public Async Function GetGlobalInfoAsync(Optional info As String = "") As Task(Of ApiResult)
            Dim p As Dictionary(Of String, String) = Nothing
            If info <> "" Then p = New Dictionary(Of String, String) From {{"info", info}}
            Return Await ApiResultAsync("globalinfo.in", p)
        End Function

        Public Async Function GetCaptchaImageAsync() As Task(Of ApiResult)
            Return Await ApiResultAsync("imga.in")
        End Function

        Public Async Function PushAddMoneyAsync(user As String, ka As String) As Task(Of ApiResult)
            Return Await ApiResultAsync("pushaddmoney.in", New Dictionary(Of String, String) From {{"user", user}, {"ka", ka}})
        End Function

        Public Async Function PushLogAsync(user As String, log As String) As Task(Of ApiResult)
            Return Await ApiResultAsync("pushlog.in", New Dictionary(Of String, String) From {{"user", user}, {"log", log}})
        End Function

        Public Async Function RemoteCancellationAsync(icid As String, icpwd As String, typeValue As String, Optional biaoji As String = "") As Task(Of ApiResult)
            Dim p As New Dictionary(Of String, String) From {{"icid", icid}, {"icpwd", icpwd}, {"type", typeValue}}
            If biaoji <> "" Then p("biaoji") = biaoji
            Return Await ApiResultAsync("remotecancellation.ic", p)
        End Function

        Public Async Function UnbindCardAsync(icid As String, icpwd As String) As Task(Of ApiResult)
            Return Await ApiResultAsync("setcarnot.ic", New Dictionary(Of String, String) From {{"icid", icid}, {"icpwd", icpwd}})
        End Function

        Public Async Function BindCardAsync(key As String, icid As String, icpwd As String) As Task(Of ApiResult)
            Return Await ApiResultAsync("setcaron.ic", New Dictionary(Of String, String) From {{"key", key}, {"icid", icid}, {"icpwd", icpwd}})
        End Function

        Public Async Function QueryCardAsync(cardid As String) As Task(Of ApiResult)
            Return Await ApiResultAsync("socard.in", New Dictionary(Of String, String) From {{"cardid", cardid}})
        End Function

        Public Async Function GetNoticeAsync() As Task(Of ApiResult)
            Return Await ApiResultAsync("gg.in")
        End Function

        Public Async Function GetVersionAsync() As Task(Of ApiResult)
            Return Await ApiResultAsync("v.in")
        End Function

        Public Async Function GetSoftInfoAsync() As Task(Of ApiResult)
            Return Await ApiResultAsync("miao.in")
        End Function

        Public Async Function GetPresetURLAsync() As Task(Of ApiResult)
            Return Await ApiResultAsync("url.in")
        End Function

        Public Async Function GetWebURLAsync() As Task(Of ApiResult)
            Return Await ApiResultAsync("weburl.in")
        End Function

        Public Async Function GetCodeEnabledAsync(Optional codeType As String = "") As Task(Of ApiResult)
            Dim p As Dictionary(Of String, String) = Nothing
            If codeType <> "" Then p = New Dictionary(Of String, String) From {{"type", codeType}}
            Return Await ApiResultAsync("getsetimag.in", p)
        End Function

        Public Async Function GetCodeEnabledAsync(types As IEnumerable(Of String)) As Task(Of ApiResult)
            Return Await GetCodeEnabledAsync(String.Join("|", types))
        End Function

        Public Async Function GetLogicAAsync() As Task(Of ApiResult)
            Return Await ApiResultAsync("logica.in")
        End Function

        Public Async Function GetLogicBAsync() As Task(Of ApiResult)
            Return Await ApiResultAsync("logicb.in")
        End Function

        Public Async Function GetLogicInfoAAsync() As Task(Of ApiResult)
            Return Await ApiResultAsync("logicinfoa.in")
        End Function

        Public Async Function GetLogicInfoBAsync() As Task(Of ApiResult)
            Return Await ApiResultAsync("logicinfob.in")
        End Function

        Public Async Function HeartbeatAsync() As Task(Of ApiResult)
            Return Await ApiResultAsync("timeout.ic")
        End Function

        ' 卡密登录：未显式传 key/maxoror 时默认使用机器码。
        Public Async Function LoginICAsync(icid As String, Optional icpwd As String = "", Optional key As String = "", Optional maxoror As String = "") As Task(Of ApiResult)
            Dim k = If(key = "", MachineCode(), key)
            Dim m = If(maxoror = "", k, maxoror)
            Return Await ApiResultAsync("login.ic", New Dictionary(Of String, String) From {
                {"icid", icid}, {"icpwd", icpwd}, {"key", k}, {"maxoror", m}
            })
        End Function

        Public Async Function FeedbackAsync(table As String, leix As String, qq As String, txt As String, Optional img As String = "", Optional user As String = "", Optional pwd As String = "") As Task(Of ApiResult)
            Dim p As New Dictionary(Of String, String) From {{"table", table}, {"leix", leix}, {"qq", qq}, {"txt", txt}}
            If img <> "" Then p("img") = img
            If user <> "" Then p("user") = user
            If pwd <> "" Then p("pwd") = pwd
            Return Await ApiResultAsync("liuyan.in", p)
        End Function

        Public Function GetCodeImageUrl() As String
            If String.IsNullOrWhiteSpace(Sessl) Then Return ""
            Return _cfg.CodeUrlPrefix & Sessl
        End Function

        ' 卡密版统一请求入口，协议流程与登录版一致。
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
                DebugApiLog("分隔", "------------------------------------------")
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
    End Class
End Namespace
