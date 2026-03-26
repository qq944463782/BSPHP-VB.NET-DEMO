Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Net.Http
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json
Imports System.Threading.Tasks

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
            Return Environment.MachineName & "_" & Environment.UserName
        End Function

        ' 计算 MD5 十六进制字符串（协议里用于 appsafecode 与摘要签名）。
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

        ' AES-128-CBC 加密并返回 Base64（key 与 iv 使用同一 16 字节字符串）。
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

        ' AES-128-CBC 解密 Base64 密文并返回 UTF-8 文本。
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

        ' 使用客户端公钥做 RSA PKCS1 加密，返回 Base64。
        Private Shared Function RsaEncryptPkcs1Base64(message As String, publicKeyB64Der As String) As String
            Using rsa = RSA.Create()
                rsa.ImportSubjectPublicKeyInfo(Convert.FromBase64String(publicKeyB64Der), Nothing)
                Dim enc = rsa.Encrypt(Encoding.UTF8.GetBytes(message), RSAEncryptionPadding.Pkcs1)
                Return Convert.ToBase64String(enc)
            End Using
        End Function

        ' 使用服务器私钥做 RSA PKCS1 解密，返回 UTF-8 文本。
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

        ' 协议核心调用：
        ' 1) 组装参数 -> 2) AES 加密 -> 3) RSA 签名段加密 -> 4) POST
        ' 5) 解析返回并解密 -> 6) 抽取 code/data/SeSsL。
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
                If res.TryGetProperty("SeSsL", ssl) AndAlso (r.Code = 1011 OrElse r.Code = 9908) Then
                    Sessl = ssl.GetString()
                End If
                Return r
            End Using
        End Function

        ' 启动初始化流程：先连通测试 internet.in，再获取会话 BSphpSeSsL.in。
        Public Async Function BootstrapAsync() As Task(Of String)
            Dim c = Await SendAsync("internet.in")
            If c.DataText <> "1" Then Return "连接失败"
            Dim s = Await SendAsync("BSphpSeSsL.in")
            Sessl = s.DataText
            Return If(String.IsNullOrWhiteSpace(Sessl), "获取会话失败", "")
        End Function

        ' 获取验证码图片 URL（依赖已拿到的 Sessl）。
        Public Function GetCodeImageUrl() As String
            Return If(String.IsNullOrWhiteSpace(Sessl), "", _cfg.CodeUrlPrefix & Sessl)
        End Function

        ' 账号密码登录（login.lg）。
        Public Async Function LoginAsync(user As String, pwd As String, coode As String) As Task(Of ApiResult)
            Return Await SendAsync("login.lg", New Dictionary(Of String, String) From {
                {"user", user},
                {"pwd", pwd},
                {"coode", coode},
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
