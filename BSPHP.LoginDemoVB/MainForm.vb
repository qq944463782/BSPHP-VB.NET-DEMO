Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace BSPHP.LoginDemoVB
    Public Class MainForm
        Inherits Form

        Private ReadOnly _client As BSPHPClient = New BSPHPClient(New BSPHPConfig())
        Private ReadOnly _notice As New TextBox()
        Private ReadOnly _tabs As New TabControl()
        Private ReadOnly _status As New Label()

        Private ReadOnly _loginUser As New TextBox()
        Private ReadOnly _loginPwd As New TextBox()
        Private ReadOnly _loginCode As New TextBox()
        Private ReadOnly _smsMobile As New TextBox()
        Private ReadOnly _smsArea As New TextBox() With {.Text = "86"}
        Private ReadOnly _smsCode As New TextBox()
        Private ReadOnly _smsCaptcha As New TextBox()
        Private ReadOnly _email As New TextBox()
        Private ReadOnly _emailCode As New TextBox()
        Private ReadOnly _emailCaptcha As New TextBox()
        Private ReadOnly _registerUser As New TextBox()
        Private ReadOnly _registerPwd As New TextBox()
        Private ReadOnly _registerPwd2 As New TextBox()
        Private ReadOnly _registerMobile As New TextBox()
        Private ReadOnly _registerQuestion As New TextBox() With {.Text = "你母亲的名字？"}
        Private ReadOnly _registerAnswer As New TextBox()
        Private ReadOnly _registerCaptcha As New TextBox()
        Private ReadOnly _registerQQ As New TextBox()
        Private ReadOnly _registerMail As New TextBox()
        Private ReadOnly _registerExtension As New TextBox()
        Private ReadOnly _codeImage As New PictureBox() With {.Width = 140, .Height = 48, .SizeMode = PictureBoxSizeMode.Zoom, .BorderStyle = BorderStyle.FixedSingle}
        Private ReadOnly _smsRegUser As New TextBox()
        Private ReadOnly _smsRegMobile As New TextBox()
        Private ReadOnly _smsRegArea As New TextBox() With {.Text = "86"}
        Private ReadOnly _smsRegCode As New TextBox()
        Private ReadOnly _smsRegPwd As New TextBox()
        Private ReadOnly _smsRegPwd2 As New TextBox()
        Private ReadOnly _smsRegCaptcha As New TextBox()
        Private ReadOnly _emailRegUser As New TextBox()
        Private ReadOnly _emailRegEmail As New TextBox()
        Private ReadOnly _emailRegCode As New TextBox()
        Private ReadOnly _emailRegPwd As New TextBox()
        Private ReadOnly _emailRegPwd2 As New TextBox()
        Private ReadOnly _emailRegCaptcha As New TextBox()
        Private ReadOnly _smsResetMobile As New TextBox()
        Private ReadOnly _smsResetArea As New TextBox() With {.Text = "86"}
        Private ReadOnly _smsResetCode As New TextBox()
        Private ReadOnly _smsResetPwd As New TextBox()
        Private ReadOnly _smsResetPwd2 As New TextBox()
        Private ReadOnly _smsResetCaptcha As New TextBox()
        Private ReadOnly _emailResetEmail As New TextBox()
        Private ReadOnly _emailResetCode As New TextBox()
        Private ReadOnly _emailResetPwd As New TextBox()
        Private ReadOnly _emailResetPwd2 As New TextBox()
        Private ReadOnly _emailResetCaptcha As New TextBox()

        Public Sub New()
            Text = "BSPHP 账号模式 VB.NET Demo"
            Width = 1180
            Height = 860
            StartPosition = FormStartPosition.CenterScreen

            _notice.Multiline = True
            _notice.Dock = DockStyle.Top
            _notice.Height = 100
            _notice.ReadOnly = True
            _notice.Text = "初始化中..."
            Controls.Add(_notice)

            _tabs.Dock = DockStyle.Fill
            Controls.Add(_tabs)
            BuildTabs()
            ApplyUnifiedUiStyle(Me)

            _status.Dock = DockStyle.Bottom
            _status.Height = 30
            _status.Text = "未连接"
            Controls.Add(_status)

            AddHandler Shown, Async Sub() Await BootstrapAsync()
        End Sub

        Private Async Function BootstrapAsync() As Task
            Try
                Dim err = Await _client.BootstrapAsync()
                If err <> "" Then
                    _status.Text = err
                    _notice.Text = err
                    Return
                End If
                Dim n = Await _client.SendAsync("gg.in")
                _notice.Text = n.DataText
                _status.Text = "服务已连接"
                RefreshCaptcha()
            Catch ex As Exception
                _status.Text = "初始化异常"
                _notice.Text = ex.Message
            End Try
        End Function

        Private Sub BuildTabs()
            _tabs.TabPages.Add(MakeLoginTab())
            _tabs.TabPages.Add(MakeSmsLoginTab())
            _tabs.TabPages.Add(MakeEmailLoginTab())
            _tabs.TabPages.Add(MakeRegisterTab())
            _tabs.TabPages.Add(MakeSmsRegisterTab())
            _tabs.TabPages.Add(MakeEmailRegisterTab())
            _tabs.TabPages.Add(MakeSimpleActionTab("解绑", "jiekey.lg"))
            _tabs.TabPages.Add(MakeSimpleActionTab("充值", "chong.lg"))
            _tabs.TabPages.Add(MakeSmsResetTab())
            _tabs.TabPages.Add(MakeEmailResetTab())
            _tabs.TabPages.Add(MakeSimpleActionTab("找回密码", "backto.lg"))
            _tabs.TabPages.Add(MakeSimpleActionTab("修改密码", "password.lg"))
            _tabs.TabPages.Add(MakeSimpleActionTab("意见反馈", "liuyan.in"))
            _tabs.TabPages.Add(MakeConsoleTab())
        End Sub

        Private Function MakeLoginTab() As TabPage
            Dim tab As New TabPage("密码登录")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 7}
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

            _loginPwd.UseSystemPasswordChar = True
            panel.Controls.Add(New Label() With {.Text = "账号", .Dock = DockStyle.Fill, .TextAlign = Drawing.ContentAlignment.MiddleRight}, 0, 0)
            panel.Controls.Add(_loginUser, 1, 0)
            panel.Controls.Add(New Label() With {.Text = "密码", .Dock = DockStyle.Fill, .TextAlign = Drawing.ContentAlignment.MiddleRight}, 0, 1)
            panel.Controls.Add(_loginPwd, 1, 1)
            panel.Controls.Add(New Label() With {.Text = "验证码", .Dock = DockStyle.Fill, .TextAlign = Drawing.ContentAlignment.MiddleRight}, 0, 2)
            panel.Controls.Add(_loginCode, 1, 2)
            Dim codeBar As New FlowLayoutPanel() With {.Dock = DockStyle.Fill, .AutoSize = True}
            Dim btnRefreshCode As New Button() With {.Text = "刷新验证码", .Width = 100}
            AddHandler btnRefreshCode.Click, Sub() RefreshCaptcha()
            codeBar.Controls.Add(_codeImage)
            codeBar.Controls.Add(btnRefreshCode)
            panel.Controls.Add(New Label() With {.Text = "验证码图", .Dock = DockStyle.Fill, .TextAlign = Drawing.ContentAlignment.MiddleRight}, 0, 3)
            panel.Controls.Add(codeBar, 1, 3)

            Dim btn As New Button() With {.Text = "登录"}
            AddHandler btn.Click, Async Sub() Await LoginAsync()
            panel.Controls.Add(btn, 1, 4)
            Dim btnPrefill As New Button() With {.Text = "预填演示参数"}
            AddHandler btnPrefill.Click, Sub() PrefillLoginDemo()
            panel.Controls.Add(btnPrefill, 1, 5)

            Dim btnWeb As New Button() With {.Text = "Web登录页", .Dock = DockStyle.Top}
            AddHandler btnWeb.Click, Sub() OpenUrl("https://demo.bsphp.com/index.php?m=webapi&c=software_auth&a=index&daihao=8888888&BSphpSeSsL=" & _client.Sessl)
            panel.Controls.Add(btnWeb, 1, 6)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeSmsRegisterTab() As TabPage
            Dim tab As New TabPage("短信注册")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 9}
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            AddLabeled(panel, 0, "账号", _smsRegUser)
            AddLabeled(panel, 1, "区号", _smsRegArea)
            AddLabeled(panel, 2, "手机号", _smsRegMobile)
            AddLabeled(panel, 3, "短信码", _smsRegCode)
            AddLabeled(panel, 4, "密码", _smsRegPwd)
            _smsRegPwd.UseSystemPasswordChar = True
            AddLabeled(panel, 5, "确认密码", _smsRegPwd2)
            _smsRegPwd2.UseSystemPasswordChar = True
            AddLabeled(panel, 6, "图形验证码", _smsRegCaptcha)
            Dim btn As New Button() With {.Text = "短信注册"}
            AddHandler btn.Click, Async Sub()
                                      Dim r = Await _client.RegisterSmsAsync(_smsRegUser.Text.Trim(), _smsRegMobile.Text.Trim(), _smsRegArea.Text.Trim(), _smsRegCode.Text.Trim(), _smsRegPwd.Text, _smsRegPwd2.Text, _smsRegCaptcha.Text.Trim())
                                      ShowResult("register_sms.lg", r)
                                  End Sub
            panel.Controls.Add(btn, 1, 7)
            Dim btnPrefill As New Button() With {.Text = "预填演示参数"}
            AddHandler btnPrefill.Click, Sub()
                                            _smsRegUser.Text = "demo_sms_user"
                                            _smsRegArea.Text = "86"
                                            _smsRegMobile.Text = "13800138000"
                                            _smsRegCode.Text = "123456"
                                            _smsRegPwd.Text = "123456"
                                            _smsRegPwd2.Text = "123456"
                                            _smsRegCaptcha.Text = "8888"
                                        End Sub
            panel.Controls.Add(btnPrefill, 1, 8)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeEmailRegisterTab() As TabPage
            Dim tab As New TabPage("邮箱注册")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 8}
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            AddLabeled(panel, 0, "账号", _emailRegUser)
            AddLabeled(panel, 1, "邮箱", _emailRegEmail)
            AddLabeled(panel, 2, "邮箱码", _emailRegCode)
            AddLabeled(panel, 3, "密码", _emailRegPwd)
            _emailRegPwd.UseSystemPasswordChar = True
            AddLabeled(panel, 4, "确认密码", _emailRegPwd2)
            _emailRegPwd2.UseSystemPasswordChar = True
            AddLabeled(panel, 5, "图形验证码", _emailRegCaptcha)
            Dim btn As New Button() With {.Text = "邮箱注册"}
            AddHandler btn.Click, Async Sub()
                                      Dim r = Await _client.RegisterEmailAsync(_emailRegUser.Text.Trim(), _emailRegEmail.Text.Trim(), _emailRegCode.Text.Trim(), _emailRegPwd.Text, _emailRegPwd2.Text, _emailRegCaptcha.Text.Trim())
                                      ShowResult("register_email.lg", r)
                                  End Sub
            panel.Controls.Add(btn, 1, 6)
            Dim btnPrefill As New Button() With {.Text = "预填演示参数"}
            AddHandler btnPrefill.Click, Sub()
                                            _emailRegUser.Text = "demo_email_user"
                                            _emailRegEmail.Text = "demo@example.com"
                                            _emailRegCode.Text = "123456"
                                            _emailRegPwd.Text = "123456"
                                            _emailRegPwd2.Text = "123456"
                                            _emailRegCaptcha.Text = "8888"
                                        End Sub
            panel.Controls.Add(btnPrefill, 1, 7)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeSmsResetTab() As TabPage
            Dim tab As New TabPage("短信找回")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 7}
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            AddLabeled(panel, 0, "区号", _smsResetArea)
            AddLabeled(panel, 1, "手机号", _smsResetMobile)
            AddLabeled(panel, 2, "短信码", _smsResetCode)
            AddLabeled(panel, 3, "新密码", _smsResetPwd)
            _smsResetPwd.UseSystemPasswordChar = True
            AddLabeled(panel, 4, "确认密码", _smsResetPwd2)
            _smsResetPwd2.UseSystemPasswordChar = True
            AddLabeled(panel, 5, "图形验证码", _smsResetCaptcha)
            Dim btn As New Button() With {.Text = "短信找回"}
            AddHandler btn.Click, Async Sub()
                                      Dim r = Await _client.ResetSmsPwdAsync(_smsResetMobile.Text.Trim(), _smsResetArea.Text.Trim(), _smsResetCode.Text.Trim(), _smsResetPwd.Text, _smsResetPwd2.Text, _smsResetCaptcha.Text.Trim())
                                      ShowResult("resetpwd_sms.lg", r)
                                  End Sub
            panel.Controls.Add(btn, 1, 6)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeEmailResetTab() As TabPage
            Dim tab As New TabPage("邮箱找回")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 6}
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            AddLabeled(panel, 0, "邮箱", _emailResetEmail)
            AddLabeled(panel, 1, "邮箱码", _emailResetCode)
            AddLabeled(panel, 2, "新密码", _emailResetPwd)
            _emailResetPwd.UseSystemPasswordChar = True
            AddLabeled(panel, 3, "确认密码", _emailResetPwd2)
            _emailResetPwd2.UseSystemPasswordChar = True
            AddLabeled(panel, 4, "图形验证码", _emailResetCaptcha)
            Dim btn As New Button() With {.Text = "邮箱找回"}
            AddHandler btn.Click, Async Sub()
                                      Dim r = Await _client.ResetEmailPwdAsync(_emailResetEmail.Text.Trim(), _emailResetCode.Text.Trim(), _emailResetPwd.Text, _emailResetPwd2.Text, _emailResetCaptcha.Text.Trim())
                                      ShowResult("resetpwd_email.lg", r)
                                  End Sub
            panel.Controls.Add(btn, 1, 5)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeRegisterTab() As TabPage
            Dim tab As New TabPage("账号注册")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 11}
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 140))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            AddLabeled(panel, 0, "账号", _registerUser)
            AddLabeled(panel, 1, "密码", _registerPwd)
            _registerPwd.UseSystemPasswordChar = True
            AddLabeled(panel, 2, "确认密码", _registerPwd2)
            _registerPwd2.UseSystemPasswordChar = True
            AddLabeled(panel, 3, "手机号", _registerMobile)
            AddLabeled(panel, 4, "密保问题", _registerQuestion)
            AddLabeled(panel, 5, "密保答案", _registerAnswer)
            AddLabeled(panel, 6, "QQ", _registerQQ)
            AddLabeled(panel, 7, "邮箱", _registerMail)
            AddLabeled(panel, 8, "推广码", _registerExtension)
            AddLabeled(panel, 9, "图形验证码", _registerCaptcha)
            Dim btn As New Button() With {.Text = "注册"}
            AddHandler btn.Click, Async Sub()
                                      Dim r = Await _client.RegisterAsync(
                                          _registerUser.Text.Trim(), _registerPwd.Text, _registerPwd2.Text, _registerCaptcha.Text.Trim(),
                                          _registerMobile.Text.Trim(), _registerQuestion.Text.Trim(), _registerAnswer.Text.Trim(),
                                          _registerQQ.Text.Trim(), _registerMail.Text.Trim(), _registerExtension.Text.Trim())
                                      ShowResult("registration.lg", r)
                                  End Sub
            panel.Controls.Add(btn, 1, 10)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeSmsLoginTab() As TabPage
            Dim tab As New TabPage("短信登录")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 6}
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            panel.Controls.Add(New Label() With {.Text = "区号", .TextAlign = Drawing.ContentAlignment.MiddleRight, .Dock = DockStyle.Fill}, 0, 0)
            panel.Controls.Add(_smsArea, 1, 0)
            panel.Controls.Add(New Label() With {.Text = "手机号", .TextAlign = Drawing.ContentAlignment.MiddleRight, .Dock = DockStyle.Fill}, 0, 1)
            panel.Controls.Add(_smsMobile, 1, 1)
            panel.Controls.Add(New Label() With {.Text = "短信验证码", .TextAlign = Drawing.ContentAlignment.MiddleRight, .Dock = DockStyle.Fill}, 0, 2)
            panel.Controls.Add(_smsCode, 1, 2)
            panel.Controls.Add(New Label() With {.Text = "图形验证码", .TextAlign = Drawing.ContentAlignment.MiddleRight, .Dock = DockStyle.Fill}, 0, 3)
            panel.Controls.Add(_smsCaptcha, 1, 3)
            Dim send As New Button() With {.Text = "发送短信码"}
            AddHandler send.Click, Async Sub()
                                      Dim r = Await _client.SendSmsAsync("login", _smsMobile.Text.Trim(), _smsArea.Text.Trim(), _smsCaptcha.Text.Trim())
                                      MessageBox.Show(r.DataText, "send_sms.lg")
                                  End Sub
            panel.Controls.Add(send, 1, 4)
            Dim login As New Button() With {.Text = "短信登录"}
            AddHandler login.Click, Async Sub()
                                       Dim r = Await _client.LoginSmsAsync(_smsMobile.Text.Trim(), _smsArea.Text.Trim(), _smsCode.Text.Trim(), _smsCaptcha.Text.Trim())
                                       MessageBox.Show("code=" & If(r.Code.HasValue, r.Code.Value.ToString(), "nil") & Environment.NewLine & r.DataText, "login_sms.lg")
                                   End Sub
            panel.Controls.Add(login, 1, 5)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeEmailLoginTab() As TabPage
            Dim tab As New TabPage("邮箱登录")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 5}
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            panel.Controls.Add(New Label() With {.Text = "邮箱", .TextAlign = Drawing.ContentAlignment.MiddleRight, .Dock = DockStyle.Fill}, 0, 0)
            panel.Controls.Add(_email, 1, 0)
            panel.Controls.Add(New Label() With {.Text = "邮箱验证码", .TextAlign = Drawing.ContentAlignment.MiddleRight, .Dock = DockStyle.Fill}, 0, 1)
            panel.Controls.Add(_emailCode, 1, 1)
            panel.Controls.Add(New Label() With {.Text = "图形验证码", .TextAlign = Drawing.ContentAlignment.MiddleRight, .Dock = DockStyle.Fill}, 0, 2)
            panel.Controls.Add(_emailCaptcha, 1, 2)
            Dim send As New Button() With {.Text = "发送邮箱码"}
            AddHandler send.Click, Async Sub()
                                      Dim r = Await _client.SendEmailAsync("login", _email.Text.Trim(), _emailCaptcha.Text.Trim())
                                      MessageBox.Show(r.DataText, "send_email.lg")
                                  End Sub
            panel.Controls.Add(send, 1, 3)
            Dim login As New Button() With {.Text = "邮箱登录"}
            AddHandler login.Click, Async Sub()
                                       Dim r = Await _client.LoginEmailAsync(_email.Text.Trim(), _emailCode.Text.Trim(), _emailCaptcha.Text.Trim())
                                       MessageBox.Show("code=" & If(r.Code.HasValue, r.Code.Value.ToString(), "nil") & Environment.NewLine & r.DataText, "login_email.lg")
                                   End Sub
            panel.Controls.Add(login, 1, 4)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeSimpleActionTab(title As String, api As String, Optional payload As Dictionary(Of String, String) = Nothing) As TabPage
            Dim tab As New TabPage(title)
            Dim box As New TextBox() With {.Multiline = True, .Dock = DockStyle.Fill, .Text = "点击按钮调用 " & api & Environment.NewLine & "下方可填参数 key=value", .ScrollBars = ScrollBars.Both}
            Dim input As New TextBox() With {.Dock = DockStyle.Top, .Height = 28, .Text = ""}
            Dim btn As New Button() With {.Text = "调用 " & api, .Dock = DockStyle.Top, .Height = 36}
            AddHandler btn.Click, Async Sub()
                                      Dim p = ParseKeyValue(input.Text)
                                      If payload IsNot Nothing Then
                                          For Each kv In payload
                                              If Not p.ContainsKey(kv.Key) Then p(kv.Key) = kv.Value
                                          Next
                                      End If
                                      Dim r = Await _client.SendAsync(api, p)
                                      box.Text = "code=" & If(r.Code.HasValue, r.Code.Value.ToString(), "nil") & Environment.NewLine & r.DataText
                                  End Sub
            tab.Controls.Add(box)
            tab.Controls.Add(input)
            tab.Controls.Add(btn)
            Return tab
        End Function

        Private Function MakeConsoleTab() As TabPage
            Dim tab As New TabPage("控制台")
            Dim panel As New FlowLayoutPanel() With {.Dock = DockStyle.Fill, .AutoScroll = True}
            Dim apis = New String() {
                "date.in", "url.in", "weburl.in", "globalinfo.in",
                "appcustom.in", "v.in", "miao.in", "getsetimag.in",
                "logica.in", "logicb.in", "logicinfoa.in", "logicinfob.in",
                "getuserinfo.lg", "vipdate.lg", "userkey.lg", "timeout.lg", "cancellation.lg"
            }
            For Each a In apis
                Dim b As New Button() With {.Text = a, .Width = 180, .Height = 36}
                AddHandler b.Click, Async Sub()
                                        Dim r = Await _client.SendAsync(a)
                                        ShowResult(a, r)
                                    End Sub
                panel.Controls.Add(b)
            Next
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Async Function LoginAsync() As Task
            Dim r = Await _client.LoginAsync(_loginUser.Text.Trim(), _loginPwd.Text, _loginCode.Text.Trim())
            ShowResult("login.lg", r)
            If r.Code.HasValue AndAlso (r.Code.Value = 1011 OrElse r.Code.Value = 9908) Then
                Dim consoleWindow As New ConsoleForm(_client)
                consoleWindow.Show()
            End If
        End Function

        Private Sub PrefillLoginDemo()
            _loginUser.Text = "demo_user"
            _loginPwd.Text = "123456"
            _loginCode.Text = "8888"
            _smsArea.Text = "86"
            _smsMobile.Text = "13800138000"
            _smsCode.Text = "123456"
            _smsCaptcha.Text = "8888"
            _email.Text = "demo@example.com"
            _emailCode.Text = "123456"
            _emailCaptcha.Text = "8888"
        End Sub

        Private Sub RefreshCaptcha()
            Try
                Dim u = _client.GetCodeImageUrl()
                If String.IsNullOrWhiteSpace(u) Then Return
                _codeImage.Load(u & "&_=" & DateTimeOffset.Now.ToUnixTimeMilliseconds().ToString())
            Catch
            End Try
        End Sub

        Private Shared Function ParseKeyValue(line As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrWhiteSpace(line) Then Return result
            Dim entries = line.Split(";"c, StringSplitOptions.RemoveEmptyEntries)
            For Each e In entries
                Dim idx = e.IndexOf("="c)
                If idx > 0 Then
                    Dim k = e.Substring(0, idx).Trim()
                    Dim v = e.Substring(idx + 1).Trim()
                    If k <> "" Then result(k) = v
                End If
            Next
            Return result
        End Function

        Private Sub AddLabeled(panel As TableLayoutPanel, row As Integer, title As String, input As Control)
            panel.Controls.Add(New Label() With {.Text = title, .Dock = DockStyle.Fill, .TextAlign = Drawing.ContentAlignment.MiddleRight}, 0, row)
            panel.Controls.Add(input, 1, row)
        End Sub

        Private Shared Sub ShowResult(title As String, r As ApiResult)
            MessageBox.Show("code=" & If(r.Code.HasValue, r.Code.Value.ToString(), "nil") & Environment.NewLine & r.DataText, title)
        End Function

        Private Shared Sub OpenUrl(url As String)
            Process.Start(New ProcessStartInfo With {.FileName = url, .UseShellExecute = True})
        End Sub

        Private Sub ApplyUnifiedUiStyle(root As Control)
            For Each c As Control In root.Controls
                If TypeOf c Is Button Then
                    c.Height = 34
                    c.Width = Math.Max(c.Width, 130)
                ElseIf TypeOf c Is TextBox Then
                    c.Height = 28
                    c.Font = New Drawing.Font("Microsoft YaHei UI", 10.0F)
                ElseIf TypeOf c Is Label Then
                    c.Font = New Drawing.Font("Microsoft YaHei UI", 10.0F)
                End If
                If c.HasChildren Then ApplyUnifiedUiStyle(c)
            Next
        End Sub
    End Class
End Namespace
