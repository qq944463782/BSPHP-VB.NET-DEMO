Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Net
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
        Private ReadOnly _recoverUser As New TextBox()
        Private ReadOnly _recoverQuestion As New TextBox() With {.Text = "你母亲的名字？"}
        Private ReadOnly _recoverAnswer As New TextBox()
        Private ReadOnly _recoverPwd As New TextBox()
        Private ReadOnly _recoverPwd2 As New TextBox()
        Private ReadOnly _recoverCaptcha As New TextBox()
        Private ReadOnly _changeUser As New TextBox()
        Private ReadOnly _changeOldPwd As New TextBox()
        Private ReadOnly _changeNewPwd As New TextBox()
        Private ReadOnly _changeNewPwd2 As New TextBox()
        Private ReadOnly _changeCaptcha As New TextBox()
        Private ReadOnly _feedbackUser As New TextBox()
        Private ReadOnly _feedbackPwd As New TextBox()
        Private ReadOnly _feedbackTitle As New TextBox()
        Private ReadOnly _feedbackQQ As New TextBox()
        Private ReadOnly _feedbackType As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList}
        Private ReadOnly _feedbackContent As New TextBox() With {.Multiline = True, .Height = 96, .ScrollBars = ScrollBars.Vertical}
        Private ReadOnly _feedbackCaptcha As New TextBox()
        Private ReadOnly _unbindUser As New TextBox()
        Private ReadOnly _unbindPwd As New TextBox()
        Private ReadOnly _rechargeUser As New TextBox()
        Private ReadOnly _rechargeUserPwd As New TextBox()
        Private ReadOnly _rechargeVerifyPwd As New CheckBox() With {.Text = "验证用户密码(防充错)", .Checked = True, .AutoSize = True}
        Private ReadOnly _rechargeCard As New TextBox()
        Private ReadOnly _rechargeCardPwd As New TextBox()
        Private ReadOnly _loginCaptchaPictures As New List(Of PictureBox)()
        Private ReadOnly _regCaptchaPictures As New List(Of PictureBox)()
        Private ReadOnly _backCaptchaPictures As New List(Of PictureBox)()
        Private ReadOnly _sayCaptchaPictures As New List(Of PictureBox)()
        Private _codeLoginEnabled As Boolean = True
        Private _codeRegEnabled As Boolean = True
        Private _codeBackEnabled As Boolean = True
        Private _codeSayEnabled As Boolean = True
        Private _connected As Boolean = False
        Private _loggedIn As Boolean = False
        Private _loginEndTime As String = ""
        Private _consoleWindow As ConsoleForm = Nothing

        Public Sub New()
            Text = "BSPHP 账号模式 VB.NET Demo"
            Width = 1040
            Height = 760
            StartPosition = FormStartPosition.CenterScreen

            _notice.Multiline = True
            _notice.Dock = DockStyle.Top
            _notice.Height = 88
            _notice.ReadOnly = True
            _notice.ScrollBars = ScrollBars.Vertical
            _notice.Font = New Drawing.Font("Microsoft YaHei UI", 9.5F)
            _notice.Text = "初始化中..."
            Controls.Add(_notice)

            Dim tabHost As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(0, 18, 0, 0)}
            _tabs.Dock = DockStyle.Fill
            ' 使用系统默认页签绘制，优先保证“页签栏一定可见”。
            _tabs.Multiline = True
            _tabs.SizeMode = TabSizeMode.Normal
            _tabs.ItemSize = New Drawing.Size(0, 0)
            _tabs.Padding = New Drawing.Point(8, 4)
            _tabs.HotTrack = True
            _tabs.Alignment = TabAlignment.Top
            tabHost.Controls.Add(_tabs)
            Controls.Add(tabHost)
            BuildTabs()
            _feedbackType.Items.AddRange(New Object() {"建议反馈", "BUG", "使用问题"})
            _feedbackType.SelectedIndex = 0
            ApplyUnifiedUiStyle(Me)
            ApplyMacLikeTheme(Me)

            _status.Dock = DockStyle.Bottom
            _status.Height = 24
            _status.Text = "服务未连接"
            Controls.Add(_status)

            AddHandler Shown, Async Sub() Await BootstrapAsync()
        End Sub

        Private Async Function BootstrapAsync() As Task
            Try
                Dim err = Await _client.BootstrapAsync()
                If err <> "" Then
                    _connected = False
                    UpdateStatusBar(err)
                    _notice.Text = err
                    Return
                End If
                Dim n = Await _client.SendAsync("gg.in")
                _notice.Text = "公告：" & Environment.NewLine & n.DataText
                _connected = True
                UpdateStatusBar()
                Await LoadCaptchaSwitchesAsync()
                RefreshCaptcha()
            Catch ex As Exception
                _connected = False
                UpdateStatusBar("初始化异常")
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
            _tabs.TabPages.Add(MakeUnbindTab())
            _tabs.TabPages.Add(MakeRechargeTab())
            _tabs.TabPages.Add(MakeSmsResetTab())
            _tabs.TabPages.Add(MakeEmailResetTab())
            _tabs.TabPages.Add(MakeRecoverTab())
            _tabs.TabPages.Add(MakeChangePwdTab())
            _tabs.TabPages.Add(MakeFeedbackTab())
            _tabs.TabPages.Add(MakeConsoleTab())
        End Sub

        Private Async Function LoadCaptchaSwitchesAsync() As Task
            Dim r = Await _client.SendAsync("getsetimag.in", New Dictionary(Of String, String) From {
                {"type", "INGES_LOGIN|INGES_RE|INGES_MACK|INGES_SAY"}
            })
            Dim parts = r.DataText.Split("|"c)
            If parts.Length >= 4 Then
                _codeLoginEnabled = parts(0).Trim().Equals("checked", StringComparison.OrdinalIgnoreCase)
                _codeRegEnabled = parts(1).Trim().Equals("checked", StringComparison.OrdinalIgnoreCase)
                _codeBackEnabled = parts(2).Trim().Equals("checked", StringComparison.OrdinalIgnoreCase)
                _codeSayEnabled = parts(3).Trim().Equals("checked", StringComparison.OrdinalIgnoreCase)
            End If
            ApplyCaptchaVisibility()
        End Function

        Private Sub ApplyCaptchaVisibility()
            _loginCode.Visible = _codeLoginEnabled
            _smsCaptcha.Visible = True
            _emailCaptcha.Visible = True
            _registerCaptcha.Visible = _codeRegEnabled
            _smsRegCaptcha.Visible = True
            _emailRegCaptcha.Visible = True
            _smsResetCaptcha.Visible = True
            _emailResetCaptcha.Visible = True
            _recoverCaptcha.Visible = _codeBackEnabled
            _feedbackCaptcha.Visible = _codeSayEnabled

            SetCaptchaVisualVisible(_codeImage, _codeLoginEnabled)

            For Each pic In _loginCaptchaPictures
                SetCaptchaVisualVisible(pic, True)
            Next
            For Each pic In _regCaptchaPictures
                SetCaptchaVisualVisible(pic, True)
            Next
            For Each pic In _backCaptchaPictures
                SetCaptchaVisualVisible(pic, True)
            Next
            For Each pic In _sayCaptchaPictures
                SetCaptchaVisualVisible(pic, _codeSayEnabled)
            Next
        End Sub

        Private Function MakeLoginTab() As TabPage
            Dim tab As New TabPage("密码登录")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 8}
            ConfigureCompactGrid(panel, 120)
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
            AddHandler btnWeb.Click, Sub()
                                         Dim web As New WebLoginForm(
                                             _client,
                                             Async Function()
                                                 _loggedIn = True
                                                 _loginEndTime = Await FetchLoginEndTimeAsync()
                                                 UpdateStatusBar()
                                                 OpenOrActivateConsoleWindow()
                                             End Function)
                                         web.Show(Me)
                                     End Sub
            panel.Controls.Add(btnWeb, 1, 6)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeSmsRegisterTab() As TabPage
            Dim tab As New TabPage("短信注册")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 10}
            ConfigureCompactGrid(panel, 120)
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
            Dim smsRegCaptchaImage As New PictureBox() With {.Width = 140, .Height = 48, .SizeMode = PictureBoxSizeMode.Zoom, .BorderStyle = BorderStyle.FixedSingle}
            _regCaptchaPictures.Add(smsRegCaptchaImage)
            panel.Controls.Add(New Label() With {.Text = "验证码图", .Dock = DockStyle.Fill, .TextAlign = Drawing.ContentAlignment.MiddleRight}, 0, 7)
            panel.Controls.Add(BuildCaptchaBar(smsRegCaptchaImage), 1, 7)
            Dim btn As New Button() With {.Text = "短信注册"}
            AddHandler btn.Click, Async Sub()
                                      Dim r = Await _client.RegisterSmsAsync(_smsRegUser.Text.Trim(), _smsRegMobile.Text.Trim(), _smsRegArea.Text.Trim(), _smsRegCode.Text.Trim(), _smsRegPwd.Text, _smsRegPwd2.Text, _smsRegCaptcha.Text.Trim())
                                      ShowResult("register_sms.lg", r)
                                  End Sub
            panel.Controls.Add(btn, 1, 8)
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
            panel.Controls.Add(btnPrefill, 1, 9)
            tab.Controls.Add(panel)
            LoadCaptchaInto(smsRegCaptchaImage)
            Return tab
        End Function

        Private Function MakeEmailRegisterTab() As TabPage
            Dim tab As New TabPage("邮箱注册")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 9}
            ConfigureCompactGrid(panel, 120)
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
            Dim emailRegCaptchaImage As New PictureBox() With {.Width = 140, .Height = 48, .SizeMode = PictureBoxSizeMode.Zoom, .BorderStyle = BorderStyle.FixedSingle}
            _regCaptchaPictures.Add(emailRegCaptchaImage)
            panel.Controls.Add(New Label() With {.Text = "验证码图", .Dock = DockStyle.Fill, .TextAlign = Drawing.ContentAlignment.MiddleRight}, 0, 6)
            panel.Controls.Add(BuildCaptchaBar(emailRegCaptchaImage), 1, 6)
            Dim btn As New Button() With {.Text = "邮箱注册"}
            AddHandler btn.Click, Async Sub()
                                      Dim r = Await _client.RegisterEmailAsync(_emailRegUser.Text.Trim(), _emailRegEmail.Text.Trim(), _emailRegCode.Text.Trim(), _emailRegPwd.Text, _emailRegPwd2.Text, _emailRegCaptcha.Text.Trim())
                                      ShowResult("register_email.lg", r)
                                  End Sub
            panel.Controls.Add(btn, 1, 7)
            Dim btnPrefill As New Button() With {.Text = "预填演示参数"}
            AddHandler btnPrefill.Click, Sub()
                                             _emailRegUser.Text = "demo_email_user"
                                             _emailRegEmail.Text = "demo@example.com"
                                             _emailRegCode.Text = "123456"
                                             _emailRegPwd.Text = "123456"
                                             _emailRegPwd2.Text = "123456"
                                             _emailRegCaptcha.Text = "8888"
                                         End Sub
            panel.Controls.Add(btnPrefill, 1, 8)
            tab.Controls.Add(panel)
            LoadCaptchaInto(emailRegCaptchaImage)
            Return tab
        End Function

        Private Function MakeSmsResetTab() As TabPage
            Dim tab As New TabPage("短信找回")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 8}
            ConfigureCompactGrid(panel, 120)
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
            Dim smsResetCaptchaImage As New PictureBox() With {.Width = 140, .Height = 48, .SizeMode = PictureBoxSizeMode.Zoom, .BorderStyle = BorderStyle.FixedSingle}
            _backCaptchaPictures.Add(smsResetCaptchaImage)
            panel.Controls.Add(New Label() With {.Text = "验证码图", .Dock = DockStyle.Fill, .TextAlign = Drawing.ContentAlignment.MiddleRight}, 0, 6)
            panel.Controls.Add(BuildCaptchaBar(smsResetCaptchaImage), 1, 6)
            Dim btn As New Button() With {.Text = "短信找回"}
            AddHandler btn.Click, Async Sub()
                                      Dim r = Await _client.ResetSmsPwdAsync(_smsResetMobile.Text.Trim(), _smsResetArea.Text.Trim(), _smsResetCode.Text.Trim(), _smsResetPwd.Text, _smsResetPwd2.Text, _smsResetCaptcha.Text.Trim())
                                      ShowResult("resetpwd_sms.lg", r)
                                  End Sub
            panel.Controls.Add(btn, 1, 7)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeEmailResetTab() As TabPage
            Dim tab As New TabPage("邮箱找回")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 7}
            ConfigureCompactGrid(panel, 120)
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            AddLabeled(panel, 0, "邮箱", _emailResetEmail)
            AddLabeled(panel, 1, "邮箱码", _emailResetCode)
            AddLabeled(panel, 2, "新密码", _emailResetPwd)
            _emailResetPwd.UseSystemPasswordChar = True
            AddLabeled(panel, 3, "确认密码", _emailResetPwd2)
            _emailResetPwd2.UseSystemPasswordChar = True
            AddLabeled(panel, 4, "图形验证码", _emailResetCaptcha)
            Dim emailResetCaptchaImage As New PictureBox() With {.Width = 140, .Height = 48, .SizeMode = PictureBoxSizeMode.Zoom, .BorderStyle = BorderStyle.FixedSingle}
            _backCaptchaPictures.Add(emailResetCaptchaImage)
            panel.Controls.Add(New Label() With {.Text = "验证码图", .Dock = DockStyle.Fill, .TextAlign = Drawing.ContentAlignment.MiddleRight}, 0, 5)
            panel.Controls.Add(BuildCaptchaBar(emailResetCaptchaImage), 1, 5)
            Dim btn As New Button() With {.Text = "邮箱找回"}
            AddHandler btn.Click, Async Sub()
                                      Dim r = Await _client.ResetEmailPwdAsync(_emailResetEmail.Text.Trim(), _emailResetCode.Text.Trim(), _emailResetPwd.Text, _emailResetPwd2.Text, _emailResetCaptcha.Text.Trim())
                                      ShowResult("resetpwd_email.lg", r)
                                  End Sub
            panel.Controls.Add(btn, 1, 6)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeRecoverTab() As TabPage
            Dim tab As New TabPage("找回密码")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 7}
            ConfigureCompactGrid(panel, 130)
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            AddLabeled(panel, 0, "账号", _recoverUser)
            AddLabeled(panel, 1, "密保问题", _recoverQuestion)
            AddLabeled(panel, 2, "密保答案", _recoverAnswer)
            AddLabeled(panel, 3, "新密码", _recoverPwd)
            _recoverPwd.UseSystemPasswordChar = True
            AddLabeled(panel, 4, "确认密码", _recoverPwd2)
            _recoverPwd2.UseSystemPasswordChar = True
            AddLabeled(panel, 5, "图形验证码", _recoverCaptcha)
            Dim btn As New Button() With {.Text = "找回密码"}
            AddHandler btn.Click, Async Sub()
                                      Dim r = Await _client.ResetPwdAsync(_recoverUser.Text.Trim(), _recoverPwd.Text, _recoverPwd2.Text, _recoverQuestion.Text.Trim(), _recoverAnswer.Text.Trim(), _recoverCaptcha.Text.Trim())
                                      ShowResult("backto.lg", r)
                                  End Sub
            panel.Controls.Add(btn, 1, 6)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeChangePwdTab() As TabPage
            Dim tab As New TabPage("修改密码")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 6}
            ConfigureCompactGrid(panel, 130)
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            AddLabeled(panel, 0, "账号", _changeUser)
            AddLabeled(panel, 1, "原密码", _changeOldPwd)
            _changeOldPwd.UseSystemPasswordChar = True
            AddLabeled(panel, 2, "新密码", _changeNewPwd)
            _changeNewPwd.UseSystemPasswordChar = True
            AddLabeled(panel, 3, "确认新密码", _changeNewPwd2)
            _changeNewPwd2.UseSystemPasswordChar = True
            AddLabeled(panel, 4, "图形验证码", _changeCaptcha)
            Dim btn As New Button() With {.Text = "修改密码"}
            AddHandler btn.Click, Async Sub()
                                      Dim r = Await _client.ChangePwdAsync(_changeUser.Text.Trim(), _changeOldPwd.Text, _changeNewPwd.Text, _changeNewPwd2.Text, _changeCaptcha.Text.Trim())
                                      ShowResult("password.lg", r)
                                  End Sub
            panel.Controls.Add(btn, 1, 5)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeFeedbackTab() As TabPage
            Dim tab As New TabPage("意见反馈")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 8}
            ConfigureCompactGrid(panel, 130)
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            AddLabeled(panel, 0, "账号", _feedbackUser)
            AddLabeled(panel, 1, "密码", _feedbackPwd)
            _feedbackPwd.UseSystemPasswordChar = True
            AddLabeled(panel, 2, "标题", _feedbackTitle)
            AddLabeled(panel, 3, "联系方式(QQ)", _feedbackQQ)
            AddLabeled(panel, 4, "类型", _feedbackType)
            AddLabeled(panel, 5, "内容", _feedbackContent)
            AddLabeled(panel, 6, "图形验证码", _feedbackCaptcha)
            Dim btn As New Button() With {.Text = "提交反馈"}
            AddHandler btn.Click, Async Sub()
                                      Dim t As String = If(_feedbackType.SelectedItem IsNot Nothing, _feedbackType.SelectedItem.ToString(), "建议反馈")
                                      Dim r = Await _client.FeedbackAsync(_feedbackUser.Text.Trim(), _feedbackPwd.Text, _feedbackTitle.Text.Trim(), _feedbackQQ.Text.Trim(), t, _feedbackContent.Text.Trim(), _feedbackCaptcha.Text.Trim())
                                      ShowResult("liuyan.in", r)
                                  End Sub
            panel.Controls.Add(btn, 1, 7)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeRegisterTab() As TabPage
            Dim tab As New TabPage("账号注册")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 11}
            ConfigureCompactGrid(panel, 140)
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
            ConfigureCompactGrid(panel, 120)
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
            Dim smsLoginCaptchaImage As New PictureBox() With {.Width = 140, .Height = 48, .SizeMode = PictureBoxSizeMode.Zoom, .BorderStyle = BorderStyle.FixedSingle}
            _loginCaptchaPictures.Add(smsLoginCaptchaImage)
            panel.Controls.Add(New Label() With {.Text = "验证码图", .TextAlign = Drawing.ContentAlignment.MiddleRight, .Dock = DockStyle.Fill}, 0, 4)
            panel.Controls.Add(BuildCaptchaBar(smsLoginCaptchaImage), 1, 4)
            Dim send As New Button() With {.Text = "发送短信码"}
            AddHandler send.Click, Async Sub()
                                       Dim r = Await _client.SendSmsAsync("login", _smsMobile.Text.Trim(), _smsArea.Text.Trim(), _smsCaptcha.Text.Trim())
                                       MessageBox.Show(r.DataText, "send_sms.lg")
                                   End Sub
            panel.Controls.Add(send, 1, 5)
            Dim login As New Button() With {.Text = "短信登录"}
            AddHandler login.Click, Async Sub()
                                        Dim r = Await _client.LoginSmsAsync(_smsMobile.Text.Trim(), _smsArea.Text.Trim(), _smsCode.Text.Trim(), _smsCaptcha.Text.Trim())
                                        ShowResult("login_sms.lg", r)
                                        Await HandleLoginResultAsync(r)
                                    End Sub
            panel.Controls.Add(login, 1, 6)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeEmailLoginTab() As TabPage
            Dim tab As New TabPage("邮箱登录")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 6}
            ConfigureCompactGrid(panel, 120)
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            panel.Controls.Add(New Label() With {.Text = "邮箱", .TextAlign = Drawing.ContentAlignment.MiddleRight, .Dock = DockStyle.Fill}, 0, 0)
            panel.Controls.Add(_email, 1, 0)
            panel.Controls.Add(New Label() With {.Text = "邮箱验证码", .TextAlign = Drawing.ContentAlignment.MiddleRight, .Dock = DockStyle.Fill}, 0, 1)
            panel.Controls.Add(_emailCode, 1, 1)
            panel.Controls.Add(New Label() With {.Text = "图形验证码", .TextAlign = Drawing.ContentAlignment.MiddleRight, .Dock = DockStyle.Fill}, 0, 2)
            panel.Controls.Add(_emailCaptcha, 1, 2)
            Dim emailLoginCaptchaImage As New PictureBox() With {.Width = 140, .Height = 48, .SizeMode = PictureBoxSizeMode.Zoom, .BorderStyle = BorderStyle.FixedSingle}
            _loginCaptchaPictures.Add(emailLoginCaptchaImage)
            panel.Controls.Add(New Label() With {.Text = "验证码图", .TextAlign = Drawing.ContentAlignment.MiddleRight, .Dock = DockStyle.Fill}, 0, 3)
            panel.Controls.Add(BuildCaptchaBar(emailLoginCaptchaImage), 1, 3)
            Dim send As New Button() With {.Text = "发送邮箱码"}
            AddHandler send.Click, Async Sub()
                                       Dim r = Await _client.SendEmailAsync("login", _email.Text.Trim(), _emailCaptcha.Text.Trim())
                                       MessageBox.Show(r.DataText, "send_email.lg")
                                   End Sub
            panel.Controls.Add(send, 1, 4)
            Dim login As New Button() With {.Text = "邮箱登录"}
            AddHandler login.Click, Async Sub()
                                        Dim r = Await _client.LoginEmailAsync(_email.Text.Trim(), _emailCode.Text.Trim(), _emailCaptcha.Text.Trim())
                                        ShowResult("login_email.lg", r)
                                        Await HandleLoginResultAsync(r)
                                    End Sub
            panel.Controls.Add(login, 1, 5)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeUnbindTab() As TabPage
            Dim tab As New TabPage("解绑")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 3}
            ConfigureCompactGrid(panel, 120)
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            AddLabeled(panel, 0, "账号", _unbindUser)
            AddLabeled(panel, 1, "密码", _unbindPwd)
            _unbindPwd.UseSystemPasswordChar = True
            Dim btn As New Button() With {.Text = "解绑"}
            AddHandler btn.Click, Async Sub()
                                      Dim r = Await _client.UnbindAsync(_unbindUser.Text.Trim(), _unbindPwd.Text)
                                      ShowResult("jiekey.lg", r)
                                  End Sub
            panel.Controls.Add(btn, 1, 2)
            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeRechargeTab() As TabPage
            Dim tab As New TabPage("充值")
            Dim panel As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 6}
            ConfigureCompactGrid(panel, 140)
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 140))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            AddLabeled(panel, 0, "账号", _rechargeUser)
            AddLabeled(panel, 1, "账号密码", _rechargeUserPwd)
            _rechargeUserPwd.UseSystemPasswordChar = True
            panel.Controls.Add(_rechargeVerifyPwd, 1, 2)
            AddLabeled(panel, 3, "充值卡号", _rechargeCard)
            AddLabeled(panel, 4, "充值卡密码", _rechargeCardPwd)
            _rechargeCardPwd.UseSystemPasswordChar = True
            Dim btn As New Button() With {.Text = "充值"}
            AddHandler btn.Click, Async Sub()
                                      Dim r = Await _client.RechargeAsync(_rechargeUser.Text.Trim(), _rechargeUserPwd.Text, _rechargeVerifyPwd.Checked, _rechargeCard.Text.Trim(), _rechargeCardPwd.Text)
                                      ShowResult("chong.lg", r)
                                  End Sub
            panel.Controls.Add(btn, 1, 5)
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
            Dim panel As New FlowLayoutPanel() With {.Dock = DockStyle.Fill, .AutoScroll = True, .FlowDirection = FlowDirection.TopDown, .WrapContents = False}
            ConfigureStackPanel(panel)

            panel.Controls.Add(MakeConsoleGroup("公用接口", New String() {
                "服务器时间::date.in", "预设URL::url.in", "Web地址::weburl.in", "全局配置::globalinfo.in",
                "验证码开关(全部)::getsetimag.in", "逻辑值A::logica.in", "逻辑值B::logicb.in", "逻辑值A内容::logicinfoa.in", "逻辑值B内容::logicinfob.in"
            }))

            panel.Controls.Add(MakeConsoleGroup("自定义配置模型", New String() {
                "软件配置::appcustom.in?info=myapp", "VIP配置::appcustom.in?info=myvip", "登录配置::appcustom.in?info=mylogin"
            }))

            panel.Controls.Add(MakeConsoleGroup("通用接口", New String() {
                "获取版本::v.in", "获取软件描述::miao.in"
            }))

            panel.Controls.Add(MakeConsoleGroup("登录模式接口", New String() {
                "取用户信息::getuserinfo.lg", "检测到期::vipdate.lg", "用户特征Key::userkey.lg", "心跳包更新::timeout.lg", "注销登录::cancellation.lg"
            }))

            Dim renewBar As New FlowLayoutPanel() With {.AutoSize = True, .FlowDirection = FlowDirection.LeftToRight, .WrapContents = True}
            renewBar.Controls.Add(New Label() With {.Text = "续费订阅推广", .AutoSize = True, .Font = New Drawing.Font("Microsoft YaHei UI", 9.5F, Drawing.FontStyle.Bold)})
            Dim bRenew As New Button() With {.Text = "续费订阅(直接)", .Width = 160, .Height = 32}
            AddHandler bRenew.Click, Sub() OpenUrl("https://demo.bsphp.com/index.php?m=webapi&c=salecard_renew&a=index&daihao=8888888")
            Dim bCard As New Button() With {.Text = "购买充值卡", .Width = 130, .Height = 32}
            AddHandler bCard.Click, Sub() OpenUrl("https://demo.bsphp.com/index.php?m=webapi&c=salecard_gencard&a=index&daihao=8888888")
            Dim bStock As New Button() With {.Text = "购买库存卡", .Width = 130, .Height = 32}
            AddHandler bStock.Click, Sub() OpenUrl("https://demo.bsphp.com/index.php?m=webapi&c=salecard_salecard&a=index&daihao=8888888")
            renewBar.Controls.Add(bRenew)
            renewBar.Controls.Add(bCard)
            renewBar.Controls.Add(bStock)
            panel.Controls.Add(renewBar)

            tab.Controls.Add(panel)
            Return tab
        End Function

        ' 通过 “中文名::api?query” 映射来生成控制台分组按钮。
        Private Function MakeConsoleGroup(title As String, apiSpecs As IEnumerable(Of String)) As GroupBox
            Dim g As New GroupBox() With {.Text = title, .AutoSize = True}
            Dim p As New FlowLayoutPanel() With {.AutoSize = True, .FlowDirection = FlowDirection.LeftToRight, .WrapContents = True, .Padding = New Padding(8)}
            For Each spec In apiSpecs
                Dim caption As String = spec
                Dim route As String = spec
                Dim splitIdx = spec.IndexOf("::", StringComparison.Ordinal)
                If splitIdx > 0 Then
                    caption = spec.Substring(0, splitIdx)
                    route = spec.Substring(splitIdx + 2)
                End If
                Dim btn As New Button() With {.Text = caption, .Width = 180, .Height = 32}
                AddHandler btn.Click, Async Sub()
                                          Dim api As String = route
                                          Dim extra As Dictionary(Of String, String) = Nothing
                                          Dim q = route.IndexOf("?"c)
                                          If q > 0 Then
                                              api = route.Substring(0, q)
                                              extra = ParseQuery(route.Substring(q + 1))
                                          End If
                                          Dim r = Await _client.SendAsync(api, extra)
                                          If String.Equals(api, "cancellation.lg", StringComparison.OrdinalIgnoreCase) Then
                                              _loggedIn = False
                                              _loginEndTime = ""
                                              UpdateStatusBar()
                                          ElseIf String.Equals(api, "vipdate.lg", StringComparison.OrdinalIgnoreCase) Then
                                              _loginEndTime = If(r.DataText, "").Trim()
                                              UpdateStatusBar()
                                          End If
                                          ShowResult(caption, r)
                                      End Sub
                p.Controls.Add(btn)
            Next
            g.Controls.Add(p)
            Return g
        End Function

        Private Async Function LoginAsync() As Task
            Dim r = Await _client.LoginAsync(_loginUser.Text.Trim(), _loginPwd.Text, _loginCode.Text.Trim())
            ShowResult("login.lg", r)
            Await HandleLoginResultAsync(r)
        End Function

        ' 登录成功后的统一状态处理：
        ' - 标记已登录
        ' - 拉取并缓存到期时间
        ' - 激活（或打开）控制台窗口
        Private Async Function HandleLoginResultAsync(r As ApiResult) As Task
            If Not r.Code.HasValue Then Return
            If r.Code.Value = 1011 OrElse r.Code.Value = 9908 OrElse r.Code.Value = 5031 Then
                _loggedIn = True
                _loginEndTime = Await FetchLoginEndTimeAsync()
                UpdateStatusBar()
                OpenOrActivateConsoleWindow()
            End If
        End Function

        ' 优先从 getuserinfo.lg(UserVipDate) 取到期时间，不行再回退 vipdate.lg。
        Private Async Function FetchLoginEndTimeAsync() As Task(Of String)
            Dim info = Await _client.SendAsync("getuserinfo.lg", New Dictionary(Of String, String) From {
                {"info", "UserVipDate"}
            })
            Dim s = If(info.DataText, "").Trim()
            If s <> "" Then
                Dim idx = s.LastIndexOf("="c)
                If idx >= 0 AndAlso idx < s.Length - 1 Then
                    s = s.Substring(idx + 1).Trim()
                End If
                If s <> "" Then Return s
            End If
            Dim vip = Await _client.SendAsync("vipdate.lg")
            Return If(vip.DataText, "").Trim()
        End Function

        ' 控制台窗口只保留一个实例，重复登录时直接激活既有窗口。
        Private Sub OpenOrActivateConsoleWindow()
            If _consoleWindow IsNot Nothing AndAlso Not _consoleWindow.IsDisposed Then
                _consoleWindow.WindowState = FormWindowState.Normal
                _consoleWindow.BringToFront()
                _consoleWindow.Activate()
                Return
            End If
            _consoleWindow = New ConsoleForm(_client)
            AddHandler _consoleWindow.FormClosed, Sub() _consoleWindow = Nothing
            _consoleWindow.Show(Me)
        End Sub

        ' 统一状态栏渲染入口，避免散落在各事件里拼接状态文本。
        Private Sub UpdateStatusBar(Optional fallback As String = "")
            If fallback <> "" Then
                _status.Text = fallback
                Return
            End If
            If Not _connected Then
                _status.Text = "服务未连接"
                Return
            End If
            If _loggedIn Then
                Dim endText = If(String.IsNullOrWhiteSpace(_loginEndTime), "-", _loginEndTime)
                _status.Text = "服务已连接 | 已登录 | 到期时间: " & endText
            Else
                _status.Text = "服务已连接"
            End If
        End Sub

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
                Dim ts As Long = CLng((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalMilliseconds)
                Dim finalUrl = u & "&_=" & ts.ToString()
                _codeImage.ImageLocation = finalUrl
                _codeImage.LoadAsync()
                For Each pic In _loginCaptchaPictures
                    If pic IsNot Nothing Then
                        pic.ImageLocation = finalUrl
                        pic.LoadAsync()
                    End If
                Next
                For Each pic In _regCaptchaPictures
                    If pic IsNot Nothing Then
                        pic.ImageLocation = finalUrl
                        pic.LoadAsync()
                    End If
                Next
                For Each pic In _backCaptchaPictures
                    If pic IsNot Nothing Then
                        pic.ImageLocation = finalUrl
                        pic.LoadAsync()
                    End If
                Next
                For Each pic In _sayCaptchaPictures
                    If pic IsNot Nothing Then
                        pic.ImageLocation = finalUrl
                        pic.LoadAsync()
                    End If
                Next
            Catch
                Try
                    Dim u = _client.GetCodeImageUrl()
                    If String.IsNullOrWhiteSpace(u) Then Return
                    Dim ts As Long = CLng((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalMilliseconds)
                    Dim finalUrl = u & "&_=" & ts.ToString()
                    Using wc As New WebClient()
                        Dim bytes = wc.DownloadData(finalUrl)
                        Using ms As New MemoryStream(bytes)
                            _codeImage.Image = Drawing.Image.FromStream(ms)
                        End Using
                    End Using
                    For Each pic In _loginCaptchaPictures
                        LoadCaptchaInto(pic)
                    Next
                    For Each pic In _regCaptchaPictures
                        LoadCaptchaInto(pic)
                    Next
                    For Each pic In _backCaptchaPictures
                        LoadCaptchaInto(pic)
                    Next
                    For Each pic In _sayCaptchaPictures
                        LoadCaptchaInto(pic)
                    Next
                Catch
                End Try
            End Try
        End Sub

        Private Function BuildCaptchaBar(pic As PictureBox) As FlowLayoutPanel
            Dim bar As New FlowLayoutPanel() With {.Dock = DockStyle.Fill, .AutoSize = True}
            Dim btnRefresh As New Button() With {.Text = "刷新验证码", .Width = 100}
            AddHandler btnRefresh.Click, Sub() LoadCaptchaInto(pic)
            bar.Controls.Add(pic)
            bar.Controls.Add(btnRefresh)
            Return bar
        End Function

        Private Sub LoadCaptchaInto(pic As PictureBox)
            Try
                Dim u = _client.GetCodeImageUrl()
                If String.IsNullOrWhiteSpace(u) Then Return
                Dim ts As Long = CLng((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalMilliseconds)
                Dim finalUrl = u & "&_=" & ts.ToString()
                pic.ImageLocation = finalUrl
                pic.LoadAsync()
            Catch
                Try
                    Dim u = _client.GetCodeImageUrl()
                    If String.IsNullOrWhiteSpace(u) Then Return
                    Dim ts As Long = CLng((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalMilliseconds)
                    Dim finalUrl = u & "&_=" & ts.ToString()
                    Using wc As New WebClient()
                        Dim bytes = wc.DownloadData(finalUrl)
                        Using ms As New MemoryStream(bytes)
                            pic.Image = Drawing.Image.FromStream(ms)
                        End Using
                    End Using
                Catch
                End Try
            End Try
        End Sub

        Private Shared Sub SetCaptchaVisualVisible(pic As PictureBox, visible As Boolean)
            If pic Is Nothing Then Return
            pic.Visible = visible
            If pic.Parent IsNot Nothing Then
                pic.Parent.Visible = visible
            End If
        End Sub

        Private Shared Function ParseKeyValue(line As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrWhiteSpace(line) Then Return result
            Dim entries = line.Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
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

        Private Shared Function ParseQuery(query As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrWhiteSpace(query) Then Return result
            Dim entries = query.Split(New Char() {"&"c}, StringSplitOptions.RemoveEmptyEntries)
            For Each e In entries
                Dim idx = e.IndexOf("="c)
                If idx > 0 Then
                    result(e.Substring(0, idx)) = e.Substring(idx + 1)
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
        End Sub

        Private Shared Sub OpenUrl(url As String)
            Process.Start(New ProcessStartInfo With {.FileName = url, .UseShellExecute = True})
        End Sub

        Private Sub ApplyUnifiedUiStyle(root As Control)
            For Each c As Control In root.Controls
                If TypeOf c Is Button Then
                    c.Height = 32
                    c.Width = Math.Max(c.Width, 120)
                ElseIf TypeOf c Is TextBox Then
                    c.Height = 26
                    c.Width = Math.Max(c.Width, 280)
                    c.Font = New Drawing.Font("Microsoft YaHei UI", 9.5F)
                ElseIf TypeOf c Is Label Then
                    c.Font = New Drawing.Font("Microsoft YaHei UI", 9.5F)
                End If
                If c.HasChildren Then ApplyUnifiedUiStyle(c)
            Next
        End Sub

        Private Sub ApplyMacLikeTheme(root As Control)
            BackColor = Drawing.Color.FromArgb(245, 246, 248)
            _tabs.Appearance = TabAppearance.Normal
            _tabs.SizeMode = TabSizeMode.Normal
            _tabs.DrawMode = TabDrawMode.Normal
            _tabs.Font = New Drawing.Font("Microsoft YaHei UI", 9.5F)
            _status.BackColor = Drawing.Color.FromArgb(236, 238, 242)
            _status.ForeColor = Drawing.Color.FromArgb(70, 70, 70)
            _notice.BackColor = Drawing.Color.FromArgb(252, 253, 255)
            _notice.BorderStyle = BorderStyle.FixedSingle
            ApplyMacLikeThemeRecursive(root)
        End Sub

        Private Sub ApplyMacLikeThemeRecursive(root As Control)
            For Each c As Control In root.Controls
                If TypeOf c Is Button Then
                    Dim b = DirectCast(c, Button)
                    b.FlatStyle = FlatStyle.Flat
                    b.FlatAppearance.BorderColor = Drawing.Color.FromArgb(198, 201, 206)
                    b.FlatAppearance.MouseOverBackColor = Drawing.Color.FromArgb(238, 243, 252)
                    b.BackColor = Drawing.Color.White
                    b.ForeColor = Drawing.Color.FromArgb(45, 45, 45)
                    b.Margin = New Padding(6)
                ElseIf TypeOf c Is TextBox Then
                    c.BackColor = Drawing.Color.White
                    c.ForeColor = Drawing.Color.FromArgb(45, 45, 45)
                    c.Margin = New Padding(6)
                ElseIf TypeOf c Is Label Then
                    c.ForeColor = Drawing.Color.FromArgb(64, 64, 64)
                    c.Margin = New Padding(4)
                ElseIf TypeOf c Is TabPage Then
                    c.BackColor = Drawing.Color.FromArgb(248, 249, 251)
                ElseIf TypeOf c Is GroupBox Then
                    c.ForeColor = Drawing.Color.FromArgb(60, 60, 60)
                    c.BackColor = Drawing.Color.White
                End If
                If c.HasChildren Then ApplyMacLikeThemeRecursive(c)
            Next
        End Sub

        Private Shared Sub ConfigureCompactGrid(panel As TableLayoutPanel, labelWidth As Integer)
            panel.Padding = New Padding(18, 14, 18, 14)
            panel.AutoScroll = True
            panel.ColumnStyles.Clear()
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, labelWidth))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            panel.GrowStyle = TableLayoutPanelGrowStyle.AddRows
        End Sub

        ' 纵向堆叠容器的统一样式（控制台页使用）。
        Private Shared Sub ConfigureStackPanel(panel As FlowLayoutPanel)
            panel.Padding = New Padding(18, 14, 18, 14)
            panel.WrapContents = False
            panel.AutoScroll = True
        End Sub
    End Class
End Namespace
