Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace BSPHP.CardDemoVB
    Public Class MainForm
        Inherits Form

        Private ReadOnly _client As BSPHP.LoginDemoVB.BSPHPClient = New BSPHP.LoginDemoVB.BSPHPClient(New BSPHPConfig())
        Private ReadOnly _notice As New TextBox()
        Private ReadOnly _tabs As New TabControl()
        Private ReadOnly _status As New Label()

        Private ReadOnly _cardId As New TextBox()
        Private ReadOnly _cardPwd As New TextBox()
        Private ReadOnly _machine As New TextBox() With {.Text = BSPHP.LoginDemoVB.BSPHPClient.MachineCode()}
        Private ReadOnly _renewCardNo As New TextBox()
        Private ReadOnly _renewCardPwd As New TextBox()
        Private ReadOnly _codeImage As New PictureBox() With {.Width = 140, .Height = 48, .SizeMode = PictureBoxSizeMode.Zoom, .BorderStyle = BorderStyle.FixedSingle}
        Private ReadOnly _loggedCard As New Label() With {.Text = "当前卡号: -", .AutoSize = True}
        Private ReadOnly _expiry As New Label() With {.Text = "到期时间: -", .AutoSize = True}

        Public Sub New()
            Text = "BSPHP 卡密模式 VB.NET Demo"
            Width = 900
            Height = 660
            StartPosition = FormStartPosition.CenterScreen

            _notice.Multiline = True
            _notice.Dock = DockStyle.Top
            _notice.Height = 72
            _notice.ReadOnly = True
            _notice.Text = "初始化中..."
            Controls.Add(_notice)

            Dim tabHost As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(0, 50, 0, 0)}
            _tabs.Dock = DockStyle.Fill
            tabHost.Controls.Add(_tabs)
            Controls.Add(tabHost)
            BuildTabs()
            ApplyUnifiedUiStyle(Me)
            ApplyMacLikeTheme(Me)

            _status.Dock = DockStyle.Bottom
            _status.Height = 24
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
                _status.Text = ex.Message
            End Try
        End Function

        Private Sub BuildTabs()
            _tabs.TabPages.Add(MakeCardLoginTab())
            _tabs.TabPages.Add(MakeMachineTab())
        End Sub

        Private Function MakeCardLoginTab() As TabPage
            Dim tab As New TabPage("制作卡密登录模式")
            Dim panel As New FlowLayoutPanel() With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.TopDown, .WrapContents = False, .AutoScroll = True}
            ConfigureStackPanel(panel)
            panel.Controls.Add(_loggedCard)
            panel.Controls.Add(_expiry)
            panel.Controls.Add(New Label() With {.Text = "卡串"})
            panel.Controls.Add(_cardId)
            panel.Controls.Add(New Label() With {.Text = "密码(可空)"})
            panel.Controls.Add(_cardPwd)
            panel.Controls.Add(New Label() With {.Text = "验证码图"})
            panel.Controls.Add(_codeImage)
            Dim btnRefresh As New Button() With {.Text = "刷新验证码", .Width = 180}
            AddHandler btnRefresh.Click, Sub() RefreshCaptcha()
            panel.Controls.Add(btnRefresh)

            Dim btnLogin As New Button() With {.Text = "验证使用", .Width = 180}
            AddHandler btnLogin.Click, Async Sub() Await VerifyCardAsync()
            panel.Controls.Add(btnLogin)
            Dim btnPrefill As New Button() With {.Text = "预填演示参数", .Width = 180}
            AddHandler btnPrefill.Click, Sub()
                                            _cardId.Text = "demo_card"
                                            _cardPwd.Text = ""
                                        End Sub
            panel.Controls.Add(btnPrefill)

            Dim btnRenew As New Button() With {.Text = "续费购买订阅", .Width = 180}
            AddHandler btnRenew.Click, Sub() OpenUrl("https://demo.bsphp.com/index.php?m=webapi&c=salecard_renew&a=index&daihao=8888888")
            panel.Controls.Add(btnRenew)

            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Function MakeMachineTab() As TabPage
            Dim tab As New TabPage("机器码账号模式")
            Dim panel As New FlowLayoutPanel() With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.TopDown, .WrapContents = False, .AutoScroll = True}
            ConfigureStackPanel(panel)
            panel.Controls.Add(New Label() With {.Text = "机器码账号"})
            panel.Controls.Add(_machine)

            Dim btnVerify As New Button() With {.Text = "机器码验证使用", .Width = 180}
            AddHandler btnVerify.Click, Async Sub() Await VerifyMachineAsync()
            panel.Controls.Add(btnVerify)
            Dim btnPrefill As New Button() With {.Text = "预填演示参数", .Width = 180}
            AddHandler btnPrefill.Click, Sub()
                                            _renewCardNo.Text = "demo_ka"
                                            _renewCardPwd.Text = "123456"
                                        End Sub
            panel.Controls.Add(btnPrefill)

            panel.Controls.Add(New Label() With {.Text = "充值卡号"})
            panel.Controls.Add(_renewCardNo)
            panel.Controls.Add(New Label() With {.Text = "充值密码"})
            panel.Controls.Add(_renewCardPwd)

            Dim btnRecharge As New Button() With {.Text = "确认充值(chong.ic)", .Width = 180}
            AddHandler btnRecharge.Click, Async Sub() Await RechargeAsync()
            panel.Controls.Add(btnRecharge)

            tab.Controls.Add(panel)
            Return tab
        End Function

        Private Async Function VerifyCardAsync() As Task
            Dim r = Await _client.SendAsync("login.ic", New Dictionary(Of String, String) From {
                {"icid", _cardId.Text.Trim()},
                {"icpwd", _cardPwd.Text.Trim()},
                {"key", BSPHP.LoginDemoVB.BSPHPClient.MachineCode()},
                {"maxoror", BSPHP.LoginDemoVB.BSPHPClient.MachineCode()}
            })
            MessageBox.Show("code=" & If(r.Code.HasValue, r.Code.Value.ToString(), "nil") & Environment.NewLine & r.DataText, "卡密登录")
            If r.Code.HasValue AndAlso (r.Code.Value = 1081 OrElse r.Code.Value = 1011) Then
                _loggedCard.Text = "当前卡号: " & _cardId.Text.Trim()
                Dim exp = Await _client.SendAsync("getdate.ic")
                _expiry.Text = "到期时间: " & If(String.IsNullOrWhiteSpace(exp.DataText), "-", exp.DataText)
                Dim panel As New ControlPanelForm(_client, _cardId.Text.Trim())
                panel.Show()
            End If
        End Function

        Private Async Function VerifyMachineAsync() As Task
            Dim id = _machine.Text.Trim()
            Dim machine = BSPHP.LoginDemoVB.BSPHPClient.MachineCode()
            Dim r1 = Await _client.SendAsync("AddCardFeatures.key.ic", New Dictionary(Of String, String) From {
                {"carid", id},
                {"key", machine},
                {"maxoror", machine}
            })
            Dim r2 = Await _client.SendAsync("login.ic", New Dictionary(Of String, String) From {
                {"icid", id},
                {"icpwd", ""},
                {"key", machine},
                {"maxoror", machine}
            })
            MessageBox.Show("AddCardFeatures: " & r1.DataText & Environment.NewLine & "login.ic: " & r2.DataText, "机器码模式")
            If r2.Code.HasValue AndAlso (r2.Code.Value = 1081 OrElse r2.Code.Value = 1011) Then
                _loggedCard.Text = "当前卡号: " & id
                Dim exp = Await _client.SendAsync("getdate.ic")
                _expiry.Text = "到期时间: " & If(String.IsNullOrWhiteSpace(exp.DataText), "-", exp.DataText)
                Dim panel As New ControlPanelForm(_client, id)
                panel.Show()
            End If
        End Function

        Private Async Function RechargeAsync() As Task
            Dim r = Await _client.SendAsync("chong.ic", New Dictionary(Of String, String) From {
                {"icid", _machine.Text.Trim()},
                {"ka", _renewCardNo.Text.Trim()},
                {"pwd", _renewCardPwd.Text.Trim()}
            })
            MessageBox.Show("code=" & If(r.Code.HasValue, r.Code.Value.ToString(), "nil") & Environment.NewLine & r.DataText, "充值结果")
        End Function

        Private Shared Sub OpenUrl(url As String)
            Process.Start(New ProcessStartInfo With {.FileName = url, .UseShellExecute = True})
        End Sub

        Private Sub RefreshCaptcha()
            Try
                Dim u = _client.GetCodeImageUrl()
                If String.IsNullOrWhiteSpace(u) Then Return
                Dim ts As Long = CLng((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalMilliseconds)
                _codeImage.Load(u & "&_=" & ts.ToString())
            Catch
            End Try
        End Sub

        Private Sub ApplyUnifiedUiStyle(root As Control)
            For Each c As Control In root.Controls
                If TypeOf c Is Button Then
                    c.Height = 32
                    c.Width = Math.Max(c.Width, 120)
                ElseIf TypeOf c Is TextBox Then
                    c.Height = 26
                    c.Width = Math.Max(c.Width, 300)
                    c.Font = New Drawing.Font("Microsoft YaHei UI", 9.5F)
                ElseIf TypeOf c Is Label Then
                    c.Font = New Drawing.Font("Microsoft YaHei UI", 9.5F)
                End If
                If c.HasChildren Then ApplyUnifiedUiStyle(c)
            Next
        End Sub

        Private Sub ApplyMacLikeTheme(root As Control)
            BackColor = Drawing.Color.FromArgb(245, 246, 248)
            _tabs.Font = New Drawing.Font("Microsoft YaHei UI", 9.5F)
            _status.BackColor = Drawing.Color.FromArgb(236, 238, 242)
            _status.ForeColor = Drawing.Color.FromArgb(70, 70, 70)
            _notice.BackColor = Drawing.Color.White
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
                End If
                If c.HasChildren Then ApplyMacLikeThemeRecursive(c)
            Next
        End Sub

        Private Shared Sub ConfigureStackPanel(panel As FlowLayoutPanel)
            panel.Padding = New Padding(18, 14, 18, 14)
            panel.WrapContents = False
            panel.AutoScroll = True
        End Sub
    End Class
End Namespace
