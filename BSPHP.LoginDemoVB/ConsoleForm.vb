Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Windows.Forms

Namespace BSPHP.LoginDemoVB
    Public Class ConsoleForm
        Inherits Form

        Private ReadOnly _client As BSPHPClient
        Private ReadOnly _panel As New FlowLayoutPanel() With {.Dock = DockStyle.Fill, .AutoScroll = True, .FlowDirection = FlowDirection.TopDown, .WrapContents = False}

        Public Sub New(client As BSPHPClient)
            _client = client
            Text = "BSPHP 控制台窗口"
            Width = 900
            Height = 600
            _panel.Padding = New Padding(12)
            Controls.Add(_panel)
            BuildButtons()
        End Sub

        Private Sub BuildButtons()
            _panel.Controls.Add(MakeConsoleGroup("公用接口", New String() {
                "服务器时间::date.in", "预设URL::url.in", "Web地址::weburl.in", "全局配置::globalinfo.in",
                "验证码开关(全部)::getsetimag.in", "逻辑值A::logica.in", "逻辑值B::logicb.in", "逻辑值A内容::logicinfoa.in", "逻辑值B内容::logicinfob.in"
            }))

            _panel.Controls.Add(MakeConsoleGroup("自定义配置模型", New String() {
                "软件配置::appcustom.in?info=myapp", "VIP配置::appcustom.in?info=myvip", "登录配置::appcustom.in?info=mylogin"
            }))

            _panel.Controls.Add(MakeConsoleGroup("通用接口", New String() {
                "获取版本::v.in", "获取软件描述::miao.in"
            }))

            _panel.Controls.Add(MakeConsoleGroup("登录模式接口", New String() {
                "取用户信息::getuserinfo.lg", "检测到期::vipdate.lg", "用户特征Key::userkey.lg", "心跳包更新::timeout.lg", "注销登录::cancellation.lg"
            }))

            Dim renewBar As New FlowLayoutPanel() With {.AutoSize = True, .FlowDirection = FlowDirection.LeftToRight, .WrapContents = True}
            renewBar.Controls.Add(New Label() With {.Text = "续费订阅推广", .AutoSize = True})
            Dim bRenew As New Button() With {.Text = "续费订阅(直接)", .Width = 160, .Height = 32}
            AddHandler bRenew.Click, Sub() OpenUrl("https://demo.bsphp.com/index.php?m=webapi&c=salecard_renew&a=index&daihao=8888888")
            Dim bCard As New Button() With {.Text = "购买充值卡", .Width = 130, .Height = 32}
            AddHandler bCard.Click, Sub() OpenUrl("https://demo.bsphp.com/index.php?m=webapi&c=salecard_gencard&a=index&daihao=8888888")
            Dim bStock As New Button() With {.Text = "购买库存卡", .Width = 130, .Height = 32}
            AddHandler bStock.Click, Sub() OpenUrl("https://demo.bsphp.com/index.php?m=webapi&c=salecard_salecard&a=index&daihao=8888888")
            renewBar.Controls.Add(bRenew)
            renewBar.Controls.Add(bCard)
            renewBar.Controls.Add(bStock)
            _panel.Controls.Add(renewBar)
        End Sub

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
                Dim b As New Button() With {.Text = caption, .Width = 180, .Height = 32}
                AddHandler b.Click, Async Sub()
                                        Dim api As String = route
                                        Dim extra As Dictionary(Of String, String) = Nothing
                                        Dim q = route.IndexOf("?"c)
                                        If q > 0 Then
                                            api = route.Substring(0, q)
                                            extra = ParseQuery(route.Substring(q + 1))
                                        End If
                                        Dim r = Await _client.SendAsync(api, extra)
                                        MessageBox.Show("code=" & If(r.Code.HasValue, r.Code.Value.ToString(), "nil") & Environment.NewLine & r.DataText, caption)
                                    End Sub
                p.Controls.Add(b)
            Next
            g.Controls.Add(p)
            Return g
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

        Private Shared Sub OpenUrl(url As String)
            Process.Start(New ProcessStartInfo With {.FileName = url, .UseShellExecute = True})
        End Sub
    End Class
End Namespace
