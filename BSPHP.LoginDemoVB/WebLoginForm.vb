Imports System
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace BSPHP.LoginDemoVB
    ' Web 登录窗口：
    ' 1) 加载 webapi 登录地址
    ' 2) 周期性 heartbeat(timeout.lg) 探测是否已登录（5031/1011/9908）
    ' 3) 登录成功后回调主窗体并自动关闭
    Public Class WebLoginForm
        Inherits Form

        Private ReadOnly _client As BSPHPClient
        Private ReadOnly _onLoginSucceeded As Func(Of Task)
        Private ReadOnly _browser As New WebBrowser() With {.Dock = DockStyle.Fill, .ScriptErrorsSuppressed = True}
        Private ReadOnly _toolBar As New FlowLayoutPanel() With {.Dock = DockStyle.Top, .Height = 40, .FlowDirection = FlowDirection.LeftToRight, .WrapContents = False, .Padding = New Padding(8, 6, 8, 6)}
        Private ReadOnly _btnRefresh As New Button() With {.Text = "刷新", .Width = 90, .Height = 26}
        Private ReadOnly _btnClose As New Button() With {.Text = "关闭", .Width = 90, .Height = 26}
        Private ReadOnly _timer As New Timer() With {.Interval = 1500}

        Public Sub New(client As BSPHPClient, onLoginSucceeded As Func(Of Task))
            _client = client
            _onLoginSucceeded = onLoginSucceeded
            Text = "Web登录"
            Width = 800
            Height = 800
            MinimumSize = New Drawing.Size(720, 640)
            StartPosition = FormStartPosition.CenterParent
            _toolBar.Controls.Add(_btnRefresh)
            _toolBar.Controls.Add(_btnClose)
            Controls.Add(_browser)
            Controls.Add(_toolBar)
            AddHandler _btnRefresh.Click, Sub() _browser.Refresh()
            AddHandler _btnClose.Click, Sub() Close()
            AddHandler Shown, AddressOf OnShownAsync
            AddHandler _timer.Tick, AddressOf OnTimerTickAsync
            AddHandler FormClosed, Sub() _timer.Stop()
        End Sub

        Private Sub OnShownAsync(sender As Object, e As EventArgs)
            Dim url As String = "https://demo.bsphp.com/index.php?m=webapi&c=software_auth&a=index&daihao=8888888&BSphpSeSsL=" & _client.Sessl
            _browser.Navigate(url)
            _timer.Start()
        End Sub

        Private Async Sub OnTimerTickAsync(sender As Object, e As EventArgs)
            Try
                Dim hb = Await _client.HeartbeatAsync()
                If hb.Code.HasValue AndAlso (hb.Code.Value = 5031 OrElse hb.Code.Value = 1011 OrElse hb.Code.Value = 9908) Then
                    _timer.Stop()
                    If _onLoginSucceeded IsNot Nothing Then Await _onLoginSucceeded()
                    Close()
                End If
            Catch
                ' Web 登录轮询异常不打断用户页面，仅保持继续轮询。
            End Try
        End Sub
    End Class
End Namespace
