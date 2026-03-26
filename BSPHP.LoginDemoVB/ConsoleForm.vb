Imports System
Imports System.Windows.Forms

Namespace BSPHP.LoginDemoVB
    Public Class ConsoleForm
        Inherits Form

        Private ReadOnly _client As BSPHPClient
        Private ReadOnly _panel As New FlowLayoutPanel() With {.Dock = DockStyle.Fill, .AutoScroll = True}

        Public Sub New(client As BSPHPClient)
            _client = client
            Text = "BSPHP 控制台窗口"
            Width = 900
            Height = 600
            Controls.Add(_panel)
            BuildButtons()
        End Sub

        Private Sub BuildButtons()
            Dim apis = New String() {
                "date.in", "url.in", "weburl.in", "globalinfo.in", "appcustom.in",
                "v.in", "miao.in", "getsetimag.in", "logica.in", "logicb.in",
                "logicinfoa.in", "logicinfob.in", "getuserinfo.lg", "vipdate.lg",
                "userkey.lg", "timeout.lg", "cancellation.lg"
            }
            For Each api In apis
                Dim b As New Button() With {.Text = api, .Width = 180, .Height = 36}
                AddHandler b.Click, Async Sub()
                                        Dim r = Await _client.SendAsync(api)
                                        MessageBox.Show("code=" & If(r.Code.HasValue, r.Code.Value.ToString(), "nil") & Environment.NewLine & r.DataText, api)
                                    End Sub
                _panel.Controls.Add(b)
            Next
        End Sub
    End Class
End Namespace
