Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace BSPHP.CardDemoVB
    Public Class ControlPanelForm
        Inherits Form

        Private ReadOnly _client As BSPHP.LoginDemoVB.BSPHPClient
        Private ReadOnly _cardId As String
        Private ReadOnly _output As New TextBox() With {.Multiline = True, .Dock = DockStyle.Fill, .ScrollBars = ScrollBars.Both}
        Private ReadOnly _header As New Label() With {.AutoSize = True}

        Public Sub New(client As BSPHP.LoginDemoVB.BSPHPClient, cardId As String)
            _client = client
            _cardId = cardId
            Text = "卡密主控制面板 - " & cardId
            Width = 900
            Height = 600

            _header.Text = "登录卡号: " & _cardId & " | 到期时间: 加载中..."

            Dim root As New TableLayoutPanel() With {.Dock = DockStyle.Top, .Height = 220, .ColumnCount = 1, .RowCount = 4}
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 60))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 90))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 40))
            root.Controls.Add(_header, 0, 0)

            Dim groupCommon As New GroupBox() With {.Text = "公用接口", .Dock = DockStyle.Fill}
            Dim commonPanel As New FlowLayoutPanel() With {.Dock = DockStyle.Fill, .Padding = New Padding(8), .AutoScroll = True}
            AddApiButton(commonPanel, "globalinfo.in")
            AddApiButton(commonPanel, "v.in")
            groupCommon.Controls.Add(commonPanel)
            root.Controls.Add(groupCommon, 0, 1)

            Dim groupCard As New GroupBox() With {.Text = "卡密接口", .Dock = DockStyle.Fill}
            Dim cardPanel As New FlowLayoutPanel() With {.Dock = DockStyle.Fill, .Padding = New Padding(8), .AutoScroll = True}
            AddApiButton(cardPanel, "getdate.ic")
            AddApiButton(cardPanel, "getlkinfo.ic")
            AddApiButton(cardPanel, "getdata.ic", New Dictionary(Of String, String) From {{"key", "demo"}})
            AddApiButton(cardPanel, "getinfo.ic", New Dictionary(Of String, String) From {{"ic_carid", cardId}, {"ic_pwd", ""}, {"info", "icid"}})
            AddApiButton(cardPanel, "setcaron.ic", New Dictionary(Of String, String) From {{"key", BSPHP.LoginDemoVB.BSPHPClient.MachineCode()}, {"icid", cardId}, {"icpwd", ""}})
            AddApiButton(cardPanel, "setcarnot.ic", New Dictionary(Of String, String) From {{"icid", cardId}, {"icpwd", ""}})
            AddApiButton(cardPanel, "remotecancellation.ic", New Dictionary(Of String, String) From {{"icid", cardId}, {"icpwd", ""}, {"type", "0"}})
            AddApiButton(cardPanel, "socard.in", New Dictionary(Of String, String) From {{"cardid", cardId}})
            AddApiButton(cardPanel, "timeout.ic")
            AddApiButton(cardPanel, "cancellation.ic")
            groupCard.Controls.Add(cardPanel)
            root.Controls.Add(groupCard, 0, 2)

            Dim groupRenew As New GroupBox() With {.Text = "续费入口", .Dock = DockStyle.Fill}
            Dim renewPanel As New FlowLayoutPanel() With {.Dock = DockStyle.Fill, .Padding = New Padding(8), .AutoScroll = True}
            AddLinkButton(renewPanel, "续费购买订阅", "https://demo.bsphp.com/index.php?m=webapi&c=salecard_renew&a=index&daihao=8888888")
            AddLinkButton(renewPanel, "购买充值卡", "https://demo.bsphp.com/index.php?m=webapi&c=salecard_gencard&a=index&daihao=8888888")
            AddLinkButton(renewPanel, "购买库存卡", "https://demo.bsphp.com/index.php?m=webapi&c=salecard_salecard&a=index&daihao=8888888")
            groupRenew.Controls.Add(renewPanel)
            root.Controls.Add(groupRenew, 0, 3)

            Controls.Add(root)
            Controls.Add(_output)
            ApplyMacLikeTheme(Me)

            AddHandler Shown, Async Sub() Await RefreshHeaderAsync()
        End Sub

        Private Sub AddApiButton(host As FlowLayoutPanel, api As String, Optional payload As Dictionary(Of String, String) = Nothing)
            Dim b As New Button() With {.Text = api, .Width = 125, .Height = 30, .Margin = New Padding(6)}
            AddHandler b.Click, Async Sub() Await RunApiAsync(api, payload)
            host.Controls.Add(b)
        End Sub

        Private Async Function RunApiAsync(api As String, payload As Dictionary(Of String, String)) As Task
            Dim p = If(payload, New Dictionary(Of String, String)())
            If api = "getdata.ic" AndAlso Not p.ContainsKey("key") Then p("key") = _cardId
            Dim r = Await _client.SendAsync(api, p)
            _output.Text = "api=" & api & Environment.NewLine &
                "code=" & If(r.Code.HasValue, r.Code.Value.ToString(), "nil") & Environment.NewLine &
                r.DataText
            If api = "getdate.ic" Then
                _header.Text = "登录卡号: " & _cardId & " | 到期时间: " & If(String.IsNullOrWhiteSpace(r.DataText), "-", r.DataText)
            End If
        End Function

        Private Sub AddLinkButton(host As FlowLayoutPanel, title As String, url As String)
            Dim b As New Button() With {.Text = title, .Width = 125, .Height = 30, .Margin = New Padding(6)}
            AddHandler b.Click, Sub() Process.Start(New ProcessStartInfo With {.FileName = url, .UseShellExecute = True})
            host.Controls.Add(b)
        End Sub

        Private Async Function RefreshHeaderAsync() As Task
            Dim d = Await _client.SendAsync("getdate.ic")
            _header.Text = "登录卡号: " & _cardId & " | 到期时间: " & If(String.IsNullOrWhiteSpace(d.DataText), "-", d.DataText)
        End Function

        Private Sub ApplyMacLikeTheme(root As Control)
            BackColor = Drawing.Color.FromArgb(245, 246, 248)
            _output.BackColor = Drawing.Color.White
            _output.Font = New Drawing.Font("Microsoft YaHei UI", 9.5F)
            _header.Font = New Drawing.Font("Microsoft YaHei UI Semibold", 10.0F)
            _header.ForeColor = Drawing.Color.FromArgb(55, 55, 55)
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
                ElseIf TypeOf c Is GroupBox Then
                    c.BackColor = Drawing.Color.White
                    c.ForeColor = Drawing.Color.FromArgb(60, 60, 60)
                    c.Font = New Drawing.Font("Microsoft YaHei UI", 9.5F)
                End If
                If c.HasChildren Then ApplyMacLikeThemeRecursive(c)
            Next
        End Sub
    End Class
End Namespace
