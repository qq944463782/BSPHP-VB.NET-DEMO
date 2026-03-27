Public Class frm_login

    Public Version As String '本程序版本号
    Public BSPHP As BSphp '网络验证类

    Private Sub frm_login_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        combMibaoQus.SelectedIndex = 0
        comboBox2.SelectedIndex = 0
        comboBox1.SelectedIndex = 0
        Version = "1.0" ''设置当前软件版本号

        '这里也可以用
        '  Version = My.Application.Info.Version.ToString 版本号需要和网站设置软件的版本对应

        BSPHP = New BSphp() ''实例化            

        '获取软件公告
        labGongGao.Text = BSPHP.AppEn_GetGongGao()

        '根据是否需要使用验证码，调整界面显示
        initImga()
    End Sub
    Private Sub initImga()

        '获取后台验证码要求状态
        Dim info As String = BSPHP.getSetImag("INGES_LOGIN登录|INGES_RE注册|INGES_MACK找回改密|INGES_SAY留言") '
        '根据验证码状态改变界面是否显示验证码,不需要验证码的可以直接不放置这些控件就可以了

        If (Not info.Contains("checked登录")) Then

            label3.Visible = False '
            txtImga.Visible = False '
            picLogin.Visible = False '
            UpdtImgaLogin.Visible = False '
        End If
        If (Not info.Contains("checked注册")) Then

            label11.Visible = False '
            txtRegImga.Visible = False '
            picReg.Visible = False '
            UpdtImgaReg.Visible = False '
        End If
        If (Not info.Contains("checked找回改密")) Then

            '更改密码界面
            label12.Visible = False '
            txtPwdImga.Visible = False '
            picPwdImga.Visible = False '
            UpdtImgaPwd.Visible = False '
            '找回密码界面
            label39.Visible = False '
            txtBacktoImga.Visible = False '
            picBackto.Visible = False '
            UpdtImgaBackto.Visible = False '
        End If
        If (Not info.Contains("checked留言")) Then

            label21.Visible = False '
            textBox10.Visible = False '
            picLiuY.Visible = False '
            UpdtImgaLiuY.Visible = False '
        End If
        If (info.Contains("checked")) Then

            updateImga() '
        End If

    End Sub

    Private Sub updateImga()

        If (BSPHP.updateImga()) Then

            picLogin.Image = BSPHP.AppEn_imga ''登陆界面的验证码图片框
            picReg.Image = BSPHP.AppEn_imga ''注册界面的验证码图片框
            picPwdImga.Image = BSPHP.AppEn_imga ''修改密码界面的验证码图片框
            picBackto.Image = BSPHP.AppEn_imga ''找回密码界面的验证码图片框
            picLiuY.Image = BSPHP.AppEn_imga ''留言界面的验证码图片框
        End If
    End Sub
    Private Sub btnTestLink_Click(sender As Object, e As EventArgs) Handles btnTestLink.Click

        Dim s As String
        s = BSPHP.AppEn_Internet() '
        If (s = "1") Then
            MessageBox.Show("网络连接正常", "测试连接") '
        Else
            MessageBox.Show("网络连接失败,请检查", "测试连接") '
        End If
    End Sub

    Private Sub butTestVersion_Click(sender As Object, e As EventArgs) Handles butTestVersion.Click

        Dim info As String = BSPHP.AppEn_v() '
        If (info <> Version) Then ' 不等于当前版本号
            MessageBox.Show("发现新版本,请打开该页下载最新版本:" + BSPHP.AppEn_in_url(), "版本检测") '
        Else
            MessageBox.Show("当前版本已经是最新版本", "版本检测") '
        End If
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click

        txtUser.Text = txtUser.Text.Trim() '
        txtPwd.Text = txtPwd.Text.Trim() '
        If (txtUser.Text = "") Then

            MessageBox.Show("请填写账号，没有账号请点击<注册账号>创建账号。") '
            Exit Sub

        ElseIf (txtPwd.Text = "") Then

            MessageBox.Show("请填写密码。") '
            Exit Sub
        End If
        Dim srt As String

        srt = BSPHP.AppEn_LogIn(txtUser.Text, txtPwd.Text, txtImga.Text) '

        ' MsgBox(BSPHP.AppEn_LogIn2(txtUser.Text, txtPwd.Text, txtImga.Text))

        If (srt.Contains("1011")) Then

            MessageBox.Show("恭喜你登录成功") '
            frm_main.Show()
            ' Me.Hide() '

        ElseIf (srt.Contains("9908")) Then

            MessageBox.Show("你的帐号已经到期了哦,请续费") '
            tabControl1.SelectTab(2) '
            txtUpUser.Text = txtUser.Text '
            txtUpPwd.Text = txtPwd.Text '

        Else

            MessageBox.Show(srt) '
        End If
        updateImga() '
    End Sub

    Private Sub btnBuyCard_Click(sender As Object, e As EventArgs) Handles btnBuyCard.Click

        '用IE打开购卡网址
        System.Diagnostics.Process.Start("IEXPLORE.EXE", "http://demo.bsphp.com/index.php?m=webapi&c=salecard_gencard&a=index&daihao=8888888") '
    End Sub


    '点击换一张或者点击验证码显示框，实现换一张
    Private Sub UpdtImgaLogin_Click(sender As Object, e As EventArgs) Handles UpdtImgaLogin.Click
        updateImga() '
    End Sub

    '注册用户名完成输入后,检验用户名是否已经存在
    Private Sub txtRegUser_Leave(sender As Object, e As EventArgs) Handles txtRegUser.Leave 'check

        Dim sult As String = BSPHP.AppEn_TestUser(txtRegUser.Text.Trim()) '
        If (sult <> "1001") Then
            MessageBox.Show("此账号已经存在，请换个再试。", "用户名错误") '
        End If
    End Sub

    Private Sub button2_Click(sender As Object, e As EventArgs) Handles button2.Click

        '本地检查，如果不想写代码，可以交给服务器去检查，但是需要网路访问时间
        If (txtRegUser.Text.Trim() <> "") Then

            If (txtRegPwdA.Text.Trim() <> "") Then

                If (txtRegPwdB.Text.Trim() = txtRegPwdA.Text.Trim()) Then

                    Dim sult As String = BSPHP.AppEn_registration(txtRegUser.Text.Trim(), txtRegPwdA.Text.Trim(), txtRegPwdB.Text.Trim(), combMibaoQus.Text.Trim(), txtMibaoAns.Text.Trim(), txtQQ.Text.Trim(), txtEmail.Text.Trim(), txtTEL.Text.Trim(), txtRegImga.Text.Trim()) '
                    BSPHP.ShowSult(sult, "注册结果") '
                Else
                    MessageBox.Show("两次输入的密码不一致。", "注册信息出错") '
                End If
            Else
                MessageBox.Show("密码不能为空。", "注册信息出错") '
            End If
        Else
            MessageBox.Show("用户名不能为空。", "注册信息出错") '
        End If
        updateImga() '
    End Sub

    Private Sub txtUpUser_Leave(sender As Object, e As EventArgs) Handles txtUpUser.Leave

        Dim sult As String = BSPHP.AppEn_TestUser(txtUpUser.Text.Trim()) '
        If (sult = "1001") Then
            MessageBox.Show("此账号不存在，请确认后再试。", "用户名错误") '
        End If
    End Sub

    Private Sub btnRecharge_Click(sender As Object, e As EventArgs) Handles btnRecharge.Click

        Dim userset As String '
        If (txtUpPwd.Text.Trim() <> "") Then
            userset = "1" '
        Else
            userset = "0" '
            Dim info As String = BSPHP.AppEn_vipchong(txtUpUser.Text, txtUpPwd.Text, userset, txtCardID.Text, txtCardPwd.Text) '
            BSPHP.ShowSult(info, "充值结果") '
        End If
    End Sub

    Private Sub button10_Click(sender As Object, e As EventArgs) Handles button10.Click

        Dim info As String = BSPHP.AppEn_Unbundling(textBox9.Text, textBox8.Text) '
        BSPHP.ShowSult(info, "解绑结果") '
    End Sub

    Private Sub btnChgPwd_Click(sender As Object, e As EventArgs) Handles btnChgPwd.Click

        Dim info As String = BSPHP.AppEn_password_S(txtChgUser.Text, txtOldPwd.Text, txtNewPwdA.Text, txtNewPwdB.Text, txtPwdImga.Text) '
        BSPHP.ShowSult(info, "密码修改") '
        updateImga() '
    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click

        Dim info As String = BSPHP.AppEn_backto(textBox1.Text, textBox4.Text, textBox3.Text, comboBox2.Text, textBox2.Text, txtBacktoImga.Text) '
        BSPHP.ShowSult(info, "密码找回") '
        updateImga() '
    End Sub

    Private Sub button12_Click(sender As Object, e As EventArgs) Handles button12.Click

        Dim info As String = BSPHP.AppEn_Feedback(textBox12.Text, comboBox1.Text, textBox11.Text, textBox13.Text, textBox10.Text, textBox14.Text, textBox15.Text) '
        BSPHP.ShowSult(info, "留言") '
        updateImga() '
    End Sub

    Private Sub picPwdImga_Click(sender As Object, e As EventArgs) Handles picPwdImga.Click

        updateImga() '
    End Sub

    Private Sub UpdtImgaPwd_Click(sender As Object, e As EventArgs) Handles UpdtImgaPwd.Click

        updateImga() '
    End Sub

    Private Sub picBackto_Click(sender As Object, e As EventArgs) Handles picBackto.Click
        updateImga() '
    End Sub

    Private Sub UpdtImgaBackto_Click(sender As Object, e As EventArgs) Handles UpdtImgaBackto.Click

        updateImga() '
    End Sub

End Class
