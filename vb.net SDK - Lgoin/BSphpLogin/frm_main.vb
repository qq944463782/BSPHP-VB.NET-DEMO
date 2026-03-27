Public Class frm_main
    Private Sub frm_main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        label1.Text = frm_login.BSPHP.AppEn_GetMiaoShu()
    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        If (frm_login.BSPHP.AppEn_Internet() = "1") Then
            MessageBox.Show("网络连接正常", "测试连接")
        Else
            MessageBox.Show("网络连接失败,请检查", "测试连接")
        End If
    End Sub

    Private Sub button2_Click(sender As Object, e As EventArgs) Handles button2.Click
        Dim info As String = frm_login.BSPHP.AppEn_GetVipDieDate()
        frm_login.BSPHP.ShowSult(info, "VIP到期信息")
    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles button5.Click
        Dim info As String = frm_login.BSPHP.AppEn_GetUserKey()
        frm_login.BSPHP.ShowSult(info, "绑定特征信息")
    End Sub

    Private Sub button3_Click(sender As Object, e As EventArgs) Handles button3.Click
        Dim info As String = frm_login.BSPHP.AppEn_GetUserInfo("#激活时间>UserReDate#激活时Ip>UserReIp#用户状态>UserIsLock#登录时间>UserLogInDate#登录Ip>UserLogInIp#到期时>UserVipDate#绑定特征>UserKey#用户分组名称>Class_Nane#用户分组别名>Class_Mark##用户QQ>UserQQ#用户邮箱>UserMAIL#购卡折扣>UserPayZhe#是否代理代理>1>UserTreasury#电话>UserMobile#帐号金额>UserRMB#帐号积分>UserPoint#密保问题>Usermibao_wenti#VIP到期时间>UserVipDate#vip是否到期没有到期返回1到期返回2>UserVipWhether")
        frm_login.BSPHP.ShowSult(info, "用户信息")
    End Sub

    Private Sub button4_Click(sender As Object, e As EventArgs) Handles button4.Click
        '你好这里的数据说明防破解重要防线,就看你怎么用
        '我这个数据如果帐号过期了取不到的哦
        '还有绑定特征不对也是取不到的哦
        '就是说解除绑定后也取不到
        '也就是1.帐号必须登录 2.帐号没过期 3.绑定特征要一致才可以获取到！ 就看你怎么用
        '开始取数据 5033=用户已经到期,无法获取重要的验证数据 ；5034=非绑定机器
        Dim info As String = frm_login.BSPHP.AppEn_GetVerification()
        frm_login.BSPHP.ShowSult(info, "验证数据信息")
    End Sub

    Private Sub button6_Click(sender As Object, e As EventArgs) Handles button6.Click
        Dim info As String = frm_login.BSPHP.AppEn_timeout()
        frm_login.BSPHP.ShowSult(info, "登陆状态信息")
    End Sub

    Private Sub button7_Click(sender As Object, e As EventArgs) Handles button7.Click
        Dim Str As String = frm_login.BSPHP.AppEn_globalinfo("信息字段：#GLOBAL_WEBURL->取Web浏览地址#GLOBAL_URL->预设URL地址#GLOBAL_LOGICA->逻辑提示A#GLOBAL_LOGICINFOA->逻辑提示A内容#GLOBAL_LOGICB->逻辑提示B#GLOBAL_LOGICINFOB->逻辑提示B内容")

        MessageBox.Show("用户信息:" + Str)

        MessageBox.Show("详细参数请参数对应API帮助文档")
    End Sub
End Class