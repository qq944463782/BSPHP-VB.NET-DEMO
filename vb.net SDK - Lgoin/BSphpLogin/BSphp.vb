Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Runtime.InteropServices '引用非托管类DLL需要的命名空间
Imports System.Drawing '验证码Bitmap类需要的命名空间
Imports System.IO
Imports System.Windows.Forms
Imports System.Net
Imports System.Management
Imports System.Web
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions

Public Class BSphp

    Public Function GetCpuID() As String
        Try
            Dim cpuInfo As String = ""
            Dim cimobject As ManagementClass = New ManagementClass("Win32_Processor")
            Dim moc As ManagementObjectCollection = cimobject.GetInstances()

            For Each mo As ManagementObject In moc
                cpuInfo = mo.Properties("ProcessorId").Value.ToString()
                Exit For
            Next

            moc = Nothing
            cimobject = Nothing
            Return cpuInfo
        Catch
            Return "unknowncpu"
        End Try
    End Function


    '此函数只为模拟值
    Public Function GetMD5() As String
        Try
            Dim exePath As String = Application.ExecutablePath
            Dim bytes As Byte() = File.ReadAllBytes(exePath)
            Using md5 As MD5 = MD5.Create()
                Dim hash As Byte() = md5.ComputeHash(bytes)
                Dim sb As New StringBuilder()
                For Each b As Byte In hash
                    sb.Append(b.ToString("x2"))
                Next
                Return sb.ToString()
            End Using
        Catch
            Return "md5test"
        End Try
    End Function

    '//定义变量
    Private AppEn_KEY As String '//用于存放机器码
    Private AppEn_MD5 As String ' //获取自身文件MD5
    Private AppEn_BSphpSeSsl As String '//连接标记，类似Cookies
    Public AppEn_imga As Bitmap ''//验证码图片。图片验证码是可选项，主要是为了防止恶意注册
    '//验证码的使用与否，需要在管理后台的“系统”->“验证码设置”中可以进行设置

    '//定义状态码查询 字典
    ' Public Dictionary<String, String> StatusCode'        
    Public StatusCode As Hashtable


    ' //5个配置参数
    Private AppEn_ImgaUrl As String '   //验证码路径
    Private AppEn_HOST_KEY As String ''  //通信认证Key
    Private AppEn_ServerKey As String ' //服务端RSA私钥（Base64 DER）
    Private AppEn_ClientKey As String ' //客户端RSA公钥（Base64 DER）
    Private AppEn_Url As String '   //服务器地址

    Public Sub New()



        '测试服务器
        AppEn_ImgaUrl = "https://demo.bsphp.com/index.php?m=coode&sessl=" ';//验证码路径

        '通信认证Key
        AppEn_HOST_KEY = "6600cfcd5ac01b9bb3f2460eb416daa8" ';
        '服务器地址
        AppEn_Url = "https://demo.bsphp.com/AppEn.php?appid=8888888&m=95e87faf2f6e41babddaef60273489e1&lang=0" ';//服务器地址
        '服务器私钥
        AppEn_ServerKey = "MIIEqAIBADANBgkqhkiG9w0BAQEFAASCBJIwggSOAgEAAoH+DEr7H5BhMwRA9ZWXVftcCWHznBdl0gQBu5617qSe9in+uloF1sC64Ybdc8Q0JwQkGQANC5PnMqPLgXXIfnl7/LpnQ/BvghQI5cr/4DEezRKrmQaXgYfXHL3woVw7JIsLpPTGa7Ar9S6SEH8RcPIbZjlPVRZPwV3RgWgox2/4lkXsmopqD+mEtOI/ntvti147nEpK2c7cdtCU5M2hQSlIXsTWvri88RTYJ/CtopBOXarUkNBfpWGImiYGsmbZI+YZ6uU0wSYlq8huu+pkTseUUiymzmv8Rpg3coi7YU+pszvB9wnQ1Rz6Z/B6Z3WN7d6OP7f9w0Q0WvgrsKcEJhMCAwEAAQKB/gHa5t6yiRiL0cm902K0VgVMdNjfZww0cpZ/svDaguqfF8PDhhIMb6dNFOo9d6lTpKbpLQ7MOR2ZPkLBJYqAhsdy0dac2BcHMviKk+afQwirgp3LMt3nQ/0gZMnVA0/Wc+Fm1vK1WUzcxEodAuLKhnv8tg4fGdYSdGVU9KJ0MU1bKQZXv0CAIhJYWsiCa5y5bFO7K+ia+UIVBHcvITQLzlgEm+Z/X6ye5cws4pWbk8+spsBDvweb5jpelbkCYs5C5TRNIWXk7+QxTXTg1vrcsmZRcmpRJq7sOd3faZltNHTIlB3HhWnsf47Bz334j9RtU8iqonbuBmcnYbD3+bvBAn891RGdAl+rVU/sJ2kPXmV4eqJOwJfbi8o1WYDp4GcK0ThjrZ1pmaZMj2WTjb3QX1VUoi+7l3389KzzDn0VXLKXZvGxmLikA1FWuuLUmwfNTxyxtGTBVeZCEaQ2lEJuaDGsK0oLi4Bo8ELfQw6JFK7jlgtTlflcYcul99P9BThDAn8y5TpSQy8/07LCgMMZOgJomYzQUmd14Zn2VQLH1u1Z4v2CPlOzGanDt7mmGZCew7iMSO1P0TrwDIreKzYyERuVvZti/IFHH1+J1hAbvk9SJGmdt46W5lyIp3xjdR2QmiK+hSsc8HF9R+zPaSe9yGA8+FwxLRfo0snGP3MC3aXxAn4n2iyABgejZlkc3EnanfzIqkHygC9gUbkCqa1tEDVZw3+Uv1G1vlJxBftyHuk4ZDmbUu1w+zM41nqiLbRxEE4LR06AKO7Yx0qlm86XOVTN/y9/WcWW1saRzs0IYIZwordhQIV463DYMgLn41B7Cdmu1gZ22TLfWCjpz9HSQosCfwMJu9l9OSzOLjV+CidPVyV3RPiKcrKOrOoPWQMkyTY8XnWP0t82APQ121cW35Mai8GT+NZy3tnFZeStH6cNbmAZ2VSnTfA45zMLHBsL2SBGHCfV9ST8yzk9BifJreIb0UceG9y2XY/k4zXeSQkDFPuOt7IXxv2W14SF9Q+Ou4ECfzfRP1hXPwq2w4YJ8sLmqWJT+3aMDucei5MJEAJNifZWhdW0GIrlKRSbhIgLAunxq+KK+mAPqqWw7Prsa21JbXSe3gugusu5d6ESURvLENRKI+Pp9TgRESsydeLy8VcPKRJ5/Ct7/p6QB3A+7F/iPNE2GagGffG9i7e+OdcToYQ="

        '客户端公钥
        AppEn_ClientKey = "MIIBHjANBgkqhkiG9w0BAQEFAAOCAQsAMIIBBgKB/g26m2hYtESqcKW+95Lr+PfCd4bwHW2Z+mM0/vcKQ5j/ZGMigqkgl3QXCEcsCaw0KFSmqAPtLbrl6p5Sp+ZUSYEYQhSxAajE5qRCd3k0r/MIQQanBaOALkP71/u6U2SZhrTXd05n1wQo6ojMH/xVunBOFOa/Eon/Y5FVh6GiJpwwDkFzTlnecmff7Y+VDqRhZ7vu2CQjApOx23N6DiFEmVZYEb/efyASngoZ+3A/DSB5cwbaYVZ21EhPe/GNcwtUleFHn+d4vb0cvolO3Gyw6ObceOT/Q7E3k8ejIml6vPKzmRdtw0FXGOJTclx1CjShRDfXoUjFGyXHy3sZs9VLAgMBAAE="


        '//AppEn_MoShi = "POST";//发包模式大写 POST  GET  




        '  //获取3个基本值
        AppEn_KEY = GetCpuID() '//获取CPU序列号作为机器码，用户也可以自己找方法获取别的绑定参数
        AppEn_MD5 = GetMD5() '//此函数没有实际编写，仅为模拟。


        '//下面这个在网络不好的时候比较耗时，会影响界面显示速度，此处使用简单调用法进行演示。
        ' //追求完美的作者可以采用其他方式初始化下面2个值，在界面显示后在初始化这2个值。
        AppEn_BSphpSeSsl = AppEn_GetBSphpSeSsL() '


        '// msgbox(AppEn_GetBSphpSeSsL())'
        ' //初始化状态码查询字典
        initStatusDic() '
    End Sub

    '//刷新验证码图片
    Public Function updateImga() As Boolean

        AppEn_imga = getImage()
        If (AppEn_imga IsNot Nothing) Then
            Return True
        Else
            Return False
        End If
    End Function

    '//初始化状态查询字典
    Private Sub initStatusDic()

        StatusCode = New Hashtable
        StatusCode.Add("1000", "用户已经存在。请选择其他！") '
        StatusCode.Add("1001", "你选择用户名可以使用！") '
        StatusCode.Add("1002", "2次密码输入不一致！") '
        StatusCode.Add("1003", "账号长度错误限制 3-15位") '
        StatusCode.Add("1004", "账号格式错误,只能选择 26位字母+数字和_下滑线") '
        StatusCode.Add("1005", "恭喜你注册成功！") '
        StatusCode.Add("1006", "账号注册失败,你要注册账号可能被抢注了.") '
        StatusCode.Add("1007", "密码长度限制 3-15位数") '
        StatusCode.Add("1008", "账号错误") '
        StatusCode.Add("1009", "密码错误") '
        StatusCode.Add("1010", "登陆账号不存在") '
        StatusCode.Add("1011", "登陆成功") '
        StatusCode.Add("1012", "密码错误") '
        StatusCode.Add("1013", "登陆失败") '
        StatusCode.Add("1014", "QQ号错误") '
        StatusCode.Add("1015", "邮箱错误") '
        StatusCode.Add("1016", "手机号码错误") '
        StatusCode.Add("1017", "保存成功") '
        StatusCode.Add("1018", "保存失败") '
        StatusCode.Add("1019", "不能输入空格") '
        StatusCode.Add("1020", "请输入密保问题或者答案") '
        StatusCode.Add("1021", "账号已经被冻结禁止登陆") '
        StatusCode.Add("1022", "记录查询失败") '
        StatusCode.Add("1023", "你的密保信息已经填写") '
        StatusCode.Add("1024", "密保信息添加成功") '
        StatusCode.Add("1025", "密保添加失败") '
        StatusCode.Add("1026", "密码不能包含空格") '
        StatusCode.Add("1027", "请输入旧密码") '
        StatusCode.Add("1028", "旧密码不能和新密码一致") '
        StatusCode.Add("1029", "密码修改失败") '
        StatusCode.Add("1030", "旧密码不正确,请重新输入") '
        StatusCode.Add("1031", "密码修改成功") '
        StatusCode.Add("1032", "表单信息不能为空,请返回填写") '
        StatusCode.Add("1033", "密码已经成功通过密保信息修改") '
        StatusCode.Add("1034", "密保信息错误") '
        StatusCode.Add("1035", "密码找回失败") '
        StatusCode.Add("1036", "验证码不能为空") '
        StatusCode.Add("1037", "验证码正确") '
        StatusCode.Add("1038", "验证码错误") '
        StatusCode.Add("1039", "检测账号不能为空") '
        StatusCode.Add("1040", "账号不能包含空格") '
        StatusCode.Add("1041", "长时空闲超时执行正常") '
        StatusCode.Add("1042", "长时空闲超时执行异常") '
        StatusCode.Add("1043", "账号不存在") '
        StatusCode.Add("1044", "账号已经被冻结禁止登陆") '
        StatusCode.Add("1045", "登陆超时,由于你长时间不停留请重新登陆") '
        StatusCode.Add("1046", "你在别处已经登陆,被迫登出！") '
        StatusCode.Add("1047", "已经登陆") '
        StatusCode.Add("1048", "你需要登陆才可以访问") '
        StatusCode.Add("1049", "没有登录或登录已超时请登陆,在继续你的操作！") '
        StatusCode.Add("1050", "密保不能少于4字符") '
        StatusCode.Add("1051", "密保信息未填写") '
        StatusCode.Add("1052", "充值账号不能为空") '
        StatusCode.Add("1053", "充值卡号不能为空") '
        StatusCode.Add("1054", "充值卡密码不能为空") '
        StatusCode.Add("1055", "充值账号不能包含空格") '
        StatusCode.Add("1056", "充值卡号不能包含空格") '
        StatusCode.Add("1057", "充值卡密码不能包含空格") '
        StatusCode.Add("1058", "充值卡号或充值卡密码错误") '
        StatusCode.Add("1059", "充值的用户账号不存在") '
        StatusCode.Add("1060", "充值卡账号密码错误或者不存在") '
        StatusCode.Add("1061", "用户没有使用过要充值软件,拒绝充值") '
        StatusCode.Add("1062", "激活成功,赶快去使用吧！") '
        StatusCode.Add("1063", "充值失败！") '
        StatusCode.Add("1064", "充值卡已经充值过了") '
        StatusCode.Add("1065", "你留言和反馈我们已经收到,谢谢你的支持") '
        StatusCode.Add("1066", "提交留言失败") '
        StatusCode.Add("1067", "请输入标题") '
        StatusCode.Add("1068", "请输入留言内容") '
        StatusCode.Add("1069", "激活成功,请在次验证就可以使用了！") '
        StatusCode.Add("1070", "添加失败") '
        StatusCode.Add("1071", "糟糕卡串已经存在了") '
        StatusCode.Add("1072", "还没有存在") '
        StatusCode.Add("1073", "卡串不存在或者错误") '
        StatusCode.Add("1074", "你使用激活串已到期作废！") '
        StatusCode.Add("1075", "卡串已经存在,无法激活") '
        StatusCode.Add("1076", "car执行错误") '
        StatusCode.Add("1077", "请检查卡串号或者密码错误！") '
        StatusCode.Add("1078", "你使用权已经被冻结,无法验证！") '
        StatusCode.Add("1079", "验证失败请重新验证,或是否已经登陆") '
        StatusCode.Add("1080", "验证成功") '
        StatusCode.Add("1081", "登录验证成功！") '
        StatusCode.Add("1082", "你的使用期已到,请购卡在使用") '
        StatusCode.Add("1083", "没有查询到用户信息") '
        StatusCode.Add("1084", "该用户不是使用本软件的") '
        StatusCode.Add("1085", "用户已经被冻结") '
        StatusCode.Add("1086", "卡串或者验证串已经被冻结,无法继续。") '
        StatusCode.Add("1087", "帐号已经到期请充值在使用。") '
        StatusCode.Add("1088", "请输入一个邮箱作为你的帐号") '
        StatusCode.Add("1089", "恭喜你注册成功,赶快去你邮箱把你帐号激活吧！") '
        StatusCode.Add("1090", "密保邮箱不能为空！") '
        StatusCode.Add("1091", "密保邮箱格式不正确,请重新输入！") '
        StatusCode.Add("1092", "QQ号不能为空！") '
        StatusCode.Add("1093", "QQ号格式输入不正确,请重新输入！") '
        StatusCode.Add("1094", "你的帐号还没激活,现在已经有一封激活邮件发到你注册邮箱上,赶快去激活吧！") '
        StatusCode.Add("5000", "无法接收到GET数据包") '
        StatusCode.Add("5001", "无法接收到POST数据包") '
        StatusCode.Add("5002", "数据包内出现系统屏蔽字符串") '
        StatusCode.Add("5003", "数据包已经过期,拒绝接收") '
        StatusCode.Add("5004", "接口不存在,连接失败") '
        StatusCode.Add("5005", "软件连接号错误,访问软件不存在") '
        StatusCode.Add("5006", "软件MD5验证失败") '
        StatusCode.Add("5007", "非法的请求,身份验证失败！") '
        StatusCode.Add("5008", "欢迎你首次使用！,请重新登陆.") '
        StatusCode.Add("5009", "绑定特征码,已经有人绑定过了,不能重复绑定,不能登陆") '
        StatusCode.Add("5010", "当前绑定特征已经有人绑定了") '
        StatusCode.Add("5011", "账号注册成功,当前机器特征已经有人绑定") '
        StatusCode.Add("5012", "当前特征已经有人绑定,无法在绑定") '
        StatusCode.Add("5013", "绑定成功！") '
        StatusCode.Add("5014", "绑定失败,请重试！") '
        StatusCode.Add("5015", "绑定特征不能为空") '
        StatusCode.Add("5016", "已经绑定了,不需要在绑定") '
        StatusCode.Add("5017", "使用软件不属于登陆模式") '
        StatusCode.Add("5018", "当前卡串已经到期,无法解除绑定") '
        StatusCode.Add("5019", "解除绑定将到期,无法解除绑定") '
        StatusCode.Add("5020", "解除绑定失败,请重试或者联系相关人员解决") '
        StatusCode.Add("5021", "解除绑定成功,已经扣除对应时间") '
        StatusCode.Add("5022", "当前的卡串已经解除绑定,无须在解除绑定,请直接绑定就可以") '
        StatusCode.Add("5023", "绑定KEY不能为空") '
        StatusCode.Add("5024", "恭喜你！绑定成功") '
        StatusCode.Add("5025", "绑定失败,当前卡串已经绑定key,必须解除绑定才能在绑定") '
        StatusCode.Add("5026", "账号登录超时,由于你长时间没有操作请种新登录") '
        StatusCode.Add("5027", "登录状态更新失败！") '
        StatusCode.Add("5028", "登录状态更新成功！") '
        StatusCode.Add("5029", "帐号没有到期！") '
        StatusCode.Add("5030", "登录帐号已经到期！") '
        StatusCode.Add("5031", "执行正常") '
        StatusCode.Add("5032", "扣点模式之扣点执行失败") '
        StatusCode.Add("9980", "用户VIP到期") '
        StatusCode.Add("BSPHP_940006", "AppID md5参数错误 不能为空！") '
        StatusCode.Add("BSPHP_940007", "AppID md5参数错误！服务器地址不正确 ") '
        StatusCode.Add("BSPHP_940008", "通信认证Key不能为空！") '
        StatusCode.Add("BSPHP_940009", "通信认证Key验证失败！") '
        StatusCode.Add("BSPHP_940010", "BSphpSeSsL连接串不可为空！") '
    End Sub

    '//充值
    Public Function AppEn_vipchong(user As String, password As String, userset As String, ka As String, pwd As String) As String

        Return HttpPost("&api=chong.lg&user=" + user + "&userpwd=" + password + "&userset=" + userset + "&ka=" + ka + "&pwd=" + pwd) '
    End Function

    ' //绑定
    Public Function AppEn_Bindkey(key As String) As String

        Return HttpPost("&api=bang.lg&key=" + key) '
    End Function

    '//心跳包
    Public Function AppEn_timeout() As String

        Return HttpPost("&api=timeout.lg") '
    End Function

    '//解绑
    Public Function AppEn_Unbundling(user As String, password As String) As String

        Return HttpPost("&api=jiekey.lg&user=" + user + "&pwd=" + password) '
    End Function

    '//注册
    Public Function AppEn_registration(user As String, pwd As String, pwdb As String, ques As String, ans As String, qq As String, mail As String, mobile As String, img As String) As String

        Return HttpPost("&api=registration.lg&user=" + user + "&pwd=" + pwd + "&pwdb=" + pwdb + "&mibao_wenti=" + ques + "&mibao_daan=" + ans + "&qq=" + qq + "&mail=" + mail + "&mobile=" + mobile + "&coode=" + img) '
    End Function

    '//修改密码
    Public Function AppEn_password_S(user As String, usepwd As String, pwda As String, pwdb As String, img As String) As String

        Return HttpPost("&api=password.lg&user=" + user + "&pwd=" + usepwd + "&pwda=" + pwda + "&pwdb=" + pwdb + "&coode=" + img) '

    End Function
    '//密码找回
    Public Function AppEn_backto(user As String, pwda As String, pwdb As String, quset As String, ans As String, img As String) As String

        Return HttpPost("&api=backto.lg&user=" + user + "&pwd=" + pwda + "&pwdb=" + pwdb + "&mibao=" + quset + "&daan=" + ans + "&coode=" + img) '
    End Function

    '//取逻辑判断A 内容
    Public Function AppEn_GetLogicText_A() As String

        Return HttpPost("&api=logicinfoa.in") '
    End Function

    ' //取逻辑判断B 内容
    Public Function AppEn_GetLogicText_B() As String

        Return HttpPost("&api=logicinfob.in") '
    End Function

    ' //取逻辑判断A
    Public Function AppEn_GetLogic_A() As String

        Return HttpPost("&api=logica.in") '
    End Function

    '//取逻辑判断B
    Public Function AppEn_GetLogic_B() As String

        Return HttpPost("&api=logicb.in") '
    End Function

    ' //取软件配置信息段
    Public Function AppEn_globalinfo(typ As String) As String


        Return HttpPost("&api=globalinfo.in&info=" + typ) '
    End Function

    ' //取验证信息
    Public Function AppEn_GetVerification() As String

        Return HttpPost("&api=lgkey.lg&key=" + AppEn_KEY) '
    End Function

    ' //留言  标题'类型'联系方式'内容'验证码,账号，密码
    Public Function AppEn_Feedback(title As String, type As String, QQ As String, text As String, img As String, user As String, pwd As String) As String

        Return HttpPost("&api=liuyan.in&table=" + title + "&leix=" + type + "&qq=" + QQ + "&txt=" + text + "&coode=" + img + "&user=" + user + "&pwd=" + pwd) '
    End Function

    ' //登录
    Public Function AppEn_LogIn(User As String, pwd As String, imga As String) As String

        If (imga = "") Then

            imga = "null" '
        End If
        Return HttpPost("&api=login.lg&user=" + User + "&pwd=" + pwd + "&key=" + AppEn_KEY + "&coode=" + imga + "&maxoror=" + AppEn_KEY) '

    End Function


    '//判断是否能与服务器进行连接
    Public Function AppEn_Internet() As String

        Return HttpPost("&api=internet.in") '
    End Function

    ' //取用户到期时间
    Public Function AppEn_GetVipDieDate() As String

        Return HttpPost("&api=vipdate.lg") '
    End Function

    '//取得登陆状态信息
    Public Function AppEn_GetLoginInfo() As String

        Return HttpPost("&api=lginfo.lg&key=" + AppEn_KEY) '
    End Function

    '//取登陆状态 
    Public Function AppEn_lginfo() As String

        Return HttpPost("&api=lginfo.lg") '
    End Function

    '//取用户绑定特征key
    Public Function AppEn_GetUserKey() As String

        Return HttpPost("&api=userkey.lg") '
    End Function

    ' //完善用户资料 QQ mail mobile
    Public Function AppEn_AddUserInfo(qq As String, mail As String, mobile As String) As String

        Return HttpPost("&api=Perfect.lg&qq=" + qq + "&mail=" + mail + "&mobile=" + mobile) '
    End Function

    ' //取用户信息  QQ mail mobile
    Public Function AppEn_GetUserInfo(info As String) As String

        Return HttpPost("&api=getuserinfo.lg&info=" + info) '
    End Function

    '//BSphpSeSsL 获取
    Public Function AppEn_GetBSphpSeSsL() As String
        Return HttpPost("&api=BSphpSeSsL.in")
    End Function

    ' //获取公告
    Public Function AppEn_GetGongGao() As String
        Return HttpPost("&api=gg.in") '
    End Function

    '  //取软件描述介绍
    Public Function AppEn_GetMiaoShu() As String
        Return HttpPost("&api=miao.in") '
    End Function

    ' //获取软件的最新版本
    Public Function AppEn_v() As String
        Return HttpPost("&api=v.in") '
    End Function

    '//获取软件的升级地址
    Public Function AppEn_in_url() As String

        Return HttpPost("&api=url.in") '
    End Function

    ' //检测账号是否已经存在
    Public Function AppEn_TestUser(user As String) As String

        Return HttpPost("&api=GetPleaseregister.lg&user=" + user) '
    End Function

    ' //获取验证码是否需要使用
    Public Function getSetImag(type As String) As String

        Return HttpPost("&api=getsetimag.in&type=" + type) '
    End Function

    '//获取验证码
    Public Function getImage() As Bitmap

        Dim url As String = AppEn_ImgaUrl + AppEn_BSphpSeSsl
        Dim img As Image = GetImageFromWeb(url)
        Return img
    End Function


    '获取验证码，此调用会将图片下载到本地（程序运行目录），然后返回本地图片路径。
    Public Function GetImageFromWeb(url As String) As Image

        '// 地址
        Dim hreq As HttpWebRequest = HttpWebRequest.Create(url)
        hreq.Method = "POST"
        hreq.ContentType = "application/x-www-form-urlencoded; charset=UTF-8"
        hreq.Accept = "Application/json, Text/javascript, */*; q=0.01"

        hreq.Timeout = 30 * 1000


        Try
            Dim byte1 As Byte() = Encoding.UTF8.GetBytes(UrlEncode(url))
            hreq.ContentLength = byte1.Length

            Dim poststream As Stream = hreq.GetRequestStream()
            poststream.Write(byte1, 0, byte1.Length)
            poststream.Close()

            Dim hres As HttpWebResponse = hreq.GetResponse()
            System.Threading.Thread.Sleep(200)

            ' // 解析返回结果
            Dim stre As Stream = hres.GetResponseStream()
            Dim img As Image = Image.FromStream(stre)

            Return img

        Catch ex As Exception
            Return Nothing
        End Try


    End Function

    Public Function UrlEncode(url As String, Optional isUpper As Boolean = False) As String
        Return UrlEncode(url, System.Text.Encoding.UTF8, isUpper)
    End Function

    Public Function UrlEncode(url As String, encoding As Encoding, Optional isUpper As Boolean = False) As String
        Dim result As String = HttpUtility.UrlEncode(url, encoding)
        If (Not isUpper) Then Return result
        Return GetUpperEncode(result)
    End Function


    Public Function GetUpperEncode(encode As String) As String
        Dim result As StringBuilder = New StringBuilder()
        Dim Index As Integer = Integer.MinValue

        For i = 0 To encode.Length - 1
            Dim character As String = encode(i).ToString()
            If character = "%" Then Index = i
            If (i - Index = 1 OrElse i - Index = 2) Then character = character.ToUpper()
            result.Append(character)
        Next

        Return result.ToString()
    End Function



    '//查状态码字典,显示结果
    Public Sub ShowSult(sult As String, text As String)

        If (StatusCode.ContainsKey(sult)) Then
            MessageBox.Show(StatusCode.Item(sult), text) '
        Else
            MessageBox.Show(sult, text) '
        End If
    End Sub


    Private Function HttpPost(ByVal postdata As String) As String
        Dim hreq As HttpWebRequest = CType(HttpWebRequest.Create(AppEn_Url), HttpWebRequest)
        hreq.Method = "POST"
        hreq.ContentType = "application/x-www-form-urlencoded; charset=UTF-8"
        hreq.Timeout = 30 * 1000

        Try
            Dim apiValue As String = ExtractApiName(postdata)
            DebugApiLog("api名称", apiValue)

            Dim sdate As String = DateTime.Now.ToString("yyyy-MM-dd#HH:mm:ss")
            Dim appsafecode As String = MD5Encrypt32(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            Dim strPost As String = BuildApiPlainData(postdata, appsafecode, sdate, apiValue)
            DebugApiLog("加密前", strPost)
            Dim requestEncoded As String = BuildEncryptedRequestPayload(strPost, apiValue)
            If requestEncoded = "" Then
                DebugApiLog("异常", "请求加密失败")
                Return ""
            End If

            Dim strPostdes = "parameter=" & requestEncoded
            Dim byte1 As Byte() = Encoding.UTF8.GetBytes(strPostdes)
            hreq.ContentLength = byte1.Length
            Dim poststream As Stream = hreq.GetRequestStream()
            poststream.Write(byte1, 0, byte1.Length)
            poststream.Close()
            Dim hres As HttpWebResponse = CType(hreq.GetResponse(), HttpWebResponse)
            System.Threading.Thread.Sleep(200)


            Dim strResponse As String = ""
            Dim reader As StreamReader = New StreamReader(hres.GetResponseStream(), Encoding.UTF8)
            strResponse = reader.ReadToEnd()
            Dim sjson As String = ParseEncryptedResponse(strResponse, apiValue)
            If sjson = "" Then
                DebugApiLog("异常", "响应解密失败")
                Return ""
            End If
            DebugApiLog("解密后", sjson)

            Dim respAppSafe As String = ExtractResponseField(sjson, "appsafecode")
            If respAppSafe <> "" AndAlso respAppSafe <> appsafecode Then Return ""

            If apiValue = "BSphpSeSsL.in" Then
                Dim sessl As String = ExtractResponseField(sjson, "SeSsL")
                If sessl <> "" Then Return sessl
            End If

            Dim code As String = ExtractResponseField(sjson, "code")
            Dim dataValue As String = ExtractResponseField(sjson, "data")
            DebugApiLog("返回核心", "code=" & code & ", data=" & dataValue)

            If code <> "" AndAlso dataValue <> "" Then Return code & "|" & dataValue
            If code <> "" Then Return code
            Return dataValue
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return ""
        End Try
    End Function

    Public Shared Function MD5Encrypt32(ByVal sourcestring As String) As String
        Dim p1 As String = sourcestring
        Dim pwd As String = ""
        Dim md5 As MD5 = MD5.Create()
        Dim s As Byte() = md5.ComputeHash(Encoding.UTF8.GetBytes(p1))

        For i As Integer = 0 To s.Length - 1
            pwd = pwd & s(i).ToString("x2")
        Next

        Return pwd
    End Function

    Private Function BuildEncryptedRequestPayload(ByVal plainData As String, ByVal apiName As String) As String
        Dim timeMd5 As String = MD5Encrypt32(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Dim aesKey As String = MD5Encrypt32(AppEn_ServerKey & timeMd5).Substring(0, 16)
        Dim encryptedB64 As String = AesCbcEncryptDoublePad(plainData, aesKey)
        If encryptedB64 = "" Then Return ""

        Dim sigMd5 As String = MD5Encrypt32(encryptedB64)
        Dim signatureContent As String = "0|AES-128-CBC|" & aesKey & "|" & sigMd5 & "|xml"
        Dim rsaB64 As String = RsaEncryptPkcs1(signatureContent, AppEn_ClientKey)
        If rsaB64 = "" Then Return ""

        Dim payload As String = encryptedB64 & "|" & rsaB64
        DebugApiLog("加密后", "api=" & apiName & " " & payload)
        Return UrlEncodeCxx(payload)
    End Function

    Private Function ParseEncryptedResponse(ByVal responseBody As String, ByVal apiName As String) As String
        Dim raw As String = UrlDecodeCxx(responseBody)
        If String.IsNullOrEmpty(raw) Then Return ""

        Dim firstPipe As Integer = raw.IndexOf("|"c)
        If firstPipe < 0 Then Return ""

        Dim respEnc As String = ""
        Dim respRsa As String = ""
        Dim secondPipe As Integer = raw.IndexOf("|"c, firstPipe + 1)
        Dim firstPart As String = raw.Substring(0, firstPipe)
        Dim hasOkPrefix As Boolean = firstPart.Length >= 2 AndAlso firstPart.Substring(0, 2).ToLower() = "ok"

        If hasOkPrefix AndAlso secondPipe > firstPipe Then
            respEnc = raw.Substring(firstPipe + 1, secondPipe - firstPipe - 1)
            respRsa = raw.Substring(secondPipe + 1)
        Else
            respEnc = firstPart
            respRsa = raw.Substring(firstPipe + 1)
        End If
        DebugApiLog("解密前", "api=" & apiName & " " & respEnc)

        Dim sigDec As String = RsaDecryptPkcs1(respRsa, AppEn_ServerKey)
        If sigDec = "" Then Return ""
        Dim parts As String() = sigDec.Split("|"c)
        If parts.Length < 4 Then Return ""

        Return AesCbcDecryptDoublePad(respEnc, parts(2))
    End Function

    Private Function AesCbcEncryptDoublePad(ByVal plain As String, ByVal keySeed As String) As String
        Try
            Dim keyString As String = If(keySeed.Length >= 16, keySeed.Substring(0, 16), keySeed.PadRight(16, "0"c))
            Dim key As Byte() = Encoding.UTF8.GetBytes(keyString)
            Dim iv As Byte() = Encoding.UTF8.GetBytes(keyString)
            Dim firstPad As Byte() = Pkcs7PadBytes(Encoding.UTF8.GetBytes(plain), 16)

            Using aes As New AesManaged()
                aes.Mode = CipherMode.CBC
                aes.Padding = PaddingMode.PKCS7
                aes.Key = key
                aes.IV = iv
                Using ms As New MemoryStream()
                    Using cs As New CryptoStream(ms, aes.CreateEncryptor(), CryptoStreamMode.Write)
                        cs.Write(firstPad, 0, firstPad.Length)
                        cs.FlushFinalBlock()
                    End Using
                    Return Convert.ToBase64String(ms.ToArray())
                End Using
            End Using
        Catch
            Return ""
        End Try
    End Function

    Private Function AesCbcDecryptDoublePad(ByVal b64 As String, ByVal keySeed As String) As String
        Try
            Dim keyString As String = If(keySeed.Length >= 16, keySeed.Substring(0, 16), keySeed.PadRight(16, "0"c))
            Dim key As Byte() = Encoding.UTF8.GetBytes(keyString)
            Dim iv As Byte() = Encoding.UTF8.GetBytes(keyString)
            Dim cipherBytes As Byte() = Convert.FromBase64String(b64)

            Using aes As New AesManaged()
                aes.Mode = CipherMode.CBC
                aes.Padding = PaddingMode.PKCS7
                aes.Key = key
                aes.IV = iv
                Using ms As New MemoryStream()
                    Using cs As New CryptoStream(ms, aes.CreateDecryptor(), CryptoStreamMode.Write)
                        cs.Write(cipherBytes, 0, cipherBytes.Length)
                        cs.FlushFinalBlock()
                    End Using
                    Dim unpadded As Byte() = Pkcs7UnpadBytes(ms.ToArray())
                    Return Encoding.UTF8.GetString(unpadded)
                End Using
            End Using
        Catch
            Return ""
        End Try
    End Function

    Private Function Pkcs7PadBytes(ByVal input As Byte(), ByVal blockSize As Integer) As Byte()
        Dim padLen As Integer = blockSize - (input.Length Mod blockSize)
        If padLen = 0 Then padLen = blockSize
        Dim output(input.Length + padLen - 1) As Byte
        Buffer.BlockCopy(input, 0, output, 0, input.Length)
        For i As Integer = input.Length To output.Length - 1
            output(i) = CByte(padLen)
        Next
        Return output
    End Function

    Private Function Pkcs7UnpadBytes(ByVal input As Byte()) As Byte()
        If input Is Nothing OrElse input.Length = 0 Then Return New Byte() {}
        Dim padLen As Integer = input(input.Length - 1)
        If padLen <= 0 OrElse padLen > 16 OrElse padLen > input.Length Then Return input
        For i As Integer = input.Length - padLen To input.Length - 1
            If input(i) <> CByte(padLen) Then Return input
        Next
        Dim output(input.Length - padLen - 1) As Byte
        Buffer.BlockCopy(input, 0, output, 0, output.Length)
        Return output
    End Function

    Private Function RsaEncryptPkcs1(ByVal plain As String, ByVal publicKeyB64Der As String) As String
        Try
            Dim publicDer As Byte() = Convert.FromBase64String(publicKeyB64Der)
            Dim rsaParams As RSAParameters = ParseSpkiPublicKey(publicDer)
            Using rsa As New RSACryptoServiceProvider()
                rsa.ImportParameters(rsaParams)
                Dim encrypted As Byte() = rsa.Encrypt(Encoding.UTF8.GetBytes(plain), False)
                Return Convert.ToBase64String(encrypted)
            End Using
        Catch
            Return ""
        End Try
    End Function

    Private Function RsaDecryptPkcs1(ByVal cipherB64 As String, ByVal privateKeyB64Der As String) As String
        Try
            Dim privateDer As Byte() = Convert.FromBase64String(privateKeyB64Der)
            Dim rsaParams As RSAParameters = ParsePkcs8PrivateKey(privateDer)
            If rsaParams.Modulus Is Nothing OrElse rsaParams.D Is Nothing Then
                DebugApiLog("异常", "RSA私钥解析失败")
                Return ""
            End If
            Using rsa As New RSACryptoServiceProvider()
                rsa.PersistKeyInCsp = False
                rsa.ImportParameters(rsaParams)
                Dim cipherBytes As Byte() = Convert.FromBase64String(cipherB64)
                If rsaParams.Modulus IsNot Nothing AndAlso cipherBytes.Length <> rsaParams.Modulus.Length Then
                    DebugApiLog("异常", "RSA密文长度不匹配，cipher=" & cipherBytes.Length & ", key=" & rsaParams.Modulus.Length)
                End If
                Dim decrypted As Byte() = rsa.Decrypt(cipherBytes, False)
                Return Encoding.UTF8.GetString(decrypted)
            End Using
        Catch ex As Exception
            DebugApiLog("异常", "RSA解密异常: " & ex.GetType().Name & " - " & ex.Message)
            Return ""
        End Try
    End Function

    Private Function ParseSpkiPublicKey(ByVal der As Byte()) As RSAParameters
        Dim idx As Integer = 0
        If Not TryReadAsn1Tag(der, idx, &H30) Then Return New RSAParameters()
        Dim topLen As Integer = 0
        If Not TryReadAsn1Length(der, idx, topLen) Then Return New RSAParameters()

        If Not TryReadAsn1Tag(der, idx, &H30) Then Return New RSAParameters()
        Dim algLen As Integer = 0
        If Not TryReadAsn1Length(der, idx, algLen) Then Return New RSAParameters()
        If idx + algLen > der.Length Then Return New RSAParameters()
        idx += algLen

        If Not TryReadAsn1Tag(der, idx, &H3) Then Return New RSAParameters()
        Dim bitStrLen As Integer = 0
        If Not TryReadAsn1Length(der, idx, bitStrLen) Then Return New RSAParameters()
        If bitStrLen < 2 OrElse idx + bitStrLen > der.Length Then Return New RSAParameters()
        idx += 1

        Dim bitStr(bitStrLen - 2) As Byte
        Buffer.BlockCopy(der, idx, bitStr, 0, bitStr.Length)
        idx = 0

        If Not TryReadAsn1Tag(bitStr, idx, &H30) Then Return New RSAParameters()
        Dim rsaSeqLen As Integer = 0
        If Not TryReadAsn1Length(bitStr, idx, rsaSeqLen) Then Return New RSAParameters()
        Dim modulus As Byte() = Nothing
        Dim exponent As Byte() = Nothing
        If Not TryReadAsn1Integer(bitStr, idx, modulus) Then Return New RSAParameters()
        If Not TryReadAsn1Integer(bitStr, idx, exponent) Then Return New RSAParameters()
        If modulus Is Nothing OrElse exponent Is Nothing Then Return New RSAParameters()

        Return New RSAParameters With {.Modulus = modulus, .Exponent = exponent}
    End Function

    Private Function ParsePkcs8PrivateKey(ByVal der As Byte()) As RSAParameters
        Dim idx As Integer = 0
        If Not TryReadAsn1Tag(der, idx, &H30) Then Return New RSAParameters()
        Dim topLen As Integer = 0
        If Not TryReadAsn1Length(der, idx, topLen) Then Return New RSAParameters()
        Dim version As Byte() = Nothing
        If Not TryReadAsn1Integer(der, idx, version) Then Return New RSAParameters()
        If version Is Nothing Then Return New RSAParameters()

        If Not TryReadAsn1Tag(der, idx, &H30) Then Return New RSAParameters()
        Dim algLen As Integer = 0
        If Not TryReadAsn1Length(der, idx, algLen) Then Return New RSAParameters()
        If idx + algLen > der.Length Then Return New RSAParameters()
        idx += algLen

        If Not TryReadAsn1Tag(der, idx, &H4) Then Return New RSAParameters()
        Dim keyOctLen As Integer = 0
        If Not TryReadAsn1Length(der, idx, keyOctLen) Then Return New RSAParameters()
        If keyOctLen <= 0 OrElse idx + keyOctLen > der.Length Then Return New RSAParameters()
        Dim keyOct(keyOctLen - 1) As Byte
        Buffer.BlockCopy(der, idx, keyOct, 0, keyOctLen)

        Return ParsePkcs1PrivateKey(keyOct)
    End Function

    Private Function ParsePkcs1PrivateKey(ByVal der As Byte()) As RSAParameters
        Dim idx As Integer = 0
        If Not TryReadAsn1Tag(der, idx, &H30) Then Return New RSAParameters()
        Dim topLen As Integer = 0
        If Not TryReadAsn1Length(der, idx, topLen) Then Return New RSAParameters()

        Dim version As Byte() = Nothing
        If Not TryReadAsn1Integer(der, idx, version) Then Return New RSAParameters()
        If version Is Nothing Then Return New RSAParameters()

        Dim modulus As Byte() = Nothing
        Dim exponent As Byte() = Nothing
        Dim d As Byte() = Nothing
        Dim p As Byte() = Nothing
        Dim q As Byte() = Nothing
        Dim dp As Byte() = Nothing
        Dim dq As Byte() = Nothing
        Dim inverseQ As Byte() = Nothing
        If Not TryReadAsn1Integer(der, idx, modulus) Then Return New RSAParameters()
        If Not TryReadAsn1Integer(der, idx, exponent) Then Return New RSAParameters()
        If Not TryReadAsn1Integer(der, idx, d) Then Return New RSAParameters()
        If Not TryReadAsn1Integer(der, idx, p) Then Return New RSAParameters()
        If Not TryReadAsn1Integer(der, idx, q) Then Return New RSAParameters()
        If Not TryReadAsn1Integer(der, idx, dp) Then Return New RSAParameters()
        If Not TryReadAsn1Integer(der, idx, dq) Then Return New RSAParameters()
        If Not TryReadAsn1Integer(der, idx, inverseQ) Then Return New RSAParameters()

        Dim keyLen As Integer = If(modulus Is Nothing, 0, modulus.Length)
        If keyLen <= 0 Then Return New RSAParameters()
        Dim primeLen As Integer = keyLen \ 2
        If keyLen Mod 2 <> 0 Then primeLen += 1

        Return New RSAParameters With {
            .Modulus = LeftPadToLength(modulus, keyLen),
            .Exponent = exponent,
            .D = LeftPadToLength(d, keyLen),
            .P = LeftPadToLength(p, primeLen),
            .Q = LeftPadToLength(q, primeLen),
            .DP = LeftPadToLength(dp, primeLen),
            .DQ = LeftPadToLength(dq, primeLen),
            .InverseQ = LeftPadToLength(inverseQ, primeLen)
        }
    End Function

    Private Function LeftPadToLength(ByVal src As Byte(), ByVal targetLen As Integer) As Byte()
        If src Is Nothing Then Return Nothing
        If targetLen <= 0 Then Return src
        If src.Length = targetLen Then Return src
        If src.Length > targetLen Then
            Dim cut(targetLen - 1) As Byte
            Buffer.BlockCopy(src, src.Length - targetLen, cut, 0, targetLen)
            Return cut
        End If

        Dim result(targetLen - 1) As Byte
        Buffer.BlockCopy(src, 0, result, targetLen - src.Length, src.Length)
        Return result
    End Function

    Private Function TryReadAsn1Tag(ByVal data As Byte(), ByRef idx As Integer, ByVal expectedTag As Byte) As Boolean
        If data Is Nothing OrElse idx < 0 OrElse idx >= data.Length Then Return False
        If data(idx) <> expectedTag Then Return False
        idx += 1
        Return True
    End Function

    Private Function TryReadAsn1Length(ByVal data As Byte(), ByRef idx As Integer, ByRef length As Integer) As Boolean
        length = 0
        If data Is Nothing OrElse idx < 0 OrElse idx >= data.Length Then Return False
        Dim first As Integer = data(idx)
        idx += 1
        If (first And &H80) = 0 Then
            length = first
            Return True
        End If

        Dim count As Integer = first And &H7F
        If count <= 0 OrElse count > 4 Then Return False
        If idx + count > data.Length Then Return False

        Dim value As Integer = 0
        For i As Integer = 0 To count - 1
            value = (value << 8) Or data(idx)
            idx += 1
        Next
        length = value
        Return True
    End Function

    Private Function TryReadAsn1Integer(ByVal data As Byte(), ByRef idx As Integer, ByRef output As Byte()) As Boolean
        output = Nothing
        If Not TryReadAsn1Tag(data, idx, &H2) Then Return False
        Dim length As Integer = 0
        If Not TryReadAsn1Length(data, idx, length) Then Return False
        If length <= 0 OrElse idx + length > data.Length Then Return False

        Dim result(length - 1) As Byte
        Buffer.BlockCopy(data, idx, result, 0, length)
        idx += length

        If result.Length > 1 AndAlso result(0) = 0 Then
            Dim trimmed(result.Length - 2) As Byte
            Buffer.BlockCopy(result, 1, trimmed, 0, trimmed.Length)
            output = trimmed
        Else
            output = result
        End If
        Return True
    End Function

    Private Function ExtractApiName(ByVal postdata As String) As String
        Dim m As Match = Regex.Match(postdata, "(?:^|&)api=([^&]+)")
        If Not m.Success Then Return ""
        Return HttpUtility.UrlDecode(m.Groups(1).Value)
    End Function

    Private Function BuildApiPlainData(ByVal postdata As String, ByVal appsafecode As String, ByVal sdate As String, ByVal apiName As String) As String
        Dim values As New Dictionary(Of String, String)(StringComparer.Ordinal)
        values("api") = apiName
        values("BSphpSeSsL") = AppEn_BSphpSeSsl
        values("date") = sdate
        values("md5") = AppEn_MD5
        values("mutualkey") = AppEn_HOST_KEY
        values("appsafecode") = appsafecode

        Dim raw As String = postdata.TrimStart("&"c)
        If raw <> "" Then
            Dim pairs As String() = raw.Split("&"c)
            For Each pair As String In pairs
                If pair = "" Then Continue For
                Dim eq As Integer = pair.IndexOf("="c)
                If eq < 0 Then
                    Dim keyOnly As String = UrlDecodeCxx(pair)
                    values(keyOnly) = ""
                Else
                    Dim key As String = UrlDecodeCxx(pair.Substring(0, eq))
                    Dim val As String = UrlDecodeCxx(pair.Substring(eq + 1))
                    values(key) = val
                End If
            Next
        End If

        Dim keys As List(Of String) = values.Keys.ToList()
        keys.Sort(StringComparer.Ordinal)
        Dim sb As New StringBuilder()
        For i As Integer = 0 To keys.Count - 1
            Dim k As String = keys(i)
            If i > 0 Then sb.Append("&")
            sb.Append(UrlEncodeCxx(k))
            sb.Append("=")
            sb.Append(UrlEncodeCxx(values(k)))
        Next
        Return sb.ToString()
    End Function

    Private Function UrlEncodeCxx(ByVal value As String) As String
        If value Is Nothing Then Return ""
        Dim sb As New StringBuilder()
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(value)
        For Each b As Byte In bytes
            Dim isUnreserved As Boolean =
                (b >= AscW("0"c) AndAlso b <= AscW("9"c)) OrElse
                (b >= AscW("A"c) AndAlso b <= AscW("Z"c)) OrElse
                (b >= AscW("a"c) AndAlso b <= AscW("z"c)) OrElse
                b = AscW("-"c) OrElse b = AscW("_"c) OrElse b = AscW("."c) OrElse b = AscW("~"c)
            If isUnreserved Then
                sb.Append(ChrW(b))
            Else
                sb.Append("%")
                sb.Append(b.ToString("x2"))
            End If
        Next
        Return sb.ToString()
    End Function

    Private Function UrlDecodeCxx(ByVal value As String) As String
        If String.IsNullOrEmpty(value) Then Return ""
        Dim bytes As New List(Of Byte)()
        Dim i As Integer = 0
        While i < value.Length
            Dim c As Char = value(i)
            If c = "%"c AndAlso i + 2 < value.Length Then
                Dim hex As String = value.Substring(i + 1, 2)
                Dim n As Integer
                If Integer.TryParse(hex, Globalization.NumberStyles.HexNumber, Globalization.CultureInfo.InvariantCulture, n) Then
                    bytes.Add(CByte(n And &HFF))
                    i += 3
                    Continue While
                End If
            ElseIf c = "+"c Then
                bytes.Add(CByte(AscW(" "c)))
                i += 1
                Continue While
            End If
            Dim chBytes As Byte() = Encoding.UTF8.GetBytes(New Char() {c})
            bytes.AddRange(chBytes)
            i += 1
        End While
        Return Encoding.UTF8.GetString(bytes.ToArray())
    End Function

    Private Function ExtractResponseField(ByVal payload As String, ByVal fieldName As String) As String
        Dim xmlPattern As String = "<" & Regex.Escape(fieldName) & ">(.*?)</" & Regex.Escape(fieldName) & ">"
        Dim mXml As Match = Regex.Match(payload, xmlPattern, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
        If mXml.Success Then Return HttpUtility.HtmlDecode(mXml.Groups(1).Value)

        Return ""
    End Function

    Private Sub DebugApiLog(ByVal stage As String, ByVal content As String)
        Dim safeContent As String = content
        If safeContent Is Nothing Then safeContent = ""
        System.Diagnostics.Debug.WriteLine("[BSPHP][" & stage & "] " & TruncateForLog(safeContent, 1200))
    End Sub

    Private Function TruncateForLog(ByVal text As String, ByVal maxLen As Integer) As String
        If text Is Nothing Then Return ""
        If text.Length <= maxLen Then Return text
        Return text.Substring(0, maxLen) & "...(truncated)"
    End Function

End Class
