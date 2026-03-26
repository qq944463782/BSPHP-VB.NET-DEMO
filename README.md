# BSPHP VB.NET Demo

当前项目跑不起来，有人可以就 调试 vs219 2022 都不可以


本目录包含对 `BSPHP-MAC-IOS-DEMO-master` 下两个 macOS 示例的 VB.NET 复刻版本：

- `BSPHP.LoginDemoVB`：账号模式（登录/注册/找回/反馈等）
- `BSPHP.CardDemoVB`：卡密模式（卡密验证、机器码账号、充值续费）

说明：

- 由于当前环境未安装 `dotnet`，我已按 VB.NET WinForms 项目结构完整生成源码。
- 你在本机 Visual Studio 打开对应 `.vbproj` 即可编译运行。
- API 加密协议已按原项目实现：`MD5 + AES-128-CBC + RSA PKCS#1`。

## 配置信息文件路径

- 账号模式配置文件：`BSPHP.LoginDemoVB\BSPHPClient.vb`
  - 配置类：`BSPHPConfig`
  - 需要改的字段：`Url`、`MutualKey`、`ServerPrivateKey`、`ClientPublicKey`、`CodeUrlPrefix`

- 卡密模式配置文件：`BSPHP.CardDemoVB\BSPHPClient.vb`
  - 配置类：`BSPHPConfig`
  - 需要改的字段：`Url`、`MutualKey`、`ServerPrivateKey`、`ClientPublicKey`（验证码地址继承 `CodeUrlPrefix`）

## 验收顺序（建议）

1. 先运行 `BSPHP.LoginDemoVB`
2. 查看顶部公告和状态是否显示“服务已连接”
3. 在“密码登录”页点击“预填演示参数”，再点“刷新验证码”
4. 切换测试：短信登录、邮箱登录、账号注册、短信/邮箱注册、短信/邮箱找回
5. 登录成功后确认独立控制台窗口弹出
6. 再运行 `BSPHP.CardDemoVB`
7. 在“制作卡密登录模式”测试登录、验证码刷新、续费入口按钮
8. 在“机器码账号模式”测试 AddCardFeatures + login.ic + chong.ic
9. 登录成功后确认“卡密主控制面板”中接口分组与续费入口可用
