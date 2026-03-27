# BSPHP VB.NET Demo




本目录包含对 `BSPHP-MAC-IOS-DEMO-master` 下两个 macOS 示例的 VB.NET 复刻版本：

- `BSPHP.LoginDemoVB`：账号模式（登录/注册/找回/反馈等）
- `BSPHP.CardDemoVB`：卡密模式（卡密验证、机器码账号、充值续费）

## 目录独立说明

- `BSPHP.LoginDemoVB` 为独立目录：可直接打开 `BSPHP.LoginDemoVB\BSPHP.LoginDemoVB.sln`
- `BSPHP.CardDemoVB` 为独立目录：可直接打开 `BSPHP.CardDemoVB\BSPHP.CardDemoVB.sln`
- 根目录 `BSPHP-VB.NET-DEMO.sln` 仅作为双项目总入口，方便切换，不影响各目录独立运行

说明：

- 由于当前环境未安装 `dotnet`，我已按 VB.NET WinForms 项目结构完整生成源码。
- 项目目标框架已调整为 `.NET Framework 4.8`（`net48`），并改为传统 `vbproj`（非 SDK）格式，兼容旧版 Visual Studio 工程加载。
- 请先安装 `.NET Framework 4.8 Developer Pack`，再用 Visual Studio 打开解决方案。
- 建议优先打开根目录 `BSPHP-VB.NET-DEMO.sln`（可一次看到并切换两个示例项目）。
- 也可单独打开对应 `.vbproj` 编译运行。
- API 加密协议已按原项目实现：`MD5 + AES-128-CBC + RSA PKCS#1`。

## 常见问题处理（先做这个）

- 如果出现“项目已卸载”“无法加载项目”“编译异常”等问题，请先关闭 Visual Studio。
- 删除各项目目录下的缓存文件夹：`.vs`、`obj`、`bin`（如果存在）。
- 重新打开 `BSPHP-VB.NET-DEMO.sln` 再尝试加载/编译。

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
