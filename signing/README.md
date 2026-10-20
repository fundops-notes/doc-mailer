# 自签名证书（内部分发用）

`LPDocMailer-Internal.cer` 是本项目的自签名代码签名证书公钥。
Release 中的 `LP_Tool_win64_signed.zip` 内的 exe 已用对应私钥签名。

在受管控的电脑上运行签名版前，先导入本证书（无需管理员，每台电脑一次）：

```
certutil -user -addstore Root LPDocMailer-Internal.cer
certutil -user -addstore TrustedPublisher LPDocMailer-Internal.cer
```

说明：
- 自签名证书只对导入过它的电脑生效
- 若导入后仍被拦截，说明是策略强制 SmartScreen / EDR 特征拦截，需联系 IT 加白
- 带 .pfx 私钥的文件绝不能提交到本仓库
