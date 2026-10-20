# LP 文档分发助手（lp-doc-mailer）

面向私募基金 / 资管机构的一站式 LP 文件分发工具：**给每位 LP 生成专属水印加密 PDF，并自动创建 Outlook 草稿**（只建草稿、绝不自动发送）。

## 功能特性

- **逐 LP 独立防护**：每位 LP 独立水印文字 + 独立打开口令 + 独立权限口令（AES-256；允许打印，禁止修改 / 复制 / 批注 / 页面拼装），泄露可溯源
- **水印自适应**：主体名过长时自动换行或缩小字号，不截断；可选单份 / 三份（防裁剪）水印
- **LP 清单管理**：Excel(xlsx)/CSV 批量导入（按表头名识别列）、口令留空自动生成、界面内增删改
- **邮件模板**：标题 + HTML 正文，`{{LP名称}}` `{{水印名称}}` `{{打开口令}}` `{{日期}}` 字段批量替换；左侧源码 + 右侧实时预览，工具栏支持加粗 / 颜色 / 口令高亮块
- **Outlook 草稿**：逐 LP 建草稿、插附件，只 `Save` 不 `Send`——发送永远需要人工复核
- **权限口令隔离**：模板字段中没有权限口令，权限口令从产品层面无法进入邮件

## 下载（免安装）

到 [Releases](../../releases) 页面下载 `LP_Tool_vX.X_win64.zip`，解压后双击
`LP_Tool.exe` 即可，**无需安装 Python 或任何依赖**（Windows 10/11 x64）。

- 首次运行若弹 SmartScreen：更多信息 → 仍要运行
- 压缩包内的 `lp_data.example.json` 与 exe 放同目录，首次启动自动载入演示数据
- Outlook 草稿功能需本机安装 Outlook 并配置账户（未安装时"仅生成加密文件"仍可用）

> exe 未做代码签名，个别杀毒软件可能误报，可在信任区添加后使用。

## 快速上手（5 分钟）

1. ① LP 信息管理 →「导入 Excel/CSV」→ 选 `examples/示例LP清单（演示数据）.xlsx`
2. ② 邮件模板 → 编辑正文，右侧实时预览
3. ③ 生成文件 / Outlook 草稿 → 添加原件 PDF →「生成加密文件 + Outlook草稿」
4. 打开 Outlook 草稿箱逐封复核后手动发送

详细说明见 `docs/使用说明.txt`。

## 示例数据

`examples/` 下的 LP 名称、邮箱（保留域名）、口令均为**虚构演示数据**，可放心使用：
- `示例LP清单（演示数据）.xlsx` — 8 条演示 LP，覆盖超长名称等边界情况
- `LP清单导入模板.xlsx` — 空白导入模板，表头即识别依据
- `lp_data.example.json` — 程序首次运行的种子数据（复制到 exe 同目录生效）
- `演示原件.pdf` — 2 页模拟文件，用于测试全流程

## 安全须知

- exe 同目录的 `lp_data.json` 保存全部 LP 信息与**权限口令**，属敏感文件，请勿外发、勿提交到仓库（已在 `.gitignore` 排除）
- 权限位由阅读器（Acrobat / Edge / WPS 等）强制执行，能阻挡常规编辑与复制；持有打开口令者理论上可借助工具破解，这是 PDF 标准的固有边界
- 发给 LP 的只有：加密文件 + 打开口令

## 从源码运行

```bash
pip install -r requirements.txt
python src/lp_tool.py
```

打包 exe：`pip install pyinstaller && pyinstaller --onefile --noconsole --name LP_Tool src/lp_tool.py`

## 目录结构

```
├── src/                 # 源码（lp_tool.py 主程序 / watermark_encrypt.py 核心逻辑）
├── examples/            # 演示数据与导入模板
├── docs/                # 使用说明
└── .github/workflows/   # 打 tag 自动构建 exe 并发布 Release
```

发布新版本：`git tag v0.2 && git push --tags`，Action 自动打包 exe 上传到 Releases。

## 许可证

AGPL-3.0（见 LICENSE）。本仓库已含全部源码，符合 PyMuPDF 的 AGPL 分发要求。
