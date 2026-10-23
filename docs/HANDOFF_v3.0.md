# HANDOFF：文档分发助手 v3.0 重构

> 交接文档。给接手实现的 agent / 同事。
> 最后更新：2026-10-23（v2.2.0 发布后，v3.0 设计完成、尚未开工）

---

## 0. 一句话现状

桌面工具"文档分发助手"当前发布版本 **v2.2.0**（可用、功能正常）。
**v3.0 的重构设计已完成并评审中，代码一行未动。** 接手者的任务是按设计稿实现 v3.0。

---

## 1. 项目坐标

| 项 | 值 |
|---|---|
| 工作区根 | `C:\Users\ian\.zcode\workspace\draft-cb79fb36-0da4-49bb-bce6-0ba69eabb0e3` |
| 主程序 | `lp_tool.py`（2450 行，GUI，tkinter） |
| 核心模块 | `watermark_encrypt.py`（427 行，水印 + AES 加密 + Word 导出） |
| Git 仓库 | `lp-doc-mailer/`（main 分支，remote = github.com/fundops-notes/doc-mailer） |
| 仓库内源码副本 | `lp-doc-mailer/src/lp_tool.py`、`src/watermark_encrypt.py`（**必须与根目录同步**） |
| 回归测试 | `lp-doc-mailer/tests/test_core.py`（9 条 pytest，CI 每次发版先跑） |
| 分发形态 | 便携版（绿色版）47MB · 签名 exe 109MB · CI exe 41MB |
| 设计稿 | `docs/设计稿_v3.0_功能设计.md`（**先读这个**） |
| 界面原型 | `C:\Users\ian\.zcode\cli\visualizations\ab53eea70e1fa7571095f9f81f363cbb6ab77bd7ee2d5993b1ff855161b31992\doc-mailer-v3-ui-mockup.html` |

**技术栈**：Python 3.14 + tkinter + PyMuPDF 1.28 + pywin32(Word/Outlook COM) + openpyxl + Pillow。
**运行环境**：Windows 10/11 x64，**用户机器是非中文 Windows**（这导致过多次编码事故，见 §6）。

---

## 2. v3.0 要解决什么

v2.x 所有配置（收件人/附件/Word模板/邮件模板/水印设置）平铺在单个 `lp_data.json` 里，
**不同业务场景无法共存**——做第二季对账单时第一季的清单、附件、模板会被覆盖。

v3.0 引入 **场景 Scenario**（可复用模板，"怎么做"）+ **发送事件 Event**（一次分发，"做给谁"），
并把文档生成重构为 **流水线 Step**（可组合步骤链），每个步骤都明确"产出给谁、字段从哪来"。

完整设计见 `docs/设计稿_v3.0_功能设计.md`，**开工前必读**。

---

## 3. 实施分期（设计稿 §7）

| 期 | 范围 | 交付物 |
|---|---|---|
| **P0 数据层** | 目录化存储、事件/场景/收件人/附件/流水线数据模型、v2 自动迁移 | 数据能存能取，v2 功能不丢 |
| **P1 场景与事件** | 场景模板、事件列表、事件工作台骨架、左侧步骤导航 | 界面骨架走通 |
| **P2 流水线** | 六种步骤、映射规则、字段映射 UI、Word 格式保持 | Word→PDF→水印→加密 链路 |
| **P3 执行引擎** | 预检、流水线执行、附件组装、草稿生成、Run 记录、只重跑失败 | 一键出草稿 |
| P4 收尾 | 收件人库、执行历史查看、示例数据更新 | 完整可用（可延后） |

**P0+P1 是地基，P2+P3 才是价值主体。** 建议 P0–P3 一次做完。

---

## 4. 关键设计决策（已定，除非用户改口否则照做）

1. **场景仅作模板**：新建事件时复制一份配置，**改场景不影响已创建的事件**（避免把已发批次改掉）。
2. **数据存储目录化**：`lp_data/{settings.json, events/, scenarios/, runs/}`，替代单文件。
   单文件会膨胀到几万行，保存慢且易损坏。
3. **v2 迁移**：首次启动检测到 `lp_data.json` 而无 `lp_data/` → 自动迁移，原文件改名保留 `lp_data.json.bak`（**不删除**）。
4. **Word 格式保持**：占位符替换必须用 Word COM 的 `Find.Execute(Replace=wdReplaceAll)`，
   只改文字不重建文档。**不要用 python-docx 重建**（会丢格式）。
5. **收件人 `fields` 字典**是本次重构的核心新增——让"文档字段"与"收件人身份"解耦。
6. **合并组语义**：同 `group` 的收件人共用一封邮件，To/CC/BCC 合并去重，附件包合并。
7. **权限口令隔离**（铁律）：权限口令**绝不能**进入邮件正文/标题/任何草稿字段。
   模板字段表里没有权限口令，是架构保证而非靠自觉。
8. **只建草稿绝不发送**：全代码库唯一的 Outlook 写操作是 `mail.Save()`，**不存在任何 Send 调用**。

---

## 5. 领域规则（回归测试必须守住）

| 规则 | 实现位置 | 怎么验 |
|---|---|---|
| 打开口令 ≠ 权限口令 | `validate_records()` + `process_one()` 入口双校验 | 同口令时 `authenticate()` 返回 6（owner 级）→ 必须被拒 |
| 权限口令不进邮件 | `render_template()` 只映射 4 个字段 | 正文中搜不到 owner_pw |
| 只建草稿不发送 | `create_outlook_draft()` 只调 `.Save()` | 全文件 grep `Send` |
| AES-256 + 仅打印 | `PDF_ENCRYPT_AES_256` + `PERM_PRINT\|PRINT_HQ\|ACCESSIBILITY` | 权限位掩码 |
| 水印不截断 | `plan_layout()` 折行/缩号 | 视觉验收 |
| 文件体积可控 | `doc.subset_fonts()` 子集化 | 不调会从 0.1MB 涨到 21MB |
| 损坏数据不覆盖 | `_load_data_file()` 返回 (data, error)，损坏时改名留证 | 坏 JSON 不被演示数据覆盖 |

---

## 6. 踩过的坑（别再踩）

**编码类（用户是非中文 Windows）**：
- ⚠️ **bat 文件名和内容必须纯 ASCII**。曾用 `执行更新.bat`（中文名）+ GBK 编码内容 → 公司电脑更新失败、显示乱码。现在固定：`run_update.bat` / `start.bat` / `app/` 目录名 / `encoding="ascii"` + `chcp 65001`。
- ⚠️ 便携包 zip 顶层目录必须是 ASCII（`app/`），robocopy 引用中文目录名会静默失败。
- ⚠️ **不要用 `Segoe UI Emoji`** 字体做首页图标 → 非中文 Windows 上 GDI 挂起白屏。首页用纯 ttk 组件。
- ⚠️ 写 Python 源码时若用 heredoc 传含 `\n` 的字符串，会被转成真实换行导致语法错误。**改代码用 Write 工具写脚本文件再执行**，不要 heredoc 拼字符串。

**打包类**：
- ⚠️ **每次改完代码必须 `cp lp_tool.py watermark_encrypt.py LP_Tool_绿色版/`**，
  否则 zip 里的还是旧代码（曾因此发布出"显示 1.8.0 但实际是新版源码"的错乱）。
- ⚠️ **版本号 `APP_VERSION` 必须同步改**。曾 tag v1.7.1 但代码里还是 "1.7.0"。
- ⚠️ 打包后**读 zip 内的 lp_tool.py 验证版本号和关键函数**再上传：
  ```python
  with zipfile.ZipFile(out) as z:
      src = z.read('app/lp_tool.py').decode('utf-8')
      assert 'APP_VERSION = "3.0.0"' in src
  ```

**功能类**：
- ⚠️ `TextWriter.write_text(morph=...)` 在 PyMuPDF 1.28 **静默丢字**。水印旋转必须走"方形临时页 + `show_pdf_page(rotate=)`"。
- ⚠️ 附件路径里的**英文逗号**不能当分隔符（基金名常有 "X, L.P."）→ 只按 `;` 和换行拆。
- ⚠️ `on_close` 里访问只在别的页面创建的变量 → AttributeError 导致窗口关不掉。访问前 try/except。

---

## 7. 发版流程（照做，别自己发明）

```bash
cd <工作区根>
# 1. 改代码 → 自测 → 同步到绿色版和仓库
cp lp_tool.py watermark_encrypt.py LP_Tool_绿色版/
cp lp_tool.py watermark_encrypt.py lp-doc-mailer/src/
cd lp-doc-mailer && python -m pytest tests/ -q && cd ..

# 2. 打包 exe（签名）
python -m PyInstaller --onefile --noconsole --name LP_Tool -y lp_tool.py
cp dist/LP_Tool.exe LP_Tool.exe
powershell -NoProfile -Command "
  \$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2('签名\LPDocMailer-Internal.pfx','LPTool2026!sign')
  Set-AuthenticodeSignature -FilePath 'LP_Tool.exe' -Certificate \$cert | Out-Null
  Write-Host ('签名: ' + (Get-AuthenticodeSignature 'LP_Tool.exe').Status)"

# 3. 打便携 zip（顶层必须是 app/）
python -c "
import zipfile, os
with zipfile.ZipFile('doc-mailer_vX.Y.Z_portable.zip','w',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
    for root,dirs,files in os.walk('LP_Tool_绿色版'):
        for f in files:
            fp=os.path.join(root,f); z.write(fp, os.path.join('app',os.path.relpath(fp,'LP_Tool_绿色版')))"
# 验证 zip 内容再上传！

# 4. 提交 + 打 tag（tag 会触发 CI 自动构建）
cd lp-doc-mailer && git add -A && git commit -m "vX.Y.Z：..." && git push origin main
git tag vX.Y.Z && git push origin vX.Y.Z

# 5. 上传便携包到 Release（命名必须是 doc-mailer_vX.Y.Z_portable.zip）
#    自动更新按此名前缀匹配，改名会让老版本更新失败
```

**下载包命名规范**（不可改）：`doc-mailer_v{版本}_{portable|win64|signed}.zip`

---

## 8. 开工前必须确认的 5 个点

设计稿 §8 有详述。**建议先问用户**，别自行决定：

1. **场景改动是否联动已创建事件**——设计稿默认"不联动"（复制式）。用户可能想要联动。
2. **`from_sheet` 模式是否必需**——外部 Excel join 收件人。砍掉可简化流水线。**优先级较高，建议先问。**
3. **自定义字段录入方式**——收件人表任意加列（灵活）vs 预定义字段模板（易用）。设计稿倾向"并存"。
4. **旧版 `lp_data.json.bak` 保留多久**。
5. **是否需要多事件并行执行**（当前设计一次一个）。

---

## 9. 已知未做/可延后

- IMAP 草稿后端（不装 Outlook 也能写草稿）——用户问过，尚未决定做
- SMTP 直发——**不建议**（跳过人工复核，与"只建草稿"的产品承诺冲突）
- DPAPI 加密 `lp_data.json`（当前明文，仅靠目录管控 + 界面红字提示）
- 单实例保护 / DPI 感知 / 界面字体自适应 / emoji 索引偏移 / HTML `<li>/<a>` 保真
- 执行历史查看界面（P4）

---

## 10. 接手检查清单

- [ ] 读 `docs/设计稿_v3.0_功能设计.md` 全文
- [ ] 打开界面原型 HTML 点一遍五个页面
- [ ] 跑 `pytest lp-doc-mailer/tests/test_core.py` 确认基线绿
- [ ] 跑一次现有程序（`python lp_tool.py`）确认 v2.2.0 正常
- [ ] 确认 §8 的 5 个决策点已与用户对齐
- [ ] 开工前先备份：`cp lp_tool.py lp_tool.py.bak`
