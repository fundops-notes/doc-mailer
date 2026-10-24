# HANDOFF：文档分发助手 v3.2（P4 收尾：收件人库 + 执行历史）

> 交接文档。给接手实现的 agent / 同事。
> **文件名沿用 `HANDOFF_v3.0.md`（历史沿革，改名会打断各处引用），内容已覆盖到 v3.2.0。**
> 最后更新：2026-10-24（**v3.1.0 / v3.1.1 已发版；v3.2.0（P4 收尾）已实施完毕并验证通过，但尚未发版**，见 §16）
>
> **当前代码版本：`3.3.0`**（`lp_tool.py` 与 `lp_v3.py` 两处一致）。
> 已发布：v2.2.0（旧）、v3.1.0 / v3.1.1、**v3.2.0**（P4 收尾，见 §16）、**v3.3.0**（附件哈希去重，见 §18）。

---

## 0. 一句话现状

桌面工具"文档分发助手"当前**已发布**版本仍是 **v3.1.1**（v3.1.0 的修复版，见 §15）。
**P4 收尾（收件人库跨事件复用 + 执行历史查看界面 + 示例数据更新）已实现完毕并通过全量验证**
（191 条 pytest 全绿 + 实机截图验收），代码版本已升到 **v3.2.0**，**但尚未发版**（用户未授权本轮发版）。
代码在 `lp_tool.py`（4764 行）+ `lp_v3.py`（2884 行）+ `watermark_encrypt.py`（698 行）。
打包 / 签名 / git tag / Release 上传**均已执行**（v3.1.0 与 v3.1.1）。
下一步 = v3.2.0 发版（等用户拍板）。

---

## 0.1 源码坐标（v3.2 后已变）

| 项 | v2.2.0 | v3.2.0（当前代码） |
|---|---|---|
| 主程序 | `lp_tool.py` 2450 行 | `lp_tool.py` 4764 行（GUI，`APP_VERSION = "3.2.0"`） |
| 核心模块 | `watermark_encrypt.py` 427 行 | `watermark_encrypt.py` + **新增 `lp_v3.py`** 2884 行（纯逻辑核心，不依赖 tkinter） |
| 数据存储 | 单文件 `lp_data.json` | 目录 `lp_data/{settings.json, events/, scenarios/, runs/, **recipients/**}` |
| 回归测试 | `tests/test_core.py` 9 条 | `tests/` **191 条**（`test_core` 9 + `test_v3_data` 17 + `test_v3_pipeline` 39 + `test_v3_engine` 25 + `test_v3_gui` 18 + `test_v3_import_export` 25 + `test_v3_attach_paths` 20 + **`test_v3_pool` 21** + **`test_v3_pool_gui` 17**） |

> ⚠️ `lp_v3.py` 是**新增的第三个源文件**。PyInstaller 跟随 import 会自动带上（已验证进包）；
> 便携包按目录整体打包，`app/` 下三个 .py 必须齐。
> ⚠️ `recipients/` 是 P4 新增的收件人库目录（`lp_data/recipients/<id>.json`），
> **口令入库**，与 `lp_data/` 其余部分同管控级别。

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
| P4 收尾 | 收件人库、执行历史查看、示例数据更新 | ✅ **已完成（v3.2.0，见 §16）** |

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
      assert 'APP_VERSION = "3.1.1"' in src
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
- 执行历史查看界面（P4）→ ✅ **已做（v3.2.0，见 §16）**

---

## 10. 接手检查清单

- [x] 读 `docs/设计稿_v3.0_功能设计.md` 全文
- [x] 打开界面原型 HTML 点一遍五个页面
- [x] 跑 `pytest lp-doc-mailer/tests/test_core.py` 确认基线绿
- [x] 跑一次现有程序（`python lp_tool.py`）确认 v2.2.0 正常
- [x] 确认 §8 的 5 个决策点已与用户对齐
- [x] 开工前先备份：`cp lp_tool.py lp_tool.py.bak`

---

## 11. 实现记录（v3.0.0，P0–P3，2026-10-23）

### 11.1 已交付

| 期 | 范围 | 落地位置 |
|---|---|---|
| **P0** | 目录化存储、数据模型与归一化、原子写 + `.bak` 轮转、损坏文件改名隔离（`_quarantine`）、v2→v3 迁移（`migrate_v2_to_v3`，原文件改名 `.bak` **不删除**）、首次启动示例数据播种（`seed_demo`） | `lp_v3.py`：`Store` / `normalize_*` / `v2_to_v3_payload` / `needs_migration` / `seed_demo` |
| **P1** | 顶栏 + 左侧固定导航 + 右侧内容区；首页场景卡片 + 发送事件表；场景模板页含「同步到事件…」；事件工作台左侧 5 步常驻导航 + 底部操作条 | `lp_tool.py`：`App` / `_build_shell` / `show_screen` / `open_event` / `_wb_show_step` |
| **P2** | 六种步骤 `word_template`/`word_to_pdf`/`pdf_watermark`/`pdf_encrypt`/`pdf_source`/`static_file`；映射规则 `per_lp`/`per_group`/`all`/`from_sheet`；占位符检测（读 docx XML，不依赖 Word）；字段映射 UI | `lp_v3.py`：`execute_step`/`run_pipeline`/`build_units`/`step_placeholders`/`auto_map_placeholder`；`watermark_encrypt.py`：`watermark_pdf`/`encrypt_pdf`/`detect_placeholders`/`WordBatch` |
| **P3** | `preflight` → 流水线 → `assemble_attachments` → `draft_units` → `create_outlook_draft`（**只 `mail.Save()`**）→ 写 Run 记录；`failed_recipients` + 只重跑失败项；`preview_plan` 清单预览 | `lp_v3.py`：`run_event`/`preflight`/`preview_plan`/`failed_recipients` |

### 11.2 关键决策（已与用户对齐，§8 五点）

1. **场景复制式**：事件各存完整配置，改场景不影响已建事件；场景页提供显式「同步到事件…」（收件人按名称合并，保留已有口令与自定义字段，只补空缺）。
2. **保留 `from_sheet`**（保住 v2 外部数据表能力）。
3. **自定义字段「模板 + 临时加列」并存**。
4. **不并行**，一次跑一个事件。
5. **`lp_data.json.bak` 永不自动删除**。

### 11.3 额外决策

- **附件默认只附最终产物**（`attach_mode="final"`）：中间产物（如未加密 .docx）进邮件包会绕过加密目的。
  设计稿「各步骤产物」的字面理解留了 `attach_mode="all"` 开关，但**默认关闭**。
- **权限口令隔离双保险**：架构上不映射（`FORBIDDEN_FIELD_KEYS`）+ 渲染时 `scrub_owner_password` 兜底抹除。
- **合并组口令不同** → 预检 warning，正文取组内第一人（沿用 v2 与设计稿 per_group 语义）。

### 11.4 验证结果（全部实测，非推断）

| 项 | 结果 |
|---|---|
| pytest | `python -m pytest lp-doc-mailer/tests -q` → **81 passed，EXITCODE=0** |
| 真实 Outlook | `verify_outlook_draft.py`：真实建 item、设字段、加附件，**正文/标题均无权限口令**；丢弃不保存，草稿箱无残留 |
| 打包 | PyInstaller onefile 构建成功（`dist/LP_Tool.exe` 115345146 字节）；`lp_v3` 已进包（`Analysis-00.toc` + `PYZ-00.toc` 均命中）；`verify_exe_boot.py` 实启动 18s 存活、无 ImportError；exe 同级目录正确生成 `lp_data/` 并播种出 8 位收件人的示例事件 |
| 视觉 | `v3_home.png` + `v3_wb_recipients/pipeline/attachments/mail/run.png`（首页 + 工作台 5 步，与原型一致） |
| 便携包 | 用自带 runtime 导入成功，`APP_VERSION=3.0.0` |
| 同步 | `sync_release.py`（复制 + 校验一步到位）通过，`APP_VERSION：3.0.0` |

### 11.5 过程中发现并修复的真实缺陷（v2 也存在）

1. **`Find.Execute(Replace=2)` 命名参数在 pywin32 动态调度下被静默忽略** → 占位符一个都没替换。改用位置参数。
2. **`mail.Close(SaveChanges=0)` 同类问题** → 一律改位置参数 `Close(0)`。
3. **合并组 + 「每人一份」时附件包只收了组内第一人的文件** → 修 `_outputs_of`/`preview_plan`，已加回归测试。
4. **tkinter grid/pack 同父容器冲突** → 拆子容器。
5. **启动未调用 `show_screen("home")`** → 右侧空白。
6. **自动更新 `robocopy` 未排除 `lp_data`** → 可能覆盖用户数据；改为一条 robocopy 同时带 `/XF` 与 `/XD lp_data`。
7. **首次启动未播种示例数据** → 新增 `seed_demo`（内置 `MOCK_LPS` 兜底）。
8. **`WordBatch.close()` 调 `Quit()` 后释放 COM 引用会触发 0x800706be 致命异常（会杀进程）** → 改为只用 `DispatchEx` 私有实例、只释放引用不 `Quit`。

### 11.6 已知折中与遗留

- **Word COM 收尾**：`WordBatch` 不调用 `Quit()`，残留的隐藏 `WINWORD.EXE` 会存活到进程结束（换取不触发 0x800706be 杀进程）。
  普通 `Dispatch` 会挂到用户已开的 Word 上并可能误关用户 Word，已改 `DispatchEx`。
- **`from_sheet` 的 v2 遗留数据表路径在 v2 未落盘**，迁移后需用户重新选择 Excel（迁移提示已写明）。
- **临时脚本残留**：项目根 `shot_v3.py` / `shot_v3_wb.py` / `fix_bom.py`（截图与 BOM 清理用）。
  本次会话的回收站只允许工作区内路径，`.zcode` 目录下不可删，故保留。无功能影响，可随手删除。
- **保留的工具脚本**：`sync_release.py`（同步+校验）、`verify_outlook_draft.py`（真实 Outlook 验证，不留草稿）、`verify_exe_boot.py`（exe 启动验证）。

### 11.7 下一步（发版前必做）

发版流程照 §7，**不要自创**。发版前注意：

- `lp-doc-mailer/.gitignore` 已忽略 `build/`、`dist/`、`*.spec`、`lp_data.json`、`out_wm/`；
  新增的 `src/lp_v3.py` 与 `tests/` 下的新测试文件**需要 `git add`**。
- 打包 zip 后**必须**读 zip 内 `app/lp_tool.py` 断言 `APP_VERSION = "3.1.1"` 与关键函数存在再上传。
- 绿色版目录里打包前记得清 `__pycache__/`（打包脚本已排除 `lp_data/` 与 `__pycache__/`）。
- 下载包命名必须是 `doc-mailer_v3.1.1_portable.zip`（自动更新按此前缀匹配）。
- git tag `v3.1.1` 会触发 CI 自动构建。
- **P4（收件人库界面、执行历史查看界面）尚未做**，发版说明里要写清。

---

## 12. 用户反馈轮（2026-10-23 晚，尚未发版）

用户试用了 P0–P3 的界面，提出 4 条反馈。本轮已全部实现并验证。

### 12.1 反馈与处理

| # | 用户反馈 | 处理 |
|---|---|---|
| 1 | 首页点「打开工作台」没反应（第一次能开，关掉后再点就不出来） | **已修**。根因：`open_event()` 重建工作台后没有重新 `pack`，而 `show_screen()` 会把其它页 `pack_forget` → 第二次点击是空操作。修法：重建后强制 `self.show_screen("event")`。已加回归测试 `test_open_event_reopens_after_leaving_workbench`。 |
| 2 | 「场景与事件」「场景模板」「收件人库」三个入口看不懂区别，只要事件库；每个事件要有发送模式、收件人列表、要调用的功能；要看得到自由组合的流程图 | **已实现**。导航收敛为 **事件库 / 设置 / 帮助**；「场景模板」收进事件库的第二个标签页（不再是独立导航项）；「收件人库」这个空壳页删除（收件人本就属于事件）。事件表新增**发送模式**列；右侧新增**可视化流程图**（收件人 → 流水线步骤 → 最终产物 → Outlook 草稿，步骤可自由增删排序，流程图自动跟着变）。 |
| 3 | 所有 Excel 导入处都要有模板下载；导入要做数据校验、可视化导入结果 | **已实现**。收件人清单与外部数据表两处导入都加了「下载模板」（含**填写说明页**）。导入改为三步：选文件 → **逐行校验**（错误/警告分级）→ **可视化导入结果弹窗**（逐行列出状态与问题，可「导入无错误的行」「全部导入」「导出问题明细」）。 |
| 4 | 导入的数据、事件设置要能导出备份 | **已实现**。事件库工具条：导出收件人 Excel、导出全部收件人 Excel、导出事件 JSON、导入事件备份、备份全部（settings+events+scenarios+runs 的 JSON）。 |

### 12.2 新增核心能力（`lp_v3.py`，纯逻辑、可测）

- **发送模式** `options.send_mode`：`per_lp`（每人一份）/ `per_group`（按组合并）/ `all`（全部合并）。
  `send_mode_of()` 认 `send_mode`，没有就退回旧的 `options.merge_groups` 布尔（v2 迁移数据仍正确）。
  `draft_units()` 按模式决定草稿单元数。`normalize_options()` 保证 `send_mode` 与 `merge_groups` 始终一致。
- **`flow_nodes(event)`** → 流程图节点序列（纯数据，GUI 只负责画）。
- **`validate_recipient_rows(rows, existing_names)`** → 导入校验结果（fatal / 每行 errors+warnings / 未识别列 / 重复名）。
  校验项：表头缺名称列、空文件、名称为空、文件内重名、与已有同名、邮箱为空、邮箱格式错（To/CC/BCC）、
  **打开口令 == 权限口令（铁律，直接判错）**、打开口令过短、附件文件不存在、合并组提示。
- **`recipients_to_rows(recipients, include_passwords)`** → 收件人导出为二维列表（自定义字段列并到末尾）。
- **`event_export_payload()` / `event_import_payload()`** → 单事件备份与还原。
- **`backup_payload(store)`** → 整库备份。

### 12.3 本轮顺带修掉的真实布局缺陷

- **收件人表 / 附件表被挤到右下角一小块**：`ttk.Treeview` 建在了 `parent` 里却 `pack` 进 `frame`，
  于是被排到 frame 下面、只拿到请求高度。修法：Treeview 直接建在 `frame` 里。附件步骤同样有这问题，已一并修。
- **`tk.PanedWindow.add(child, weight=…)`** 在本机 Tk 上报 `unknown option "-weight"` → 改为 `add` + `paneconfigure`（失败也不影响使用）。
- 事件表「最近执行」的成功数一直统计 `item["status"]`，但 Run 记录里根本没有 `status` 字段（是 `draft`）→ 永远显示 0/N，已改。

### 12.4 本轮验证

- `python -m pytest lp-doc-mailer/tests -q` → **112 passed，EXITCODE=0**（新增 `tests/test_v3_import_export.py` 25 条）。
- `verify_v4_features.py`：**30 项端到端检查全过，0 失败**（真实写 xlsx/json 文件，覆盖导出收件人、
  下载模板、导入三步流程、导出/导入事件备份、整库备份、导出全部收件人，并复核权限口令不进映射与正文）。
- 视觉：`v4_library.png`（事件库 + 流程图）、`v4_recipients.png`（收件人步骤 + 发送模式）、`v4_import_result.png`（导入结果弹窗）。

### 12.5 本轮保留/新增的工具脚本

- `verify_v4_features.py` —— 新增/导出/备份的端到端验证（可重跑）。
- `shot_v4.py` —— 截图脚本（同 `shot_v3.py` 那批）。
- ⚠️ 同上一轮：**临时脚本删除受本会话安全策略限制**（可恢复删除只允许工作区内路径，
  `.zcode` 目录下的文件删不掉），故 `shot_v4.py` 等仍留在项目根，无功能影响。

---

## 13. 概念澄清轮（2026-10-23 晚，同属 3.0.0，仍未发版）

用户对界面上的两个概念提出疑问，排查后确认了两处**真实设计缺陷**，已一并修掉。

### 13.1 「源 PDF」vs「静态文件」——代码上本来就是同一件事

核实结果：`execute_step()` 里两者是**同一个分支**（`if stype in ("pdf_source", "static_file")`），
行为完全一致——把磁盘上的文件复制到输出目录，文件名按收件人/组名生成。差别只有文案和「预期用法」：

| 步骤 | 实际行为 | 预期用法 |
|---|---|---|
| **源 PDF** | 从磁盘选一个 PDF → 复制成每人一份 | 后面接「加水印」「加密」，把外部 PDF 处理成每人一份的加密件 |
| **静态文件** | 从磁盘选一个任意文件（Word/Excel/图片）→ 原样复制成每人一份 | 通常放最后一步，直接当附件 |

> ⚠️ 用户原本以为「源 PDF = 从上一步拿 PDF」——**不是**。它是从磁盘选外部原件。
> 「从上一步拿」是 `word_to_pdf` / `pdf_watermark` / `pdf_encrypt` 这三个步骤的语义。

已做的处理：
- 「添加步骤」弹窗里把两者说明写清楚，并加了一段提示：两者代码上同一件事，
  区别只在用法；**要往邮件里塞文件但不想走流水线，用「附件」那一步更省事**。
- 步骤参数里的标签：`源 PDF：` / `源文件：`。

### 13.2 修掉的真缺陷：`final` 模式会把加密件挤出附件包

**旧逻辑**：`attach_mode="final"` 只取流水线**最后一步**的产物。
若流水线是 `源 PDF → 加密 → 静态文件`，最终附件里只剩静态文件的副本，
**加密后的 PDF 根本没进邮件**（等于加密白做了，而且发出的是未加密原件）。

**新规则**：`final = 最后一步的产物 ∪ 所有加密步骤的产物`。
回归测试 `test_final_attach_mode_keeps_encrypted_pdf_when_static_file_is_last`。

### 13.3 把「输出粒度」和「数据来源」拆成两个下拉

**旧问题**：步骤参数里一个下拉塞了 4 个值，其中 `from_sheet`（按 Excel 行匹配）
**根本不是粒度，是数据来源**——它的产出单位仍然是「每人一份」（已核实：3 位收件人 → 3 个产出单元）。
用户看到的「每人1份 / 每合并组一份 / 全员公用一份 / 按excel匹配」因此无法理解。

**新结构**（`lp_v3.py`）：

| 字段 | 取值 | 含义 |
|---|---|---|
| `mapping["mode"]`（输出粒度） | `per_lp` / `per_group` / `all` | **这一步产出几份文件** |
| `mapping["source"]`（数据来源） | `recipient` / `sheet` | **这一步的占位符字段从哪来** |

- 选「从外部 Excel 数据表取」→ 输出粒度自动锁定「每人一份」（数据表是按人 join 的），并显示数据表配置区。
- **旧数据兼容**：`mapping["mode"] == "from_sheet"` 在 `normalize_step()` 里自动拆成
  `mode="per_lp" + source="sheet"`，老文件读取后无需手工迁移。
- `execute_step()` 现在只在 `source == "sheet"` 时才去查外部数据表行。
- 界面文案：`OUTPUT_MODE_LABELS` / `OUTPUT_MODE_DESC` / `DATA_SOURCE_LABELS` / `DATA_SOURCE_DESC`，
  步骤卡与流程图节点同步显示（外部数据表来源显示为「每人一份 · 外部数据表」）。

### 13.4 顺带修的

- `sync_release.py` 之前把「绿色版存在 `lp_data/`」当**同步失败**（exit 1）——
  但用户从绿色版测试时必然会生成 `lp_data/`，于是同步脚本永远红。
  改为：平时只**提醒**，`--pack`（真正打包前）才**拦截**；`__pycache__/` 同理。

### 13.5 两处「模式」的对照（回答用户提问）

| | 事件信息里的**发送模式** | 步骤参数里的**输出粒度** |
|---|---|---|
| 作用对象 | 整个事件 | 单个步骤 |
| 决定什么 | **几封邮件** | **这一步产出几份文件** |
| 取值 | 每人一份 / 按组合并 / 全部合并 | 每人一份 / 每组合并一份 / 全员共用一份 |

例：3 位收件人（甲乙同组 G1、丙单独）
- 发送模式=每人一份 → 3 封草稿；发送模式=按组合并 → 2 封；全部合并 → 1 封
- 某个步骤粒度=每人一份 → 3 个文件；每组合并一份 → 2 个；全员共用一份 → 1 个

### 13.6 本轮验证

- `python -m pytest lp-doc-mailer/tests -q` → **115 passed，EXITCODE=0**（新增 3 条：
  粒度往返、数据来源切换与锁定、老 from_sheet 数据迁移；另加 final 模式加密件回归 1 条）。
- `verify_concepts.py` → 概念与修复复核通过。
- 视觉：`v4_pipeline.png`（步骤参数区两个下拉 + 说明文字）。


---

## 14. v3.1 流水线重构实施记录（2026-10-23 晚，Q1–Q5 用户已确认）

> 设计稿：`docs/设计稿_v3.1_流水线重构.md`。用户逐条回答 Q1–Q5 后指示「不用考虑旧配置，
> 我们就是全新开发。开始实施」。本节记录**已落地**的实现与验证结果。

### 14.1 核心模型：六种步骤 → 两种产出条目

| 旧（v3.0） | 新（v3.1） |
|---|---|
| `word_template` / `word_to_pdf` / `pdf_watermark` / `pdf_encrypt` | `kind="generate"`（生成文档）+ 条目上的勾选项 `to_pdf` / `watermark` / `encrypt` |
| `pdf_source` / `static_file` | `kind="attach"`（附加文件），`params.files` 是路径列表 |
| `mapping.mode`（输出粒度） | `granularity`（条目级：per_lp / per_group / all） |
| `mapping.source`（数据来源） | **概念消解**：来源下沉到字段级，`mapping = {占位符: 来源键}` |
| `mapping.sheet` | 条目级 `sheet = {path, key_col, match_col}` |

新条目结构：

```python
{"id": "it-xxx", "kind": "generate" | "attach", "name": "季度对账单",
 "enabled": True,
 "params": {"template_path": "", "files": [], "to_pdf": False,
            "watermark": False, "encrypt": False, "permissions": ["print"]},
 "granularity": "per_lp" | "per_group" | "all",
 "mapping": {"{{Call_Amount}}": "fields.Call_Amount"},
 "sheet": {"path": "", "key_col": "", "match_col": "name"}}
```

核心层新接口：`normalize_item` / `new_item` / `item_uses_sheet` / `execute_item` /
`run_items` / `field_source_options` / `item_placeholders` / `item_output_name` /
`_item_units_for` / `_item_output_key` / `_item_preview_files` / `default_word_batch`。
**删除**：`attach_mode` / `upstream_lookup_key` / `step_label` / `STEP_TYPES` / `STEP_LABELS` /
`MAPPING_MODE_LABELS` / `OUTPUT_MODE_LABELS` / `DATA_SOURCE_LABELS`。

处理顺序固定为 **转 PDF → 加水印 → 加密**，用户不用排。

### 14.2 界面改造（工作台从 5 步变 6 步）

| 步骤 | 变化 |
|---|---|
| (1) 收件人 | 不变；「发送模式」加括号说明「决定几封邮件」 |
| (2) 文档流水线 | **步骤卡 → 条目表格**（# / 条目名称 / 类型 / 来源·模板 / 处理 / 产出粒度 / 启用）；单击行下方参数区切换；按住行上下拖动改顺序（带插入线）；Ctrl/Shift 多选 → 「删除选中」批量删；双击条目名改名；点「启用」列切换启用。**每行右侧的 ↑↓/参数/删除 四个按钮已删** |
| (3) 附件 | 降级为快捷入口（文案改为「附件（快捷挂载）」） |
| (4) 邮件模板 | 不变 |
| (5) **总览（新增）** | 顶部四项统计（收件人/发送模式/条目数/待生成草稿）+ 预检结论；左表逐人列出最终附件包与打开口令；右侧邮件预览（主题 + 正文渲染图）+ 完整附件清单 |
| (6) 预览与生成 | 原第 5 步，索引由 4 改为 5（已抽常量 `WB_STEP_RUN` 等，避免写死数字） |

- 条目参数区：名称/改名、类型、Word 模板（选择…/检测占位符）或文件列表（添加文件…/从 Excel 导入路径…/保存）、
  处理勾选、**权限位**（允许打印/允许编辑/允许复制内容）、产出粒度下拉。
- 字段映射表三列：Word 占位符 / 填什么（数据来源）/ **预览值（以第一位收件人为例）**；
  双击改来源；来源候选 = 内置 + 收件人字段 + 事件常量 + 外部数据表列（只有真选了 sheet. 来源才出现外部表配置区）。

### 14.3 权限位可配置（Q1）

`watermark_encrypt.encrypt_pdf(..., permissions=None)` 新增 `permissions` 参数，
`permission_mask()` 把 `print/edit/copy` 映射到 PyMuPDF 权限位。**默认仍是「仅打印」**（铁律默认不变），
用户主动勾才放开。全部取消 = 收件人只能打开查看（`normalize_item` 保留空列表，
区别于「没配过 permissions 键」→ 默认 `["print"]`）。

### 14.4 实施中发现并修掉的真缺陷

1. **`tk.StringVar` 被 GC 后 Entry 变空**（新参数区实测截图发现：名称框是空的，而 LabelFrame 标题正确）。
   根因：`tk.StringVar.__del__` 会把 Tcl 变量一起删掉；控件还活着，值没了。
   修法：新增 `_wb_keep_var()` 把参数区所有 tk 变量吊在 `self._wb_keep` 上，
   重建参数区时统一清空。**已加回归测试**（GC 前后 Entry 值不变）。
2. **两个「生成文档」条目用同一个 Word 模板时产物同名**（附件包里撞名）。
   修法：`item_output_name(..., orig_name=...)`；生成条目的「原件名」改用**条目名**
   （`季度对账单_甲公司.pdf` 而非 `对账单模板_甲公司.pdf`），预览清单同步。**已加回归测试**。
3. **执行完成后跳错页**：`_poll_queue` 里 `run_done` 分支写死 `_wb_show_step(4)`，
   v3.1 加了总览页后 4 变成「总览」。改为 `WB_STEP_RUN` 常量（并把 6 个步骤索引全抽成常量）。
4. **`run_items` 空流水线的错误不是元组**：`[("流水线为空…")]` 其实是字符串列表，
   调用方按 `(unit_key, item_id, msg)` 取 `errs[0][2]` 会拿到第 3 个字符。改为真元组。
5. **`granularity="all"` 时加密拿不到口令**：`build_units` 的 all 单元 `rep=None`，
   `_passwords_for` 取到空口令 → 报「收件人没有打开口令」。修法：all 单元回退取第一位收件人。
6. **Word 进程泄漏（机器级问题）**：`run_event` 每次都新建 `WordBatch`，而 `close()` 只释放
   COM 引用、不调 Quit（Quit 会触发 0x800706be 致命异常杀进程），于是**每次生成残留一个
   WINWORD.EXE（约 170MB）**。实测本机堆积到 **71 个**，空闲内存从 7.6GB 掉到 4.1GB，
   机器被拖垮（截图脚本直接卡死）。修法：`lp_v3.default_word_batch()` 按线程缓存共享实例 +
   GUI 改成**常驻工作线程**（`_submit_job` / `_worker_loop`），一个进程最多留 1 个 Word 进程。
   实测：pytest 一轮后残留从 15+ 降到 **1**。
7. **`ITEM_KIND_DESC` 漏导入**导致参数区渲染 NameError（测试抓到）。
8. `run_items` 里 `outputs` 的注释仍写 `step_id`，已改为条目语义。

### 14.5 本轮验证

- `python -m pytest lp-doc-mailer/tests -q` → **132 passed，EXITCODE=0**（v3.0 时 115；
  新增 17 条覆盖条目模型、字段级来源、权限位可配、粒度、Word 复用、GC 变量保活、撞名修复等）。
  **v3.1.1 修复后 → 152 passed**（再新增 20 条 `test_v3_attach_paths.py`）。
- `sync_release.py` → 同步 + 校验通过（`APP_VERSION` = **3.1.0**，三处源码一致）。
  **版本号决定（v3.1.0 当时）**：用户明确指出「不是 3.1.0 吗」——本轮就是 v3.1，设计稿/界面/交接文档
  口径都叫 v3.1，且 3.0.0 与 3.1.0 **都从未发布过**（仓库 tag 最高 v2.2.0），
  所以对外的第一个 3.x 版本直接叫 **3.1.0**，不用再跳一个小版本。
  **后续**：v3.1.0 发布后用户实测发现缺陷（§15），按用户决定 **保留 v3.1.0 不动**，
  修复版另发 **v3.1.1**（`APP_VERSION` 六处同步为 3.1.1）。
- 视觉验收：`v31_library.png` / `v31_pipeline.png` / `v31_overview.png`
  （截图用 Win32 `PrintWindow(PW_RENDERFULLCONTENT)` 抓窗口，不依赖桌面可见——
  本机当前整屏 `ImageGrab` 抓出来是纯黑，会话不可见时整屏截图不可用）。

### 14.6 已知折中与遗留

- **未做旧配置迁移器**（用户明确「不用考虑旧配置，全新开发」）。`normalize_item` 遇到未知
  `kind` 一律当 `generate`；v2→v3 迁移只产出一个「生成文档」条目（带 to_pdf）。
- 附件页（步骤 3）与流水线的「附加文件」条目功能有重叠，已按设计稿把前者降级为快捷入口，
  文案里写清区别（收件人表里的「附件路径」= 每人不同；流水线「附加文件」= 一份文件发多人，
  除非粒度设成「每人一份」只是文件名带名字）。
- 总览页表格列宽有限，长附件名会被截断 → 右侧已加「完整附件清单」文本框兜底。
- **Word 进程仍会残留 1 个/进程（未彻底解决）**：本轮把「每次生成残留 1 个」压到
  「每个 App/脚本进程残留 1 个」，但进程退出后那个 WINWORD.EXE 仍不会自己退出
  （`WordBatch.close()` 只释放 COM 引用、不调 Quit——调 Quit 再释放引用会触发
  0x800706be 致命异常直接杀进程，这条路已验证是雷）。
  残留是无头实例（MainWindowTitle 为空、不可见、不含用户文档）。
  清理办法（只杀无头的，别杀用户自己开着的 Word）：
  ```powershell
  Get-Process WINWORD -EA 0 | ? { [string]::IsNullOrEmpty($_.MainWindowTitle) } | Stop-Process -Force
  ```
  实测本轮堆积到 71 个时机器被拖垮（空闲内存 7.6GB → 4.1GB，截图脚本直接卡死）；
  清理后回到 7.4GB。要彻底解决需要另找 Word 退出路径（风险高，建议单独一轮处理）。
- `.zcode` 目录下的临时脚本（`shot_v4.py` / `verify_v4_features.py` / `verify_concepts.py` /
  `_rename_step*.py` 等）因本会话回收站仅限工作区内路径而无法删除，可随手删。
  （v3.1 的 `shot_v31.py` + `win_grab.py` 是有意保留的：可复现界面验收。）
- **P4（收件人库界面、执行历史查看界面）仍未做**，发版说明要写清。
- v3.1.0 的打包 / 签名 / git tag / Release 上传 **已执行**（2026-10-23 晚，见 §15.3）。
- 仓库状态（`lp-doc-mailer/`，branch=main，remote=github.com/fundops-notes/doc-mailer）：
  v3.1.0 已 push（main = `0c7d139`，tag `v3.1.0`）；v3.1.1 的改动见 §15.4。

---

## 15. v3.1.1 修复版：「附件路径」链路（2026-10-23 深夜）

### 15.1 触发原因

用户在 v3.1.0 实测后反馈：**导入模板里有「附件路径」列，但界面上没有维护的地方，
流水线里也看不到每个收件人的文件地址配置。**

排查后确认是**三个问题叠在一起**（不是单纯"界面没做"）：

| # | 问题 | 位置 | 后果 |
|---|---|---|---|
| ① | 收件人表**没有「附件路径」列** | `lp_tool.py` 收件人表格列写死 | 路径其实导进来了（`HEADER_ALIASES["files"]` → `rec["attachments"]`），但用户在表格上看不见、也没法在表格上维护；编辑框里又叫「专属附件」，与模板列名对不上号 |
| ② | 流水线**没有 per-recipient 路径这一层** | `execute_item()` 的 attach 分支只读 `params["files"]` | 附加文件条目只能手填一份固定清单，所有人拿一样的；per-recipient 路径只在最后组装附件包时被悄悄并进去，流水线里既配不了也看不见 |
| ③ | **配了附件路径就点「导出收件人清单」必崩** | `lp_v3.recipients_to_rows()` | `attachments` 是**路径字符串列表**，代码却按 dict 用（`a.get("path")`）→ `AttributeError`；exe 是 `--noconsole` 打包的，异常被 tkinter 吞掉 → **用户点按钮后连报错都看不到，什么都不会发生** |

### 15.2 改了什么

**`lp_v3.py`**
- `recipients_to_rows()`：attachments 兼容「字符串 / dict」两种形态，不再崩。
- `normalize_recipient()`：v2 的 `files` 与 v3 的 `attachments` 合并后去重。
- 新增**附加文件条目的「路径来源」**（与占位符的字段来源同一套口径）：
  - `manual`（默认，行为与旧版完全一致）/ `recipient.attachments`（收件人清单的「附件路径」列）
    / `fields.<列名>`（收件人自定义字段）/ `sheet.<列名>`（外部数据表列）
  - `path_source_options()` 列出可选项，`path_source_label()` 出人话，`resolve_item_paths()` 解析取值。
- `execute_item()` attach 分支：按来源取路径。按人取时某人没填/文件不存在 → **warning 跳过**，不让整条失败；
  手填路径不存在仍然报错（配置错误要暴露）。
- `event_uses_recipient_paths()`：若流水线已有条目按「收件人清单附件路径」取，
  `assemble_attachments()` 就不再原样重复附加收件人那列路径 —— **否则同一个人会收到两份同样的文件**
  （一份是条目产物、一份是原始路径）。
- `item_uses_sheet()`：路径来源是 `sheet.` 时也算"用到外部数据表"（否则不会去读那张表）。

**`lp_tool.py`**
- 收件人表新增「附件路径」列，显示每人各自的路径（文件名）；**文件不存在时那个名字前加 `⚠`**。
- 编辑收件人对话框的标签由「专属附件」改为「**附件路径**」，与导入模板列名一致。
- 附加文件条目参数区新增「**路径来源**」下拉 + 预览（"以第一位收件人为例会取到什么"）；
  选中非手填来源时，手填文本框与三个按钮一起禁用（disabled 的 Text 再 insert 会抛 TclError）。
- 流水线表格的「来源 / 模板」列按路径来源显示（不再显示手填清单）。
- **导出收件人清单加了异常兜底**：出错弹「导出失败」框，不再静默失败。
- 帮助页写清三处附件来源的区别。

**模板 / 文档**
- 导入模板表头改为与 `EXPORT_BASE_HEADERS` 一致（**补上「合并组」列** —— 导出表头本来有、
  模板没有，说明页却写着它）；「附件路径」说明写清"每人专属、会进这个人的附件包"。
- `examples/收件人清单导入模板.xlsx` 重生成（含「填写说明」页），
  `examples/收件人清单示例（演示数据）.xlsx` 补「合并组」列。
- 两份《使用说明.txt》补充「路径来源」与三处附件来源说明。

### 15.3 验证

- `pytest lp-doc-mailer/tests -q` → **152 passed**（新增 `test_v3_attach_paths.py` 20 条），
  `LP_SKIP_WORD_TESTS=1` 模拟 CI → **143 passed + 9 skipped**，EXITCODE=0。
- 实机验证脚本 `shot_attach_paths.py`（可复现）：收件人页「附件路径」列显示各自路径 + 缺失标 ⚠；
  流水线参数区「路径来源」下拉含「收件人清单的『附件路径』列」；导出表第 7 列正确写出路径。
  截图 `v31_recipients.png` / `v31_pathsource.png`。
- 打包：`doc-mailer_v3.1.1_portable.zip`（44.2 MB / 2550 文件，zip 内断言 `APP_VERSION = "3.1.1"`）；
  `doc-mailer_win64.zip`（44.7 MB，用**本地已签名** exe，zip 内 exe 签名 Valid、PYZ 里版本 3.1.1）。

### 15.4 版本策略（用户拍板）

v3.1.0 的 tag 与 Release 在修复前已由 CI 建出（`v3.1.0` → `0c7d139`，
Release id=421973099，资产是 CI 打的**未签名** win64.zip）。
用户选择：**保留 v3.1.0 不动**，修复版另发 **v3.1.1**（`APP_VERSION` 六处同步为 3.1.1）。
→ v3.1.0 的 Release 说明里加一句指向 3.1.1 的提示。

### 15.5 本轮踩到的新坑

- ⚠️ **PyInstaller onefile 会派生子进程**：`verify_exe_boot.py` 里 `proc.terminate()` 只杀父进程，
  子进程继续活着并**锁住 `dist/LP_Tool.exe`** → 下次打包 `PermissionError: [WinError 5]`。
  已在 `verify_exe_boot.py` 里改成"按可执行文件路径"用 PowerShell 一并清残留。
- ⚠️ **exe 体积**：本机 PyInstaller 会把 `PIL.Image` 里函数体内的 `import pyarrow/numpy` 也算进依赖，
  于是 pandas/pyarrow/numpy/scipy 全被打进 exe（**115 MB**）。CI 是干净 venv 所以只有 44 MB。
  已在构建命令里加 `--exclude-module pandas pyarrow numpy scipy IPython matplotlib polars win32com.test`，
  exe 回到 **45.1 MB**。依据：本项目源码无 numpy 直接依赖、`openpyxl` 那边是 try/except、
  便携版 runtime 本来就没装 numpy（一直跑得好好的）。
- ⚠️ **exe 字节码里搜不到版本号字符串**是正常的（PYZ 是 zlib 压缩的）；
  要验证包里版本号得用 `PyInstaller.archive.readers` 取出模块字节码看常量。
- ⚠️ **发布 Release 资产必须发到 `uploads.github.com`，发 `api.github.com` 一律 404**：
  ```
  POST https://uploads.github.com/repos/{owner}/{repo}/releases/{id}/assets?name=...
  ```
  其它 release 操作（建/改/删 Release、删资产、读资产）发 `api.github.com` 都正常。
  症状极具迷惑性：同一枚 token 能建 Release、能改说明、能删资产，唯独「新增资产」一律
  `404 {"message":"Not Found"}`；不带 `name` 也是 404（正常应是 422），说明**端点根本没被授权**，
  不是参数问题。排查顺序：换 host → 换 header 形态 → 换仓库 → 换 token → 试草稿 Release。
- ⚠️ **CI 会自己传一个未签名的 `doc-mailer_win64.zip`**：CI 用 `GITHUB_TOKEN`（权限足），
  本地这枚 GCM token 也传得动资产（只要走对 host）。本轮用**本地已签名** exe 自建的 win64.zip
  替换了 CI 那份（先 DELETE 同名资产再上传）。便携包只能人工传 —— CI 里没有绿色版的
  121MB runtime，造不出来。

### 15.6 v3.1.1 发版结果

- commit `63d5565`（v3.1.1）已 push main；tag `v3.1.1` → `63d5565` 已推远端。
- CI run `39977597018` **success**（Run tests / Build exe / Package / Upload to Release 全绿）。
- Release `v3.1.1`：id=**422000210**，已发布（非草稿、非预发布）。
  - `doc-mailer_v3.1.1_portable.zip`（46,384,250 字节）
  - `doc-mailer_win64.zip`（46,829,079 字节，**本地已签名** exe 打包）
  - 匿名 HEAD 下载地址 → 302（= 公开可下）。
- Release `v3.1.0`（id=421973099）**原样保留**（资产仍是 CI 的未签名 win64.zip），
  但说明正文已改成开头就提示「请升级到 v3.1.1」。
- 产物均已读回校验：便携 zip 内 `APP_VERSION = "3.1.1"`、无 `lp_data/`、无 `__pycache__/`；
  win64.zip 内 exe 签名 Valid、PYZ 里版本常量 3.1.1、**没有** pandas/pyarrow/numpy。

---

## 16. v3.2.0 P4 收尾：收件人库 + 执行历史（2026-10-23 深夜 ~ 2026-10-24）

### 16.1 范围

v3.0 设计稿 §7 把 P4 定为三件事：**收件人库（跨事件复用）、执行历史查看、示例数据更新**。
P0–P3 已铺好数据底座（`lp_data/runs/` 有 Run 记录），但**一直只有数据、没有界面**。

**关键约束（用户 2026-10-23 反馈②）**：用户明确说过「场景与事件 / 场景模板 / 收件人库
三个入口看不懂区别，只要事件库」，当时把收件人库那个空壳页整个删了。
→ 所以本轮**不复活独立导航页**，收件人库做**事件库里的第三个标签页**，
执行历史做**第四个标签页**，顶层导航仍然只有「事件库 / 设置 / 帮助」三项。

设计稿：`docs/设计稿_v3.2_P4收尾.md`。

### 16.2 数据层（`lp_v3.py`）

| 新增 | 说明 |
|---|---|
| `RECIPIENTS_DIR = "recipients"` + `Store.recipients_dir` | 库文件落在 `lp_data/recipients/<id>.json`，与 events/scenarios/runs 同级 |
| `normalize_pool_recipient()` / `new_pool_recipient()` | 规范化 + 新建（id 前缀 `rcp-`，格式 `rcp-YYYYmmdd-HHMMSS-xxxxxx`） |
| `pool_merge_recipient(pool, rec)` | 按名称并入库，返回 `(pool, **新增的那条记录 or None**, 是否更新)` |
| `merge_recipient_into_list(list, rec)` | 铁律实现：按名称匹配，**只补空缺**，**绝不覆盖已有打开口令与权限口令** |
| `merge_recipients_into_event(event, records)` | 把若干收件人并入事件，返回 `(event, 新增数, 更新数)` |
| `Store.list_pool()` / `save_pool_recipient()` / `delete_pool_recipient()` | 库存取；损坏文件**改名留证**不删除（与 events 同策略） |

> ⚠️ `pool_merge_recipient` **返回「新增的那条记录」而不是计数** —— 调用方需要知道该存哪条。
> 早先一版返回计数，导致调用方误存 `pool[-1]`（库里最后一条 ≠ 刚合并的那条）。已改。

### 16.3 界面（`lp_tool.py`）

- 抽出共用表单 `_recipient_dialog(rec, field_cols, title, on_save, no_field_hint)`，
  事件里的「编辑收件人」与库里的「新增/编辑」**共用同一套表单**（字段、口令校验完全一致）。
- **事件库第三个标签页「收件人库」**：
  - 列：收件人名称 / To / CC / BCC / 水印 / 组 / **附件路径** / 自定义字段列 / 打开口令 / 权限口令(勿发)。
    缺失文件名前标 `⚠`；权限口令在表里显示 `••••`。
  - 工具条 8 键：新增 / 编辑 / 删除 / 从事件导入… / 添加到事件… / 导入 Excel-CSV / 导出 Excel / 下载导入模板。
  - 底部红字安全提示（库里含口令，与 `lp_data/` 同管控级别）。
- **事件库第四个标签页「执行历史」**：
  - 列：开始时间 / 事件 / 状态 / 草稿 / 成功 / 失败 / 跳过 / 附件包。
  - 双击出详情弹窗：每位收件人的附件包数、缺失、草稿名、错误，以及执行日志 / 错误 / 提醒。
  - 「打开该事件」按钮；执行记录**只看不删**（审计材料）。
- **事件工作台「收件人」步骤**工具条新增「**从收件人库添加…**」「**存入收件人库**」。

### 16.4 示例数据

- `lp_data.example.json` 新增 `pool` 段（2 条虚构收件人，口令空则自动补随机）。
  四份副本同步：项目根 / `LP_Tool_绿色版/` / `lp-doc-mailer/examples/` / `LP_Tool_分发包/`。
- `seed_demo()` 播种**示例收件人库**（读 `pool` 段）。
- 两种产出条目的示例（generate + attach 各一）放在**场景模板**里且**默认禁用**；
  **事件流水线保持为空** —— 预检对「流水线为空」和「未选模板」都判 error，
  放进事件会挡住用户第一次运行；禁用的条目在场景里不参与预检，所以安全。

### 16.5 本轮顺带修掉的真实缺陷

1. **`save_run` 静默丢执行记录（真缺陷，跨版本）**：
   `run["id"]` 未清洗，里面带 `:` `/` `\` 等 Windows 文件名非法字符 → 写文件失败，
   而**调用方忽略返回值** → 执行记录**无声消失**。修法：`sanitize_filename(run.get("id"))`；
   并且 `run_event()` 两处 `save_run` 失败会写进 `run["warnings"]` 并记日志，不再吞掉。
2. **界面文案里漏出 Markdown 记号**：Tk 标签不支持 `**加粗**`，原样显示成 `**……**`。
   修掉 3 处（收件人库说明、从收件人库添加说明、「确认生成」提示框）。
3. **`pack_portable.py` 版本号写死**（写死过一次就漏发过一次）→ 改成从 `lp_tool.py`
   的 `APP_VERSION` 读，并新增 P4 关键函数断言。

### 16.6 验证（全部实测）

- `sync_release.py` → 同步 + 校验通过（`APP_VERSION` = **3.2.0**，四处一致）。
- 全量 pytest：**191 passed**（新增 `test_v3_pool.py` 21 条 + `test_v3_pool_gui.py` 17 条）。
- CI 模拟（`LP_SKIP_WORD_TESTS=1`）：**182 passed, 9 skipped**。
- 实机截图验收脚本 `shot_p4.py`（新增，固化在项目根）：四张截图
  `v32_pool.png` / `v32_runs.png` / `v32_run_detail.png` / `v32_wb_recipients.png`，
  逐条断言四标签页、附件路径 ⚠、权限口令打码、执行历史计数、详情弹窗、工作台两个新按钮。
- 便携包已重打包并读回断言：`doc-mailer_v3.2.0_portable.zip`（2550 文件 / 44.2 MB），
  内含 `APP_VERSION = "3.2.0"` 与 P4 关键函数、无 `lp_data/`、无 `__pycache__/`。

### 16.7 版本与发版状态

- 对外版本号定为 **3.2.0**（新增功能，不是修复）。
- **v3.2.0 尚未发版** —— 用户本轮只说「继续做完 P4」，没说发版，故未擅自发布。
  代码、测试、便携包都已就绪；发版时按 §7 流程走（tag `v3.2.0` 会触发 CI）。
- 注意：本地 `APP_VERSION` 已是 3.2.0 > 远端最新 Release 3.1.1，
  自动更新检查会显示「当前已是最新版本（v3.2.0）」，**不会**误提示有新版（逻辑 `remote <= local`）。

### 16.8 本轮改动的文件

| 文件 | 改动 |
|---|---|
| `lp_v3.py` | 收件人库数据层、`save_run` 修复、`seed_demo` 播种库与示例条目 |
| `lp_tool.py` | 收件人库/执行历史两个标签页、共用收件人表单、工作台两个新按钮、文案修正、版本号 |
| `lp_data.example.json`（×4 份） | 新增 `pool` 段 |
| `tests/test_v3_pool.py` | **新建** 21 条（规范化 / 合并铁律 / 库存取 / 损坏隔离 / seed_demo / 执行历史排序过滤） |
| `tests/test_v3_pool_gui.py` | **新建** 17 条（收件人库页渲染 / 执行历史页渲染 / 工作台互操作） |
| `tests/test_v3_gui.py` | 标签页断言 2 → 4，并断言顶层导航仍只有三项 |
| `docs/设计稿_v3.2_P4收尾.md` | **新建**（P4 设计） |
| `docs/使用说明.txt` | 版本号 + 新增「收件人库」「执行历史」两节 |
| `shot_p4.py`（项目根） | **新建** 实机验证 + 截图脚本 |

### 16.9 已知限制与后续

- 收件人库**不做双向自动同步**：只有显式的「存入 / 添加」按钮，避免意外改动已发批次。
- 收件人库与执行历史**只在事件库内**，没有独立入口（尊重用户已否掉的方案）。
- 执行历史**只读**，没有清理入口；`lp_data/runs/` 会随时间变大（当前无归档策略）。
- 下一轮若发版：按 §7 走 → 重出 exe（签名）→ 两个 zip → commit + tag `v3.2.0` → CI → 上传资产 → 写 Release 说明。

---

### 16.10 v3.2.0 发版结果（2026-10-24）

- 发版前校验全绿：`sync_release.py` [OK] / `pytest tests -q` **191 passed** / `pack_portable.py` 读回断言全部 [OK]。
- exe：`--exclude-module`（pandas/pyarrow/numpy/scipy/IPython/matplotlib/polars/win32com.test）→ **45.1 MB**（目标区间内），签名 **Valid**。
- win64 包：3 项在根（与 v3.1.1 布局一致），zip 内 exe 签名 Valid、`lp_data.example.json` 含 `pool` 段（读回断言通过）。
- tag `v3.2.0` → `1764a73` 已推远端（**本轮 github.com 直连可达，未动 FlClash**）；
  远端 main 同为 `1764a73`。
- CI run `39989823347` **success**，自动创建 Release id=**422091489** 并上传未签名 win64.zip。
- `release_tool.py upload 3.2.0`：先删 CI 的同名未签名资产再上传两个本地资产（win64 为已签名版顶替）。
- 终态验证：Release 非草稿非预发布；资产 2 个
  （`doc-mailer_v3.2.0_portable.zip` 46,392,519 字节 + `doc-mailer_win64.zip` 46,845,800 字节）；
  匿名 HEAD 均 302；说明正文为 v3.2.0 版、无旧包名残留。
- ⚠️ 本轮网络注意：FlClash mixed-port 已关（7890 无监听），但**注册表系统代理仍指向 127.0.0.1:7890 且 ProxyEnable=1**
  → Python urllib 会捡到死代理而 `connection refused`。绕行：命令前加 `no_proxy="*"`（不动用户系统设置）。
  curl 不读注册表代理，直连正常。此状态与 §17 描述略有差异（§17 是 github.com 被屏、API 通；
  本轮是 github.com 也通了，但系统代理残留导致 urllib 全挂）。

---

## 17. 网络环境坑：github.com 被 IP 级屏蔽（2026-10-24 实测）

### 17.1 症状

`git push` / `git tag` 报：

```
fatal: unable to access 'https://github.com/fundops-notes/doc-mailer.git/':
Failed to connect to github.com:443 after 21097 ms: Could not connect to server
```

**但 API 是通的** —— `api.github.com`、`codeload.github.com`、`uploads.github.com` 全部可连，
只有 `github.com` 连不上。

### 17.2 根因（不是 DNS 污染，也不是代理没开）

| 域名 | 解析结果 | TCP 443 |
|---|---|---|
| `github.com` | 20.205.243.**166** | **超时** |
| `api.github.com` | 20.205.243.**168** | 通 |

同一 /24，只差最后一位 IP —— **IP 级封锁**，不是域名级，也不是 DNS 问题
（换 FlClash 的 DNS 1053 解析，结果一样）。

### 17.3 绕行办法（FlClash 核心 + 控制器 API）

本机 FlClash 核心在跑，但 `mixed-port` 被置 **0**（只做 DNS/规则，没开代理口），
系统代理注册表 `ProxyEnable=0`（`ProxyServer=127.0.0.1:7890`）。

**临时开一个本地代理口**（用 Clash 控制器 API，secret 在
`%APPDATA%\com.follow\clash\config.yaml` 的 `secret:` 字段）：

```powershell
$sec = (Select-String "$env:APPDATA\com.follow\clash\config.yaml" -Pattern '^secret:\s*(.*)$').Matches[0].Groups[1].Value.Trim().Trim('"')
$hdr = @{ Authorization = "Bearer $sec"; 'Content-Type' = 'application/json' }
Invoke-RestMethod 'http://127.0.0.1:9090/configs' -Method Patch -Headers $hdr -Body '{"mixed-port":7890}'
```

然后带代理推：

```powershell
git -c http.proxy=http://127.0.0.1:7890 -c https.proxy=http://127.0.0.1:7890 push origin main
```

**推完立刻恢复**（别留着一个用户没打算开的代理口）：

```powershell
Invoke-RestMethod 'http://127.0.0.1:9090/configs' -Method Patch -Headers $hdr -Body '{"mixed-port":0}'
```

### 17.4 影响面

- `git push` / `git tag` / `git push --tags` → 走 `github.com`，**需要代理**
- `release_tool.py` 走 `api.github.com` + `uploads.github.com` → **不需要代理**（直连可用）
- GitHub Actions CI → 跑在 GitHub 自己的机器上，**不受影响**
- 其它 Release 操作（建/改/删 Release、删资产）→ `api.github.com`，不需要代理

### 17.5 顺带记录：工具脚本不再写死版本号

`pack_portable.py` 与 `release_tool.py` 的版本号都改成从 `lp_tool.py` 的 `APP_VERSION` 读
（**写死过一次就漏发过一次** —— v3.1.1 那次就差点打包成旧版本）。



## 18. v3.3.0：附件内容级去重（2026-10-24）

### 18.1 背景（用户拍板）

用户确认发送语义：有合并组 → 同组一封（共用模板、附件并入、收件人去重、**附件按内容哈希去重**）；
没组 → 各发各的；不同组 → 不同邮件。实测确认  现有语义已完全符合
（无组收件人自成一组 = 各发各的），唯一差距是附件去重只按路径、不按内容。

### 18.2 改动

- ：
  去重升级为两层——先按路径（大小写不敏感），再按内容指纹（**大小 + SHA256**）；
  同内容不同文件名的副本只保留第一份； 传入 list 时记录 (保留, 被丢弃)。
- ：内容指纹辅助（读不了的文件返回 None，不参与内容去重）。
-  组装附件包时传  并经  写日志：
  。
- **兼容**： 可选，旧调用不受影响；收件人去重（To/CC/BCC 保序去重）与
  合并组语义（：无组 = §+名字 自成一组）本来就对，未动。

### 18.3 验证

- 新增  10 条：混合分组语义（G1 合并/无组各发/G2 合并 = 4 单元）、
  地址去重、同内容不同名只留一份、不同内容都保留、同路径一次、跨来源（静态+专属）哈希合并、
  旧签名兼容、指纹稳定性。
-  的测试数据原来三个文件
  **内容全部是 b"1"**——在新语义下被正确合并成一份而失败；已改为三个不同内容
  （该测试本意是验证 scope 匹配，不是去重）。**这是新语义生效的证据，不是回归。**
- 全量 **201 passed**；真实  实测输出去重日志且 G1 单元只挂一份。
- 版本号两处（ + ——**这次差点又漏掉 lp_v3 那处**，收集阶段断言拦住了）
  与使用说明已同步 3.3.0。

### 18.4 发版结果

- Release v3.3.0（id=422474282）： + （签名版顶替 CI），
  匿名 HEAD 302，CI run 40055394494 success。


## 19. v3.3.1：指纹提示改为只提示不去重（2026-10-24 第二轮用户拍板）

### 19.1 决策变更

v3.3.0 做的是「同内容自动只留一份」。用户澄清：**不要去重、不改路径**——
两份同内容附件都发出去，但在「检查附件」时对同一封邮件内不同路径、指纹
（大小+SHA256）相同的文件做特别提示，由人决定是否有意如此。

### 19.2 改动

- assemble_attachments：内容指纹重复【全部保留】，dup_log 只记 (首份, 重复份)；
  路径级去重保留（同一路径=同一文件只挂一份）。
- 新增 lp_v3.attachment_fingerprint_warnings(event)：按草稿单元（=每封邮件）扫
  静态附件+收件人专属附件（不含流水线产物——每人一份的产物内容天然不同），
  返回提示文案。接入三处：preflight warnings、工作台两个「检查附件」按钮
  （收件人步骤与静态附件步骤，弹窗里带【内容重复提示】段）、run 日志
  （[附件重复提示] … 已全部保留，请人工确认）。
- 分组设置唯一性核实：合并组只在收件人一处设置（编辑框/Excel列），
  发送模式与条目粒度只引用不重设——不存在重复设置问题。

### 19.3 验证

- 测试重写为 warn-only 语义（14 条）：同内容不同名两份都保留+记录、同路径一份、
  不同内容无提示、同封才提示（跨封不提示）、跨来源（静态+专属）提示、
  preflight 接入。全量 204 passed。
- 真实 run_event：同内容三来源（流水线产物副本+专属×2）→ 全保留 +
  3 条提示日志；G1 单元挂 2 份（用户文件）符合预期。

### 19.4 发版

Release v3.3.1（id=422486899），两资产 302，CI run 40057711228 success。


## 20. v3.3.2：空名导入行可见化（2026-10-24 第三轮）

### 20.1 用户要求

「名称为空就不导入」的逻辑**不变**；只要求导入结果显示时能看到这一行
（红/绿提示足够），并提示会导入多少条。

### 20.2 改动

- validate_recipient_rows：parse 阶段被丢弃的空名行【补回结果表】——
  status=error（红色），错误文案「名称为空（该行不会导入）」，计入 error_count；
  完全空白的行与「示例：」模板占位行仍静默跳过（无意义/有意为之）。
- 行号修正：丢行后 enumerate 会错位 → 改为按原始行号对齐（kept_nos）。
- 导入结果弹窗汇总新增「将导入 N 条（通过+警告）」。
- 导入逻辑零改动（apply 的 only_ok 与空名 guard 本来就跳过这些行）。

### 20.3 验证

新增 4 条测试（空名行可见+计数、全空行不显示、示例行不显示、行号对齐）→ 全量 **208 passed**。
实测：表头/甲/空名/全空/乙/示例 六行 → items 3 条 = 行2甲(ok) 行3空名(error) 行5乙(ok)，行号正确。

### 20.4 发版

Release v3.3.2（id=422539456），两资产 302，CI run 40066681775 success。


## 21. v3.4.0/v3.4.1：发件人视角重构（2026-10-24 深夜）

### 21.1 用户 8 项需求全部落地

① 收件人页（表格/编辑框/Excel 映射）不再出现「附件路径」——附件统一到流水线。
② 「生成文档」改名 **Word-MailMerge**：条目参数区有「Word-MailMerge 说明」按钮
（占位符写法/处理顺序/水印口令来源一页讲清）；加密勾选后新增**「发送人统一密码」**
选项（勾选+输入框，即 GP 统一权限口令机制，方便发送人打开核对）。
②b 预设占位符 **Client_Name**（=收件人名称）进字段来源列表。
③ 「产出粒度」全部改为「产出颗粒度」。
④⑤ 流水线页新增：**统一附件区**（所有人同一份文件，替代「附件」页功能）、
**▶ 处理按钮**（按当前参数全量重跑、覆盖上次结果）、
**收件人↔文件对应表**（每人将收到的文件+路径+状态，处理完成自动回填；
Run items 按组拆回成员、按收件人清单顺序去重）。
⑥ 事件表/流水线表横向滚动条、root.minsize(960,620)。
⑦ 默认正常窗口 1280x800，不再最大化。
⑧ 自绘 icon.ico（信封+文档+锁，6 尺寸），exe --icon 挂载，窗口 iconbitmap。

### 21.2 过程中修掉的 4 个真缺陷

1. **normalize_event 白名单洗掉 process_results** → 处理后对应表永远是空。
   白名单加入 process_results。
2. **run_done 消息 tuple 化后旧分支 .get 崩** → 分支改 tuple 安全。
3. **_apply_process_results 用 self._wb_event** → 工作台重建后为 None；
   改为 run_done 消息携带 ev。
4. **_wb_run（预览生成路径）漏传 ev** → 补上。
5. src/ 副本不同步导致 CI 测试失败（旧 pool 行构造）→ sync 后重跑 CI 绿。

### 21.3 测试与发版

- 全量 **208 passed**（新增 test_v3_attach_dedup.py 14 条）。
- 真实端到端：Word 模板 → 处理按钮 → 每人一份加密 PDF（口令/权限/水印验证）
  → process_results 逐人回填持久化。
- v3.4.0 CI 失败（src 未同步）→ 修复后以 **v3.4.1** 重新发布
  （Release id=422581991，两资产 302）。
- 版本号升版时注意：**lp_tool.py + lp_v3.py + 使用说明.txt 三处**。
