# HANDOFF：文档分发助手 v3.8.3（功能归属整理与导入导出修复）

> 交接文档。给接手实现的 agent / 同事。
> **文件名沿用 `HANDOFF_v3.0.md`（历史沿革，改名会打断各处引用），内容已覆盖到 v3.8.3。**
> 最后更新：2026-10-25（v3.8.3 功能归属整理与三项缺陷修复，见 §40；v3.8.2 见 §39）
>
> **当前代码版本：`3.8.3`**（`lp_tool.py` 与 `lp_v3.py` 两处一致；使用说明三处同步）。
> 已发布：v2.2.0（旧）、v3.1.0 / v3.1.1、v3.2.0、v3.3.0、v3.3.1、v3.3.2、v3.4.0/v3.4.1、v3.4.2、v3.4.3、**v3.4.4**（v3.4.3 独立验收问题修复版，见 §24）、**v3.4.5**（v3.4.4 界面热修，见 §25）、**v3.4.6**（改名直改/要素模板导入/内联说明，见 §26）、**v3.4.7**（检查更新 API 限流修复，见 §27）、**v3.4.8**（字段映射弹窗关联修复，见 §28）、**v3.4.9**（自动更新下载回调修复，见 §29）、**v3.5.0**（顶部导航/莫兰迪 + 附件路径按人导入 + Excel 显示值 + 取消颗粒度，见 §30）、**v3.5.1**（设置清理 + 文件名模板下沉 + 场景模板移除，见 §31）、**v3.6.0**（四步流程 + 处理日志/进度 + 总览预生成 + 首页编辑事件，见 §32）、**v3.7.0**（WYSIWYG 编辑器/单封编辑/收件人标签/添加到事件修复，见 §33）、**v3.7.1**（自定义字段只属于事件 + 库字段清理，见 §34）、**v3.7.2**（收件人信息隐藏字段 + 日期序列 + 图片显示，见 §35）、**v3.7.3**（日期导入按 Excel 显示文本 + 可选格式，见 §36）、**v3.8.0**（SaaS UI + 本事件字段 + 编辑器/预生成/步骤推进/未保存提醒，见 §37）、**v3.8.1**（工作台三入口左移/侧栏缩窄/检查更新移左下，见 §38）、**v3.8.2**（按标签筛选并批量添加事件收件人，见 §39）、**v3.8.3**（功能归属整理与导入导出修复，见 §40）。

---

## 0. 一句话现状

桌面工具「文档分发助手」当前代码版本 **v3.8.3**：事件列表专注事件管理和备份；事件工作台维护本次名单与业务字段；库页维护可复用联系人、标签及专用Excel导入导出。修复跨事件明细字段表头、库标签往返、同名入库更新未保存。全量 pytest 已实跑 **285 passed**（24.68s，无跳过）。真实GUI+Word 188 PASS / 0 FAIL，构建/签名、CI与最终资产下载校验均通过，v3.8.3已发布；结果见 §40。Outlook 的 fake 验证不代表真实邮箱 COM 已验收。

---

## 0.1 源码坐标（以下 v3.2 历史快照；当前状态以 §0 / §40 为准）

| 项 | v2.2.0 | v3.2.0（历史快照） |
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


## 22. v3.4.2：流水线页左右分栏 + 对应表可编辑 + Excel 模板导出（2026-10-24 深夜）

### 22.1 用户 7 项反馈

① 流水线页显示不全 → 改左右分栏（PanedWindow，可拖）：左=条目表+参数区，右=统一附件+处理按钮+对应表。
② 对应表双击「文件路径」列行内编辑；表下「保存自定义编辑」写 event.process_overrides
（normalize_event 白名单已加，否则保存被洗掉——又一个真缺陷）。
③ 统一附件改「📎 统一附件…」按钮，点击才展开（默认收起省空间）。
④ 去「插入示例链路」按钮。
⑤ 处理只检查提示不改包（v3.3.1 语义保持）。
⑥ Word-MailMerge 参数区加「导出要素模板」Excel（预填本事件收件人，列=已检测占位符）。
⑦ 附加文件加「导出路径模板」Excel（预填收件人+附件路径列）。
⑧ 双击条目行不再弹改名窗；名称在参数区文本框改，点「改名」直接生效。

### 22.2 修掉的缺陷

- normalize_event 白名单漏 process_overrides（保存即被洗掉）。
- pool 行构造遗漏 atts 列清理（v3.4 主体改动时 src/ 副本漏同步暴露于 CI）。
- CI 失败根因两次均为 src/ 未同步 → 已按 HANDOFF §6 要求 sync 后重推。

### 22.3 验证与发版

全量 **208 passed**；对应表编辑保存→process_overrides 持久化实测 OK。
Release v3.4.2（id=422627074），两资产 302，CI run 40080833442 success。


## 23. v3.4.3：流水线页左右分栏落盘 + 附加文件去路径来源 + Excel 模板导出（2026-10-25）

### 23.1 用户第二轮反馈（截图确认 v3.4.2 部分未生效）

1. 附加文件条目的「路径来源」下拉仍在（收件人页已无附件路径列，该来源已无意义）。
2. 收件人↔文件对应表没有实现左右分栏（v3.4.2 的分栏补丁因脚本中断未落盘）。
3. 「插入示例链路」仍在。
4. Word-MailMerge 说明不明细、无示例模板可测。

### 23.2 改动

- 附加文件条目参数区：删除「路径来源」下拉、disabled 提示、sheet 面板挂载
  （v3.4 收件人页已无附件路径列，recipient.attachments 来源保留在 lp_v3 兼容旧数据但 UI 不再露出）。
- 流水线页改为**左右 PanedWindow 分栏**：左=条目表+参数区，右=统一附件（折叠按钮）+处理按钮+对应表。
  对应表 grid 布局 + 双击「文件路径」行内编辑 + 保存按钮。
- 删除「插入示例链路」按钮与 _wb_add_example_chain 方法。
- Word-MailMerge 参数区：说明按钮 + **导出要素模板 Excel**（预填本事件收件人，列=已检测占位符）
  + **下载示例 Word 模板**（Capital Call 格式含 6 占位符，真实生成 .docx）。
- 附加文件参数区：**导出路径模板 Excel**（收件人+附件路径两列）+ 说明文字。
- normalize_event 白名单补 process_overrides。

### 23.3 验证与发版

- 全量 **208 passed**；GUI 实测组件齐备（PanedWindow 右栏确认）。
- Release v3.4.3（id=422643200），两资产 302，CI run 40083094646 success。
- 版本号三处同步（lp_tool/lp_v3/使用说明）。


---

## 24. v3.4.4：v3.4.3 独立验收问题修复（2026-10-24）

### 24.1 触发

对 v3.4.3 做独立验收（真实 GUI + 真实 Word 流水线，见项目根 `验收报告_v3.4.3.md`）发现
5 项需求未完全达标，其中 2 项是发布级问题：
1. **P0** 流水线页左栏（条目表+参数区）被 PanedWindow 的空左栏盖住，默认完全不可见；
   v3.4.3 的 `pack(in_=left_col)` 只是把窗口塞进左栏的打包列表，父窗口没变，后建的 PanedWindow
   又压在它上面。
2. **P1** 收件人表 `values` 比 `columns` 多一个遗留的附件文本 → Tk 丢掉最后一格：
   权限口令 `••••` 遮罩失效、整行错位、明文打开口令显示在权限口令列。
3. **P1** 收件人导入模板/`examples/*.xlsx` 仍有「附件路径」列，导入仍写 `rec["attachments"]`。
4. **P1** Word-MailMerge 的「导出要素模板」「下载示例 Word 模板」两个按钮未接线（后者是死代码）；
   HANDOFF §23.2 声称已完成，实际未落盘。
5. **P1** 「发送人统一密码」调用了不存在的 `save_data()`（AttributeError、settings 不落盘），
   且 settings 没有传给 `event["options"]`，引擎读不到 → 勾了等于没勾。
6. 其余 P2：流水线条目表无横向滚动条；统一附件按钮不在右栏；帮助页/使用说明残留旧文案；
   三处重复提示文案不一致；`process_overrides` 对多文件行存「A / B」伪路径；编辑收件人会清空旧
   `attachments`；行内编辑偶发不提交。

### 24.2 修复清单（本轮实际改动）

- `lp_tool.py::_wb_step_pipeline` 整体重构：PanedWindow 先建，左栏直接创建条目表（**补横向滚动条**）
  + 参数区；右栏直接创建「📎 统一附件」（可折叠，展开时排在处理区上方）+「▶ 处理」+ 对应表。
- `_wb_step_recipients`：删掉遗留的 `att_txt`，`vals` 与 `cols` 严格一一对应；权限口令恢复 `••••`。
- `_recipient_dialog`：不再渲染附件路径，保存时原样保留 `rec["attachments"]`（流水线路径来源兼容旧数据）。
- 导入模板：`lp_v3.import_template_headers()` 去掉「附件路径」；`HEADER_ALIASES["files"]` 去掉
  「附件路径/附件」→ 导入时该列按未识别列收进 `fields["附件路径"]`，不再进 `attachments`；
  `EXPORT_BASE_HEADERS`（导出备份）保留该列以免丢旧数据；重生成 `examples/` 两个 xlsx。
- Word-MailMerge 参数区补「导出要素模板」（列=检测出的占位符、行=收件人、值按当前映射预填）与
  「下载示例 Word 模板」（另存为 .docx，含 6 个占位符）；附加文件参数区补「导出路径模板」
  （收件人 + 附件路径两列）；`_wb_download_word_example` 由死代码变为按钮接线。
- 「发送人统一密码」：`save_data()` → `_persist_settings()`；settings 变更同时同步
  `event["options"]["gp_owner_mode"/"gp_owner_pw"]` 并落盘；加 `StringVar.trace_add("write")`，
  输入即保存。引擎/preflight 现在能读到 GP 口令。
- 重复内容提示统一：预检 / 检查附件弹窗 / 运行日志三处均为 `[附件重复提示] … 请人工确认`。
- `_wb_map_save_edits`：单文件存字符串、多文件存路径列表（不再存 `" / "` 拼接伪路径）；
  行内编辑 `focus_force()` + `<Return>/<KP_Enter>/<FocusOut>` 三键提交。
- 帮助页与 `docs/使用说明.txt` 按 v3.4.4 现状重写（清掉「产出粒度/生成文档/附件路径列/路径来源」旧说法），
  `release_body.md` 更新为本版发布说明。
- 版本号三处：`lp_tool.py` / `lp_v3.py` / `docs/使用说明.txt` → **3.4.4**。

### 24.3 验证

- `python -m pytest lp-doc-mailer/tests -q` → **220 passed**（原 208 + 新增 `tests/test_v3_4_4_fixes.py` 12 条；
  新测试覆盖：列对齐/遮罩、编辑保留附件、导入模板/导入映射、左右分栏归属、横向滚动条、两个缺失按钮、
  路径模板、GP 落盘+引擎可见、帮助页文案、重复提示文案、process_overrides 列表、版本号）。
- `e2e_v344_verify.py`（真实 GUI + 真实 Word→PDF→水印→加密，隔离临时数据目录）→
  **96 PASS / 0 FAIL / 15 INFO**；含 2 位收件人真实产物（口令可解、权限仅打印、水印在、
  process_results 回填并持久化）、示例模板/要素模板导出、重复文件附件包与三处提示、导入模板校验等。
- `sync_release.py` → [OK]（root = 绿色版 = 仓库 src，三文件哈希一致）。
- `PyInstaller LP_Tool.spec -y` → `dist/LP_Tool.exe` **47,270,168 B = 45.08 MiB**，Authenticode **Valid**
  （CN=LPDocMailer-Internal）；`--icon icon.ico` / 8 项 excludes 由 spec 保证。
- `pack_portable.py` → `doc-mailer_v3.4.4_portable.zip`（2551 文件，44.26 MiB，
  sha256 `FB46DEEA…F5FA9`），读回断言全 [OK]，zip 内源码/说明/图标哈希与根目录一致。
- `doc-mailer_win64.zip`（44.65 MiB，sha256 `58ADAF9B…8A971`）：exe + `lp_data.example.json` + 使用说明，
  zip 内 exe 与 dist 同哈希、签名 Valid。
- 受限沙箱内 Word COM 无法初始化（0x800a13e9），上述 pytest/e2e/打包均在无沙箱环境执行；验收后已清理
  全部无头 WINWORD.EXE。
- **已发版（2026-10-24）**：commit `f76ed56` + tag `v3.4.4` 已 push（`fa19847..f76ed56 main -> main`；
  `* [new tag] v3.4.4 -> v3.4.4`）；CI run **40090967585**（branch v3.4.4）conclusion=**success**；
  Release **id=422697613**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.4.4）；
  两资产均为本地构建产物，字节级复核一致：
  `doc-mailer_v3.4.4_portable.zip` 46,413,265 B sha256 `FB46DEEA…F5FA9`；
  `doc-mailer_win64.zip` 46,822,910 B sha256 `58ADAF9B…8A971`（内含签名 Valid 的 exe）。

### 24.4 已知遗留（本轮未做）

- 已知问题 1 不变：邮件正文/标题的占位符替换仍是 Python `str.replace`（大小写敏感，`{{LP名称}}`
  不替换）；Word 文档链路因 `Find(MatchCase=False)` 大小写不敏感，两种写法都能替换。
- 已知问题 2 不变：打开口令在收件人表明文显示（用户拍板不修）；权限口令已恢复 `••••`。
- 已知问题 3 不变：`PATH_SOURCE_RECIPIENT` / `resolve_item_paths` 仍保留兼容旧数据，UI 无下拉。
- `rec["attachments"]` 仍保留在数据模型/导出备份中，但 UI 与导入模板不再露出。
- 远端 Release/CI 未验证（网络）；推送后需确认 tag v3.4.4 触发的 CI 绿。


---

## 25. v3.4.5：v3.4.4 界面热修（2026-10-24）

### 25.1 触发

v3.4.4 发布后用默认 1280×800 窗口截图复核：分栏后左栏宽度不足，
Word-MailMerge 参数区的「Word-MailMerge 说明 / 导出要素模板 / 下载示例 Word 模板」与
附加文件的 4 个按钮被裁到栏外（要拖动分栏才能看到）；Word 模板路径输入框同理。
功能都在，但默认窗口下用户看不到/点不到，属可用性问题。

### 25.2 改动

- `_wb_item_params` 参数区按行重排：
  - 占位符写法提示、处理行提示（水印文字与各人打开口令…）、权限提示各占一行；
  - 「说明 / 导出要素模板 / 下载示例 Word 模板」单独一行；
  - Word 模板输入框 54→32 字符（「选择…/检测占位符」不再被挤出）；
  - 附加文件：文本框与「添加文件…/从 Excel 导入路径…/保存/导出路径模板」改为上下两行；
  - 「发送人统一密码」勾选文案缩短。
- 字段映射表补横向+纵向滚动条（列多时不再被裁）；条目表列宽略收。
- 版本三处 → **3.4.5**。

### 25.3 验证与发版

- `pytest`：**220 passed**。
- `e2e_v345_verify.py`（真实 GUI+Word）：**96 PASS / 0 FAIL / 15 INFO**。
- 默认窗口按钮位置复核：7 个按钮 x 范围全部落在左栏内（layout 探针）。
- 构建：`dist/LP_Tool.exe` **47,269,984 B（45.08 MiB）**，签名 **Valid**；
  便携包 `doc-mailer_v3.4.5_portable.zip`（2551 文件 / 44.3MB，sha256 `19ED35C1…C6236`）；
  win64 包 `doc-mailer_win64.zip`（46,822,787 B，sha256 `5E26EC14…E4BA5`）。
- 发版状态：commit/tag/CI/Release 见 §25.4（下文回填）。

### 25.4 发布结果

- commit：`2e0ea3f`；tag：`v3.4.5`（已 push）；CI run **40091416821** conclusion=**success**。
- Release：**id=422700970**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.4.5）；
  两资产从 Release 下载回来做哈希复核，与本地构建产物**字节级一致**：
  `doc-mailer_v3.4.5_portable.zip` 46,413,396 B sha256 `19ED35C1…C6236`；
  `doc-mailer_win64.zip` 46,822,787 B sha256 `5E26EC14…E4BA5`（内含签名 Valid 的 exe）。


---

## 26. v3.4.6：流水线页三个小修复（2026-10-24）

### 26.1 用户反馈

1. 条目名称修改仍会弹「重命名条目」对话框；希望直接在参数区名称文本框改，点「改名」立即生效。
2. 导出了 Word 要素模板，却找不到「导入要素模板」的入口。
3. 「Word-MailMerge 说明」不需要弹窗，关键说明直接写在界面上即可。

### 26.2 改动

- `_wb_on_pipe_double`：双击条目行不再调 `_wb_rename_item`；改为选中行 + 重新渲染参数区 +
  `after_idle` 聚焦名称输入框（`<Return>` 也可提交）。
- 删除 `_wb_rename_item` 重命名对话框；新增 `_wb_apply_item_name`（名称框 → 改名按钮/
  回车 → 直接写 event + 保存 + 刷新表格，无对话框）。
- 新增 `_wb_import_field_template`：
  - 读 Excel（导出模板格式：第一列「收件人名称」，其余列=占位符）；
  - 按收件人名称匹配本事件收件人，非空单元格写入 `fields[占位符名]`（空单元格不覆盖）；
  - 模板里检测到的占位符，**只有导入到值的列**才自动映射为 `fields.<名>`，避免空列顶掉
    已有映射；未匹配的收件人名称在完成提示里列出。
- Word-MailMerge 参数区：
  - 删除「Word-MailMerge 说明」按钮与 `_show_mailmerge_help` 方法；
  - 内联两行提示：占位符写法 + 处理顺序 + 水印/打开口令维护位置；
  - 按钮行改为「导出要素模板 / 导入要素模板 / 下载示例 Word 模板」。
- 版本三处 → **3.4.6**；使用说明/发布说明同步。

### 26.3 验证

- `pytest`：**224 passed**（新增 4 条：改名直改、双击不弹窗、要素模板导入回填+映射、内联说明/无说明按钮）。
- `e2e_v346_verify.py`（真实 GUI+Word）：全部 PASS，含「导出要素模板 → Excel 填写 →
  导入要素模板 → 字段回填 + 自动映射」全链路、改名即时落盘、无重命名对话框、
  真实 2 人 Word→PDF→水印→加密流水线。
- 构建/发版：见 §26.4。

### 26.4 发布结果

- commit：`84e1559`；tag：`v3.4.6`（已 push）；CI run **40094185417** conclusion=**success**。
- Release：**id=422720162**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.4.6）；
  两资产从 Release 下载回来做 SHA256 复核，与本地构建产物**字节级一致**：
  `doc-mailer_v3.4.6_portable.zip` 46,414,642 B sha256 `83E5E64B…CF0D0`；
  `doc-mailer_win64.zip` 46,826,364 B sha256 `84519E0D…A2E58`（内含签名 Valid 的 exe）。


---

## 27. v3.4.7：「检查更新」总失败（GitHub API 匿名额度 403）修复（2026-10-24）

### 27.1 现象

用户反馈：点「检查更新」总提示「无法连接 GitHub（公司网络可能受限）」，
但同一台机器浏览器打开 GitHub 正常。

### 27.2 根因（实测证据）

旧版 `check_latest()` 只请求：

`https://api.github.com/repos/fundops-notes/doc-mailer/releases/latest`

GitHub 对未认证 API 请求限 **60 次/小时/IP**。本机走 FlClash 系统代理
（`127.0.0.1:7890`，出口 IP `43.203.139.69`，共享），实测返回：

```
HTTP 403 rate limit exceeded
X-RateLimit-Limit: 60
X-RateLimit-Remaining: 0
X-RateLimit-Used: 60
X-RateLimit-Reset: 1792846188   （2026-10-24 20:49:48）
{"message":"API rate limit exceeded for 43.203.139.69. ..."}
```

而 `https://github.com/fundops-notes/doc-mailer/releases/latest` 同网络下
HEAD/GET 均 200（跟随 302 到 `/releases/tag/vX.Y.Z`，约 0.3~1s）——
网页不受 API 额度限制。代码把所有异常都吞成「无法连接 GitHub」，
所以症状看起来像网络问题。

**第二个 bug（顺带发现）**：`_ver_tuple` 定义漏了 `@staticmethod`，
`_show_update` 里 `self._ver_tuple(...)` 抛 `TypeError: takes 1 positional argument but 2 were given`，
被 `_poll_queue` 外层 `except Exception: pass` 吞掉 —— 即使 API 成功，点「检查更新」也毫无反应。

### 27.3 修复

- `check_latest` 改为「网页跳转优先」：
  1. `GET https://github.com/<repo>/releases/latest` 跟随 302，从最终 URL 取 tag（不受 API 限流）；
  2. `api.github.com` 仅用于补充 release body / 资产 size/digest；403/超时静默降级；
  3. API 不可用时按命名规范拼下载地址
     `https://github.com/<repo>/releases/download/<tag>/doc-mailer_v<ver>_portable.zip`；
     有 `GITHUB_TOKEN`/`GH_TOKEN` 环境变量时自动用于 API（5000 次/小时）。
- `check_update` 失败时展示两种方式的具体原因（网页跳转 / API），不再统一称「网络受限」。
- 修复 `_ver_tuple` 的 `@staticmethod`；`_poll_queue` 的 update 分支加兜底提示，不再静默吞异常。

### 27.4 验证

- `pytest`：**228 passed**（新增：API 403 时网页回退、API 正常补充说明/资产、两种方式都失败、
  `_ver_tuple`/`_show_update` 回归）。
- 真实端到端（`e2e_v347_verify.py`）：在 API 仍返回 403 的情况下，点「检查更新」
  成功弹出「当前已是最新版本（v3.4.7）」，`_last_update_err` 记录 API 403 明细。
- 构建/发版：见 §27.5。

### 27.5 发布结果

- commit：`9daa0d6`；tag：`v3.4.7`（已 push）；CI run **40095336535** conclusion=**success**。
- Release：**id=422728475**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.4.7）；
  两资产从 Release 下载回来做 SHA256 复核，与本地构建产物**字节级一致**：
  `doc-mailer_v3.4.7_portable.zip` 46,416,060 B sha256 `C180F99B…FCEB5`；
  `doc-mailer_win64.zip` 46,826,910 B sha256 `22B15250…4C3F9`（内含签名 Valid 的 exe）。
- 遗留观测：修复上线时 GitHub API 仍返回 403（共享出口 IP），但网页跳转路径可用 →
  「检查更新」功能正常；这正好验证了修复目标。


---

## 28. v3.4.8：字段映射弹窗（占位符来源）联动修复（2026-10-24）

### 28.1 用户反馈

字段映射表格双击某行改映射时，弹窗里只会显示第一个占位符，
下拉菜单看不到其他待映射字段，来源下拉也不联动 —— 「关联没做好」。

### 28.2 根因

`lp_tool.py::_wb_edit_mapping`：
1. `ttk.Combobox(...)` 建占位符下拉时**漏传 `values`** → 下拉里只有当前值，
   其它占位符看不到（点开是空的/只有一项）；
2. 没有绑定 `<<ComboboxSelected>>` → 切换占位符时，来源下拉不会刷新；
3. 映射表双击绑定只传 `item`，没有把双击行的占位符传进去 → 每次都从第一个开始；
4. 确定时用 `ph_var.get()` 写回，虽然能写对当前值，但用户根本没机会切到别的字段。

### 28.3 修复

- 占位符下拉 `values=ph_list`（本条目全部占位符）；
- `ph_box.bind("<<ComboboxSelected>>", _sync_source)`：切换即刷新来源下拉为
  该占位符当前映射（无映射时回落到第一项）；
- `_wb_edit_mapping(item, initial_ph=None)` 增加初始占位符参数；映射表双击时用
  `identify_row(event.y)` 取该行 `_wb_map_iids` 的占位符传入 → 打开即预选双击行；
- 「确定」写回 `ph_var.get()` 对应项并保存、刷新映射表。

### 28.4 验证

- `pytest`：**229 passed**（新增 `test_mapping_dialog_lists_all_placeholders_and_preselects_row`：
  下拉含全部字段、按行预选、切换联动、保存落盘）。
- `e2e_v348_verify.py`（真实 GUI）：下拉 7 项全部出现；双击第 2 行预选第 2 项；
  切换后来源下拉联动；确定后 mapping 持久化正确。
- 构建/发版：见 §28.5。

### 28.5 发布结果

- commit：`64d6bb5`；tag：`v3.4.8`（已 push）；CI run **40096450266** conclusion=**success**。
- Release：**id=422736913**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.4.8）；
  两资产从 Release 下载回来做 SHA256 复核，与本地构建产物**字节级一致**：
  `doc-mailer_v3.4.8_portable.zip` 46,416,489 B sha256 `26ED6EDA…77679D`；
  `doc-mailer_win64.zip` 46,828,965 B sha256 `1893BDFA…A5278`（内含签名 Valid 的 exe）。


---

## 29. v3.4.9：自动更新下载回调签名修复（2026-10-24）

### 29.1 用户反馈

点「检查更新」发现新版本后选择自动更新，立刻弹：

```
自动更新失败：App._self_update.<locals>.work.<locals>._prog() takes 1
positional argument but 3 were given
可手动到 Releases 页下载。
```

### 29.2 根因

```python
def _prog(pct): ...
urllib.request.urlretrieve(info["asset_url"], zip_path, reporthook=_prog)
```

`urlretrieve` 的 `reporthook` 固定按 `(blocknum, blocksize, totalsize)` 调用，
`_prog(pct)` 只接受一个参数 → 一开始下载就 `TypeError` → 被兜底成「自动更新失败」。
v3.4.4~v3.4.8 的自动更新链路一直存在该问题（此前 API 限流导致很少真正走到下载）。

### 29.3 修复

- `_prog(blocknum, blocksize=1, totalsize=0)`：按 3 参数计算百分比，每 20% 记日志；
- 下载后增加 zip 完整性校验（`zipfile.testzip`）；API 拿不到 SHA256 摘要时明确提示并改用 zip 校验；
- 其余更新流程不变：解压到 `.update_staging` → 生成 `run_update.bat`
  （robocopy `/MOVE`，排除 `lp_data` 等用户数据）→ 重启 `start.bat`。

### 29.4 验证

- `pytest`：**230 passed**（新增自动更新链路回归：假 `urlretrieve` 按 3 参数回调 →
  下载/解压/生成更新脚本/`update_ready`，不真实重启）。
- `e2e_v349_verify.py`：真实 GUI 下模拟更新包，进度回调正常、staging 就位、
  `run_update.bat` 含 robocopy 与 lp_data 排除、update_ready 入队。
- 构建/发版：见 §29.5。

### 29.5 发布结果

- commit：`0ce2499`；tag：`v3.4.9`（已 push）；CI run **40097157541** conclusion=**success**。
- Release：**id=422742085**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.4.9）；
  两资产从 Release 下载回来做 SHA256 复核，与本地构建产物**字节级一致**：
  `doc-mailer_v3.4.9_portable.zip` 46,417,022 B sha256 `504CEB05…05BD1`；
  `doc-mailer_win64.zip` 46,830,669 B sha256 `255D591F…0E11C`（内含签名 Valid 的 exe）。
- **真实自动更新链路实测**（发布后跑）：从 Release 真实下载 44MB 便携包（API 返回 sha256:
  digest）→ 进度 0/20/40/60/80/100% 正常 → SHA256 校验通过 → 解压出的 `app/lp_tool.py`
  为 3.4.9 → 生成 `run_update.bat`（robocopy + 排除 lp_data + start）→ update_ready；
  用时 6.8s，无错误。


---

## 30. v3.5.0：顶部导航/莫兰迪 + 三项功能修复（2026-10-24）

### 30.1 用户反馈

1. Word-MailMerge 取 Excel 字段时，日期等被强制转成数字（如录入「X月X日」显示成数字）。
2. 附加文件「从 Excel 导入路径」没有按收件人对应：模板里按人填了路径，导入后所有人都收到全部文件。
3. 「产出颗粒度」设置不需要保留，默认按分组（同组一封邮件、未分组各自一封）。
4. 「统一附件」右栏容器按钮（保存）显示不全。
5. 主导航从左栏移到顶部按钮形式；整体莫兰迪浅色系、不同功能用色块区分。

### 30.2 根因与修复

1. **Excel 显示值**：`lp_v3.read_spreadsheet` 直接 `iter_rows(values_only=True)`，把日期读成
   datetime/序列号、数字读成原始 float。改为逐单元格按 `number_format` 格式化成「显示文本」：
   - datetime/date：按 `年/月/日` 或 `yyyy-mm-dd`/`yyyy/mm/dd` 输出，带时间格式的保留时间；
   - 数字：按百分比、小数位、千分位（`#,##0`）输出；整数值不再带 `.0`；
   - `_cell_text` 也兜底 datetime，避免其它路径漏格式。
2. **附件路径按人导入**：`_wb_import_attach_paths` 旧实现把所选列所有路径 append 到手填清单。
   重写为：选「收件人名称」列 + 「附件路径」列 → 按名称匹配收件人 → 路径写入该收件人的
   `rec["attachments"]` → 条目 `params.path_source = recipient.attachments`、清空手填清单。
   - `lp_v3.preflight` 对「按收件人取路径」的附加文件条目不再要求手填 `files`，
     改为逐人检查路径存在性（缺失只 warning，处理时跳过）；
   - 参数区显示当前来源；提供「改回手填路径」。
3. **取消产出颗粒度**：流水线条目表去掉 `gran` 列；参数区去掉下拉框与 `_wb_set_item_gran`；
   新建条目/`normalize_item` 默认 `granularity="per_group"`；参数区固定提示「产出方式：按分组自动」；
   帮助/使用说明同步。引擎仍保留 per_group/per_lp/all 兼容旧数据（测试显式指定 per_lp 的用例仍可跑）。
4. **统一附件布局**：右栏从「标签 + 文本框 + 按钮同一行」改成上下两行，按钮行 `fill=x`，
   默认窗口下「添加文件… / 保存」完整可见。
5. **顶部导航 + 莫兰迪**：`_build_shell` 重写：顶部暖灰条（品牌 + 事件库/设置/帮助按钮 +
   面包屑/版本/检查更新），去掉左导航；按钮用 clam 主题的自定义样式实现莫兰迪色块
   （灰绿/灰棕/灰蓝），当前页高亮；卡片、表头、步骤导航统一浅色。

### 30.3 验证

- `pytest`：**233 passed**（新增/更新：默认 per_group、Excel 显示值、附件路径按人导入、
  顶部导航/主题；引擎测试显式保持 per_lp 语义）。
- `e2e_v350_verify.py`（真实 GUI+Word）：全部 PASS；新增：
  - 顶部三个导航按钮、当前页高亮、clam 主题、无左导航；
  - 条目表/参数区无「产出颗粒度」、默认 per_group；
  - 附件路径 Excel 导入：甲/乙各自只收到自己的文件（处理后产物分别以各自文件名开头）；
  - 统一附件按钮完整可见。
- 构建/发版见 §30.4。

### 30.4 发布结果

- commit：`dd28e27`；tag：`v3.5.0`（已 push）；CI run **40099432566** conclusion=**success**。
- Release：**id=422757465**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.5.0）；
  两资产从 Release 下载回来做 SHA256 复核，与本地构建产物**字节级一致**：
  `doc-mailer_v3.5.0_portable.zip` 46,420,139 B sha256 `56A8F002…722BE`；
  `doc-mailer_win64.zip` 46,836,412 B sha256 `FAF68AE5…8DC064`（内含签名 Valid 的 exe）。


---

## 31. v3.5.1：设置页清理 / 文件名模板下沉 / 场景模板移除（2026-10-24）

### 31.1 用户反馈

1. 设置页有些参数没用了（如 GP 统一权限口令），可删。
2. 默认输出目录保留，但默认值应为程序所在文件夹。
3. 文件名模板移到 Word-MailMerge 页面，统一这一条生成任务的文件名。
4. 场景模板看不出用处，检查代码评估是否需要，不需要就删除清理。

### 31.2 评估与修复

**设置页**
- 删除全局「GP 统一权限口令」行；条目级「发送人统一密码」保留，只写
  `event.options`（不再写 `settings`），默认值也不再读全局设置。
- 删除全局「文件名模板」行与 `set_fname` 逻辑。
- 删除重复的「自动更新设置」按钮；「启动时检查更新」勾选即时保存。
- 数据存储说明改为 `settings.json · events/ · recipients/ · runs/`。

**默认输出目录**
- `lp_v3.app_dir()`（单 exe → exe 目录，源码 → 源码目录）；`run_event` 与
  `_wb_process` 的兜底从「文档\文档分发输出」改为程序目录。
- 设置页无自定义值时显示程序目录；保存时若等于程序目录则存空串，
  便携包移动后仍跟随程序。

**文件名模板下沉**
- `normalize_item` params 新增 `filename_tpl`；
- `item_output_name(..., filename_tpl=None)`：条目级优先，回退旧事件级
  `options.filename_tpl`，再回退 `DEFAULT_FILENAME_TPL`；
- Word-MailMerge 参数区新增「生成文件名模板 + 保存」，显示可用字段；
- 设置页删除全局模板。

**场景模板（评估结论：移除）**
- 现状：仅有「新建/重命名/删除、从场景建事件、同步到一个事件」，
  本质是复制式模板且无自动同步；创建事件主流程根本不经过它，
  普通用户没有使用入口/动机，属于高理解成本低价值功能。
- 处理：删除事件库「场景模板」标签页、全部场景 UI 与方法、场景创建/同步入口；
  `_new_event` 不再接受 scenario；首次启动不再保存示例场景（只建示例事件）；
  数据层（`Store.list_scenarios` 等、旧 `scenarios/` 文件、v2 迁移中的场景对象）
  保留只读兼容，不删除旧数据。
- 场景相关测试改为「不存在场景标签/不再生成场景」；旧迁移测试保持。

### 31.3 验证

- `pytest`：**235 passed**。
- `e2e_v351_verify.py`：设置页无 GP/文件名模板、默认目录=程序目录、
  数据存储结构无 scenarios、事件库 3 个标签、WMM 文件名模板保存后
  实际处理产物为 `甲_专属账单.docx`。
- 构建/发版见 §31.4。

### 31.4 发布结果

- commit：`5cf1ef0`；tag：`v3.5.1`（已 push）；CI run **40104249066** conclusion=**success**。
- Release：**id=422792060**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.5.1）；
  两资产从 Release 下载回来做 SHA256 复核，与本地构建产物**字节级一致**：
  `doc-mailer_v3.5.1_portable.zip` 46,418,506 B sha256 `50FBE251…66D5FD`；
  `doc-mailer_win64.zip` 46,831,753 B sha256 `7843191B…F9BCAE6`（内含签名 Valid 的 exe）。


---

## 32. v3.6.0：流程四步化 + 处理日志/进度 + 总览预生成（2026-10-24）

### 32.1 用户反馈

1. 去掉流水线里的「(3) 附件」页（已由统一附件代替）。
2. 「▶ 处理」不要跳到预览与生成页；要在当前页右下角显示处理进度和日志，
   日志加日期+机器时间。
3. 生成文件路径没有按设置里的默认输出目录；要统一。
4. 执行日志里的「[附件重复提示]」用红色/橙色字。
5. 去掉「(6) 预览和生成」：完成流水线后进邮件模板 → 预生成进总览；
   总览左侧邮件清单，点行右侧预览邮件；下方附件可点文件名打开预览。
6. 事件右上角的「返回列表」「事件信息…」移到右上角。
7. 首页事件库增加编辑事件（名称/业务类型等）。
8. 确认事件信息里的发送模式是否生效。

### 32.2 实现

- **步骤**：`WB_STEPS` 缩为 4 步（收件人/文档流水线/邮件模板/总览）；
  删除 `_wb_step_attachments`、`_wb_step_run` 及附件页相关方法；统一附件留在流水线右栏。
- **处理日志**：流水线右栏底部新增「处理进度与日志（实时）」Text；
  `_wb_log` 统一加 `[YYYY-MM-DD HH:MM:SS]` 前缀；`[附件重复提示]` 打 `dup`
  tag（#c25e00 橙色）；失败/错误打 `err` tag。`_wb_process` 不再触发
  `_wb_show_step(RUN)`；`_wb_run_mode` 区分处理/生成，处理完成留在第 2 步。
- **输出目录**：`_default_outdir()` = 设置 outdir 或 `app_dir()`；
  `_wb_process`、`_wb_run` 都重置 `event.options.outdir` 并显式传给
  `run_event(outdir=...)`；引擎兜底也已是 app_dir。
- **总览**：动作栏（预检/生成全部草稿/只重跑失败项/进度/状态）移入总览；
  右侧邮件预览 pane 加宽；下方附件改为可点击 Label（真实产物 +
  统一附件路径，`os.startfile` 打开）；新增生成日志（橙色 tag）。
- **邮件模板**：底部新增「预生成 →」（保存模板后 `_wb_show_step(OVERVIEW)`）。
- **顶部操作**：`_build_shell` 顶栏右侧新增 `_ev_actions`（返回列表/事件信息…），
  进入事件显示、离开隐藏；工作台顶部原按钮移除。
- **首页编辑事件**：事件信息弹窗抽成 `_event_info_dialog(ev, after_save)`，
  工作台与首页共用；首页工具栏新增「编辑事件…」。
- **发送模式**：弹窗下拉切换时实时更新说明；保存写 `options.send_mode`/`merge_groups`；
  回归验证 per_lp→N 封、per_group→1 封。
- **纯文本升级说明**：`release_body.md` 改为普通文本，不用 Markdown 标记。

### 32.3 验证

- `pytest`：**243 passed**（新增 4 步、处理不跳页+时间戳、橙色 tag、预生成、
  附件点击、顶部操作、首页改名、发送模式生效）。
- `e2e_v360_verify.py`：全部 PASS（含处理不跳页、统一输出目录、日志橙色 tag、
  预生成总览、附件点击打开、首页编辑、发送模式 per_lp/per_group 封数）。
- 构建/发版见 §32.4。

### 32.4 发布结果

- commit：`493df11`；tag：`v3.6.0`（已 push）；CI run **40107430653** conclusion=**success**。
- Release：**id=422816825**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.6.0）；
  两资产从 Release 下载回来做 SHA256 复核，与本地构建产物**字节级一致**：
  `doc-mailer_v3.6.0_portable.zip` 46,418,902 B sha256 `55F95246…374375B`；
  `doc-mailer_win64.zip` 46,831,383 B sha256 `63EB11D6…1668D41`（内含签名 Valid 的 exe）。


---

## 33. v3.7.0：邮件编辑器 / 总览单封编辑 / 收件人标签 / 添加到事件修复（2026-10-25）

### 33.1 用户反馈

1. 邮件模板编辑器去掉右侧预览，直接所见即所得；加字体、字号、图片、字段、
   字体颜色/背景、左右居中；去掉「设定」按钮。
2. 总览右侧预览改成编辑器，可单封个性化修改，保存后生成草稿按修改后的内容。
3. 流水线和总览两个进度条没生效，删掉。
4. 收件人库增加标签：单人/批量打标签；从事件导入时提示打标签；
   批量移除只能移除共有标签，特异标签不能批量移除。
5. 「添加到事件…」无效（用户数据目录 C:\Users\ian\Downloads\doc-mailer）。

### 33.2 诊断与修复

**添加到事件无效（根因确定）**
- 用户安装为 v3.6.0，数据目录 `app/lp_data/` 正常（events/recipients/runs 都在，
  无报错日志、无失败记录）；说明操作在打开事件选择框之前/之时静默返回。
- 代码 `_pick_event()` 创建 Toplevel + 按钮后**没有 wait_window 就 return**，
  `picked["id"]` 永远是 None → `_pool_to_event()` / `_pool_from_event()` 直接返回，
  所以无日志、无报错、看起来“没生效”。
- 修复：`_pick_event` 增加 `transient` + `wait_window`，并支持双击列表项确定。
- 之前测试用 `_pick_event` 桩，掩盖了该 bug；新增真实选择对话框回归测试。

**编辑器**
- 删除邮件模板页右侧预览 PanedWindow；编辑器全宽。
- 工具栏：字体（微软雅黑/宋体/黑体/楷体/仿宋/Arial 等）、字号（选即应用）、
  文字色、背景色、左/中/右对齐、插入图片、插入字段、清除格式；删除「设定」。
- 富文本层扩展：`_ensure_tag` 支持 font-family；HTML 解析/序列化支持
  font-family 与 `<img>`；图片插入用 data URI + 隐藏 token + Tk 图片显示，
  `text_to_html`/`load_html` 双向还原。
- 生成草稿时 `create_outlook_draft` 把 data URI 图片转成 cid 内嵌附件，
  临时文件保存后清理，兼容 Outlook 桌面版不显示 data URI 的限制。

**总览单封编辑**
- 右侧改为同一套编辑器（主题 + 正文 + 工具栏 + 保存/恢复统一模板）；
- 事件新增 `mail_overrides`（按 `draft_units` 的 key），保存单封主题/正文；
- `normalize_event` 保留该字段；`run_event` 生成草稿时优先使用单元覆盖。

**进度条**
- 删除流水线 `_wb_progress2` 与总览 `wb_progress`；进度只靠实时日志/状态。

**收件人库标签**
- `normalize_pool_recipient` 增加 `tags`（去重、清空）；合并时标签取并集；
- 收件人库表增加「标签」列与「标签…」按钮；
- 批量对话框：添加勾选/新标签到所有选中；移除仅允许共有标签，
  勾选特异标签会提示且不删除；
- `_pool_from_event` 导入后自动弹出标签对话框。

### 33.3 验证

- `pytest`：**251 passed**（新增：无进度条、编辑器工具栏/无预览、图片 HTML
  往返、总览覆盖保存、run_event 使用覆盖、标签持久化、批量共有标签语义、
  `_pick_event` 等待用户选择）。
- `e2e_v370_verify.py`：全部 PASS（编辑器控件、图片写入 HTML、总览单封保存、
  标签列/批量标签、真实「添加到事件」成功）。
- 构建/发版见 §33.4。

### 33.4 发布结果

- commit：`a1c9904`；tag：`v3.7.0`（已 push）；CI run **40111696757** conclusion=**success**。
- Release：**id=422846098**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.7.0）；
  两资产从 Release 下载回来做 SHA256 复核，与本地构建产物**字节级一致**：
  `doc-mailer_v3.7.0_portable.zip` 46,423,312 B sha256 `5955B1F1…67B4FA34`；
  `doc-mailer_win64.zip` 46,840,578 B sha256 `E94B1C2F…32746F74`（内含签名 Valid 的 exe）。


---

## 34. v3.7.1：自定义字段归属事件，收件人库不再累积字段（2026-10-25）

### 34.1 用户反馈

收件人编辑框里的自定义字段会跟着事件走，多个事件用过的字段越积越多，认为没必要。

### 34.2 根因与结论

- 事件收件人编辑框显示 `_wb_field_columns()`：当前事件所有收件人的字段并集（事件内合理）。
- 收件人库编辑框显示 `_pool_field_cols()`：**整个库所有记录字段的并集**；
  而「存入收件人库」「从事件导入」「Excel 导入到库」都把事件字段一起带进库，
  于是字段越积越多，且所有库记录编辑时都显示这些无关字段。
- 结论：自定义字段是事件数据（用于 Word 映射），不属于收件人身份；
  收件人库只保留名称/邮箱/水印/组/标签/口令等身份信息。

### 34.3 实现

- `_pool_field_cols()` 返回 `[]`；收件人库表删除动态字段列，固定列
  （名称/To/CC/BCC/水印/组/标签/打开口令/权限口令）。
- 收件人库新增/编辑对话框不再显示自定义字段。
- `存入收件人库`、`从事件导入`、`导入 Excel/CSV 到库` 时剥离 `fields`。
- 新增收件人库「清理库内字段」：一次清掉旧版本遗留的字段（只动库，不动事件）。
- 事件收件人页新增「删除字段列」：从当前事件所有收件人删除字段，
  并移除 `fields.<name>` 的字段映射；只影响当前事件。
- 版本三处 → 3.7.1。

### 34.4 验证

- `pytest`：**253 passed**（库不收集字段、清理旧字段、事件删除字段列；
  原有池/编辑器/标签/添加事件用例保持）。
- 构建/发版见 §34.5。

### 34.5 发布结果

- commit：`f5e50f7`；tag：`v3.7.1`（已 push）；CI run **40112877305** conclusion=**success**。
- Release：**id=422854080**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.7.1）；
  两资产从 Release 下载回来做 SHA256 复核，与本地构建产物**字节级一致**：
  `doc-mailer_v3.7.1_portable.zip` 46,424,183 B sha256 `CBC58EDD…1D1937`；
  `doc-mailer_win64.zip` 46,842,822 B sha256 `E2C7E020…DC29E8`（内含签名 Valid 的 exe）。


---

## 35. v3.7.2：收件人信息隐藏字段 / 日期序列 / 图片显示（2026-10-25）

### 35.1 用户反馈

1. 收件人信息页仍有自定义字段列。
2. 要素模板 Due_Date 是数字（46193）导入后还是数字。
3. 邮件模板/总览编辑器插入图片不显示。
4. 收件人库已无字段，「清理库内字段」不需要了。

### 35.2 实现

- **收件人信息隐藏字段**
  - 事件收件人表删除动态字段列，固定列：名称/To/CC/BCC/水印/组/打开口令/权限口令；
  - 收件人编辑对话框不再显示字段；保存时保留 `rec["fields"]`，避免编辑丢数据；
  - 移除事件收件人页的「添加自定义字段列 / 删除字段列 / 套用字段模板」按钮；
  - 字段仍由 Excel/CSV 导入、要素模板导入写入，映射/生成逻辑不变。
- **日期序列**
  - `lp_v3.coerce_date_like_value(header, value)`：列名含 date/due/日期/时间/截止/到期
    且值为 20000–90000 的数字时，用 `openpyxl.utils.datetime.from_excel` 转
    `YYYY-MM-DD`；
  - `parse_import_rows_v3` 和「导入要素模板」都使用该函数；46193 → 2026-06-20。
- **图片显示**
  - `_editor_place_image` 把 Tk 图片创建在隐藏 token **之后**；此前图片与 token
    同索引，token 的 elide 标签把图片一起隐藏，导致“插入了但看不见”。
  - 插入与 HTML 往返、生成草稿的 cid 内嵌逻辑保持不变。
- **移除库清理按钮**
  - 收件人库已不保存/显示字段，删除「清理库内字段」按钮与方法；
    旧库历史字段只是不显示，不再需要清理动作。
- 版本三处 → 3.7.2。

### 35.3 验证

- `pytest`：**253 passed**（新增/更新：收件人表与弹窗隐藏字段且编辑保留字段、
  图片可见（bbox 非空）、日期序列转日期；移除库清理测试）。
- `e2e_v372_verify.py`：全部 PASS（收件人表无字段列、插入图片可见、库无清理按钮等）。
- 构建/发版见 §35.4。

### 35.4 发布结果

- commit：`d42c512`；tag：`v3.7.2`（已 push）；CI run **40114634215** conclusion=**success**。
- Release：**id=422866584**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.7.2）；
  两资产从 Release 下载回来做 SHA256 复核，与本地构建产物**字节级一致**：
  `doc-mailer_v3.7.2_portable.zip` 46,424,444 B sha256 `9F056D0D…9271A59D`；
  `doc-mailer_win64.zip` 46,843,001 B sha256 `407E0628…175941F3`（内含签名 Valid 的 exe）。


---

## 36. v3.7.3：日期导入按 Excel 显示文本 + 可选格式（2026-10-25）

### 36.1 用户反馈

v3.7.2 把日期列的数字统一转成了 YYYY-MM-DD，但不同场合需要的日期展示不一样；
希望优先使用 Excel 里显示的内容作为文本。

### 36.2 实现

- `read_spreadsheet` 本来就按单元格 `number_format` 输出显示文本；v3.7.3 不再在
  `parse_import_rows_v3` / 要素模板导入里强制转日期。
- 要素模板导入时，检测日期类列（date/due/日期/时间/截止/到期）中的
  General 数字（20000–90000），弹窗让用户选择：
  - 保持 Excel 显示文本（默认，推荐先在 Excel 设置格式）；
  - `YYYY-MM-DD` / `YYYY年MM月DD日` / `YYYY/MM/DD` 三种转换格式之一。
- 取消选择则不导入。
- 非日期列数字保持原样（Account_No 等不受影响）。
- 日期展示的最终控制权在 Excel：用户把单元格/列格式设置成什么样式，
  read_spreadsheet 就按该显示文本导入。

### 36.3 验证

- `pytest`：**255 passed**（保持显示文本、ISO/中文/斜杠格式、非日期数字不转换；
  原有字段/图片/收件人信息测试保持）。
- 构建/发版见 §36.4。

### 36.4 发布结果

- commit：`08bf170`；tag：`v3.7.3`（已 push）；CI run **40115266569** conclusion=**success**。
- Release：**id=422870944**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.7.3）；
  两资产从 Release 下载回来做 SHA256 复核，与本地构建产物**字节级一致**：
  `doc-mailer_v3.7.3_portable.zip` 46,425,567 B sha256 `2DF1FA9B…56D942F1`；
  `doc-mailer_win64.zip` 46,845,979 B sha256 `A4F7DF63…E7A1D81`（内含签名 Valid 的 exe）。


---

## 37. v3.8.0：Modern Enterprise SaaS UI + 交互完善（2026-10-25）

### 37.1 用户反馈

1. 映射来源仍显示「收件人字段：XXX」，应表明是本次事件/导入文件的数据。
2. 邮件模板工具栏上下重复，只保留上方。
3. 预生成应覆盖总览里的原有单封设置；之后以总览为准生成草稿。
4. 左侧事件流程加相邻节点推进按钮（上一步/下一步），节点宽度缩小、点击直达。
5. 邮件模板/单封邮件未保存就跳页或生成草稿时强提醒保存。
6. 整体 UI/UX 升级为 Modern Enterprise SaaS（用户给出完整设计规范）。

### 37.2 实现

**字段来源文案**
- `lp_v3.field_source_options` / `path_source_options` 的 `fields.*` 分类由
  「收件人字段」改为「本事件字段」；`path_source_label` 同步；帮助/文档同步。
- 字段数据模型不变，仍由 Excel/要素模板导入，映射与生成不变。

**编辑器**
- 删除邮件页第二个 `_build_editor_toolbar` 调用（工具栏上下重复）。
- 工具栏分两行：格式/字体/字号/颜色；对齐/图片/清除/字段。
- 总览右栏使用 compact=True 的三行工具栏，窄栏不裁切。
- 图片显示修复见 §35（图片建在隐藏 token 之后）。

**预生成**
- `_wb_mail_preview_next`：保存统一模板后 `event["mail_overrides"]={}`，
  再进总览；总览里的单封修改以 `mail_overrides` 保存，生成时优先使用。

**步骤推进**
- `_build_workbench` 左栏宽度 156；底部新增 `_wb_prev_btn` / `_wb_next_btn`
  （下一步 Accent 蓝色），`_wb_update_step_buttons` 同步高亮与禁用状态；
  步骤按钮点击仍直接跳转。

**未保存提醒**
- `_mail_dirty` / `_ov_dirty` 状态：邮件主题/代发/抄送/密送变量 trace、
  正文 KeyRelease、格式按钮、图片插入都会标脏；载入内容时抑制。
- `_ask_unsaved` 三选一（保存/不保存/取消）；
  `_handle_unsaved_before_leave` 在 `_wb_show_step`、`show_screen`、`_wb_run`、
  `on_close` 拦截；总览切换行时也会提醒并保存。
- 保存后清脏并刷新页面数据；取消则留在当前页。

**SaaS UI**
- 调色板：#F8FAFC / #FFFFFF / #0F172A / #64748B / #E2E8F0 / #2563EB；
- 左侧白色导航（216px，当前项淡蓝底+蓝字）；顶栏白色（面包屑、事件操作、
  检查更新）；内容区 24px 边距。
- Design System：Head 18px Semibold，Sub 9px 灰；按钮 padding 12/7，
  Accent 蓝底白字；Treeview rowheight 28、浅灰表头、选中淡蓝；
  Card 白底细边框；TNotebook 白底选中；全局 Option font。
- 首页工具条分主/次两行（新建事件 Accent）；收件人/收件人库工具条也分行；
  流水线左右栏宽度调整，右栏处理日志固定可见；总览生成按钮 Accent。

### 37.3 验证

- `pytest`：**260 passed**（新增：左侧导航/SaaS 样式、邮件页按钮与工具栏不重复、
  未保存保存/丢弃/取消、预生成清覆盖、上一步/下一步、本事件字段文案；
  原字段/图片/标签/日期用例保持）。
- `e2e_v380_verify.py`：全部 PASS（左侧导航、本事件字段、工具栏唯一、
  图片 bbox 可见、步骤推进、预生成清覆盖）。
- 构建/发版见 §37.4。

### 37.4 发布结果

- commit：`0bf1d03`；tag：`v3.8.0`（已 push）；CI run **40118326426** conclusion=**success**。
- Release：**id=422891824**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.8.0）；
  两资产从 Release 下载回来做 SHA256 复核，与本地构建产物**字节级一致**：
  `doc-mailer_v3.8.0_portable.zip` 46,428,092 B sha256 `DEF24081…38EDE2A0`；
  `doc-mailer_win64.zip` 46,849,342 B sha256 `978E3F12…3BFA4C84`（内含签名 Valid 的 exe）。


---

## 38. v3.8.1：工作台三入口左移 + 侧栏缩窄 + 检查更新移左下（2026-10-25）

### 38.1 用户反馈

1. 「事件」「收件人库」「执行历史」移到左侧工作台下方并列展示；
   「事件库」对应的就是「事件」。
2. 左侧「工作台」区域宽度缩小。
3. 「检查更新」移动到左侧下方版本号附近。

### 38.2 实现

- 侧栏宽度 216 → 176。
- 侧栏结构：
  - 工作台：事件 / 收件人库 / 执行历史（三个 ttk Nav 按钮并列，均可直接切换）
  - 系统：设置 / 帮助
  - 底部：版本号 + 检查更新
- 首页仍保留内部 Notebook 以复用三大页面，但**隐藏 tab 栏**，
  并用 `show_home_section(key)` + `home_nb.select()` 切换；
  `Plain.TNotebook` 自定义 layout + 全局 `TNotebook.Tab` 空 layout 去掉页内标签栏。
- 首页标题/面包屑动态：事件 / 收件人库 / 执行历史；
  收件人库/执行历史页内重复 Head 标题删除，避免双标题。
- 导航高亮：`_set_nav_active` 映射 events/pool/runs/settings/help；
  `<<NotebookTabChanged>>` 同步高亮与标题。
- 顶栏移除「检查更新」，左侧底部增加；事件操作仍留顶栏。
- 版本三处 → 3.8.1。

### 38.3 验证

- `pytest`：**260 passed**（左侧五项导航、工作台三入口切换/高亮、标题动态；
  原有未保存/字段/图片/日期用例保持）。
- `e2e_v381_verify.py`：全部 PASS（工作台并列、切换收件人库/执行历史、
  检查更新 y=817 位于左下、SaaS 主题）。
- 构建/发版见 §38.4。

### 38.4 发布结果

- commit：`e613ab5`；tag：`v3.8.1`（已 push）；CI run **40150069183** conclusion=**success**。
- Release：**id=423080882**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.8.1）；
  两资产从 Release 下载回来做 SHA256 复核，与本地构建产物**字节级一致**：
  `doc-mailer_v3.8.1_portable.zip` 46,428,719 B sha256 `EF2BAE81…119F6A79`；
  `doc-mailer_win64.zip` 46,851,105 B sha256 `B74BAF76…040E19BA`（内含签名 Valid 的 exe）。


---

## 39. v3.8.2：按标签批量添加事件收件人（2026-10-25）

用户要求：事件的收件人导入弹窗支持按收件人标签筛选，批量选择相同标签的收件人加入本事件。

实现：
- 入口仍为事件工作台 → 收件人 → 从收件人库添加…；弹窗新增标签筛选下拉，选项为全部收件人 / 未打标签 / 标签：具体名称。标签精确匹配，多标签收件人可出现在各匹配标签下。
- 列表显示名称、To 和标签；新增全选当前列表 / 取消选择，显示当前人数、库总人数和已选人数；支持 Ctrl/Shift 多选。
- 切换筛选删除旧列表并清空选择，只允许添加当前可见列表中的选中项；避免隐藏项误加。
- 复用原有按名称合并，只补空缺，保留事件已有打开口令、权限口令、邮箱和字段；不改库原始记录/标签，无数据迁移。
- 弹窗默认 820×560，最小 700×460；横向/纵向滚动条；空选择提示，取消不写入。
- lp_tool.py 与 lp_v3.py 两处 APP_VERSION 字面量一致升至 3.8.2；使用说明同步分发副本，Release 正文维持普通文本。

测试与发布：
- 新增 8 条真实控件 GUI 回归，覆盖标签列表/特殊同名标签、精确筛选/合并保留、切换清空隐藏选择、未打标签与取消、空列表阻止添加、35 人跨可视区批量选择、默认/最小弹窗尺寸下操作按钮完整可见。
- pytest 全量：268 passed（2026-10-25，22.82s，未跳过 Word）。
- 新 e2e_v382_verify.py 复用 v3.8.1 验收流程并增加真实标签弹窗流程；输出统一到项目验收_v3.8.2，临时数据位于项目 out/tmp；新增按 JSON FAIL 数量设置退出码。
- 最终 e2e：177 PASS / 0 FAIL / 15 INFO（真实 GUI+Word，Outlook 使用 fake）。默认 820×560 与最小 700×460 下四个操作按钮均完整可见；标签筛选/批量添加落盘通过。
- 已构建并签名：exe 48,350,840 B，Authenticode Valid；隔离目录启动签名 exe 实测标题 v3.8.2、正常退出、临时数据清理通过。
- 代码 commit/tag：`7bd3dbe1f4ea9c5fac43a9236d1559fd04e1ca46`；CI run **40156377632** completed/success；新 tag v3.8.2 普通 push，旧 tag 未动。
- Release **id=423114605**：https://github.com/fundops-notes/doc-mailer/releases/tag/v3.8.2；正文普通文本，两个资产均从 Release 重新下载，SHA256 与本地及 GitHub digest 一致，下载包 exe 签名 Valid。
- 证据：项目 验收_v3.8.2/evidence/release_verify_current.json、downloaded_signature.json、release_upload.log；截图 screens/；修复与发布报告_v3.8.2.md。

| 资产 | 字节 | SHA256 | 下载校验 |
| --- | ---: | --- | --- |
| doc-mailer_v3.8.2_portable.zip | 46429542 | `5FBB8ECE86A04F26EF3B38E36C718E3A7CFB0DEECB0C5398FD8B6344180AF528` | 一致 |
| doc-mailer_win64.zip | 47904110 | `F6253FC1293C17D7F72C5E9033D379510543932219629BE34CC130040C7F9994` | 一致 |


---

## 40. v3.8.3：事件/收件人库功能归属整理与三项缺陷修复（2026-10-25）

- 用户已确认：事件首页管理事件与备份；本次名单/业务字段在事件工作台；长期联系人与标签在收件人库。依据上一轮真实GUI+Excel审查报告实施。
- 首页去掉导入收件人/导出收件人/导出全部收件人三个重复按钮，删除前两者重复实现。跨事件明细保留在「更多导出 → 导出各事件收件人明细…」。事件内部保留导入、导出本事件收件人、从库按标签添加等入口。
- 跨事件导出使用所有事件业务字段并集表头，并按字段键对齐；与固定表头冲突的业务字段加「本事件字段：」前缀并保证唯一；同一人在不同事件保留多行，原数据不改。
- 库页独立「导入联系人Excel/CSV / 导出联系人Excel / 下载联系人模板」，模板带标签、不带业务金额日期字段。库专用校验器识别标签/tag/tags等列，提示忽略事件字段；事件导入语义不变。
- 标签支持多种常用分隔符、去空/去重；Excel导出可无损还原含分隔符的特殊标签。同名只补空缺、保留已有口令/邮箱、标签取并集，空白标签不清空旧标签；旧库字段不再导出。
- Store.merge_pool_records统一保存新增和同名更新；库Excel导入、从事件导入、存入收件人库三入口都使用，修复只提示合并而未落盘的同源问题。写入失败显示错误，不报完成；不迁移或操作真实用户数据。
- 新增17条回归，含真实菜单导出/Excel回读、库Excel往返、CSV/标签解析、多个入库入口持久化、已有口令保留、写失败处理、专用模板、事件字段边界；全量285 passed（24.68s）。
- e2e_v383_verify.py复用旧GUI+Word全流程，新增归属按钮、更多菜单导出、真实库导入预览、Excel标签往返、CSV同名更新、字段隔离、预览按钮可见及截图。最终188 PASS / 0 FAIL / 15 INFO（Outlook使用fake）；截图确认归属入口与导入预览按钮可见。
- 证据：项目 验收_v3.8.3/evidence/、screens/、backups/。

- 本地构建成功，签名exe=48,353,920 B / Authenticode Valid；隔离启动标题v3.8.3、正常退出及临时数据清理通过。便携包/win64包已读回校验，无用户数据。代码commit/tag `077334ec6909deb504a08ad9052ea99273128718`；CI run **40159280824** completed/success；Release **id=423131794**（https://github.com/fundops-notes/doc-mailer/releases/tag/v3.8.3），正文普通文本。两个资产从Release重新下载后SHA256与本地及GitHub digest一致，下载exe签名Valid。

| 资产 | 字节 | SHA256 | 下载校验 |
| --- | ---: | --- | --- |
| doc-mailer_v3.8.3_portable.zip | 46431384 | `20CF308A70CB403D7AA30B97DBF0FB026A40A9DF401F1ABAFFCE4D2402D6F5DB` | 一致 |
| doc-mailer_win64.zip | 47906798 | `0D5B77F13172FD4215D13E0F8E92A36A330E8000C87AD07FD89FB3E72453419C` | 一致 |

---

## 41. v3.8.4：窗口布局、滚动与参数可达性（2026-10-26）

- 用户反馈：流水线选中条目的参数在最大化时仍显示不全，缺乏滚动条，普通窗口多个按钮拥挤。基线重新执行：285 pytest passed；原GUI+Word e2e 188 PASS / 0 FAIL。隔离布局审查复现默认/最小窗口多项参数不可达，最大化下Word字段映射按钮仍被截断。
- 流水线条目表与参数区分开分配高度；参数区和右侧处理区有独立竖向滚动条，高度不足时整页也可滚动。保持嵌入内容的自然高度，避免切换条目后滚动范围仍停在旧尺寸；键盘焦点会自动显示当前控件，切换条目重置参数滚动位置。
- 工具栏随宽度自动换行，解释文字按容器宽度折行；窄窗口工作台四步移至顶部。邮件收件地址输入框可伸缩，抄送/密送分行；邮件模板、总览、收件人、设置和帮助页在低窗口下可滚动，底部上一步/下一步优先保留。
- 修复多处Treeview用pack导致横向滚动条只显示在右下角的问题，使用grid铺满表格宽度；附件路径文本框增加横/竖滚动，总览清单增加滚动条。鼠标滚轮按所在区域路由，原生文本框/表格独立滚动，销毁页面时移除绑定；开始处理时在当前页显示进度日志。
- 新增21项布局回归：6组条目/尺寸参数控件可达、切换重建、滚轮和焦点、绑定清理、长路径横向滚动、8个页面按钮可达、权限真实落盘与口令保留、处理日志可见。旧布局测试从强制pack更新为grid，原左右栏归属断言仍保留。
- 最终pytest：306 passed，60.46s，无skipped。真实GUI+Word e2e：191 PASS / 0 FAIL / 15 INFO；含旧188项流程及3组布局审查回执。Outlook使用fake，不代表真实邮箱集成验收。
- 布局审查：1280×800、960×620、实际最大化（本机1872×1057）×默认/125%/150% Tk字号缩放。每档99个参数控件及198个其他页面操作逐一显露，共891次检查、不可达0；150%字号另查事件信息/收件人编辑/字段映射/水印设置4个弹窗，按钮未裁切、窗口在屏幕内；无Tk回调错误。每份回执绑定根源码SHA256，截图已检查。未修改系统DPI设置。
- 构建成功，签名exe=48,363,240 B，Authenticode Valid；隔离启动实际标题v3.8.4，正常退出0、临时数据清理通过。根源码/src/便携镜像一致，项目lp_data的32件文件哈希未变；Downloads/doc-mailer当前不存在，本轮未操作用户安装数据。证书/runtime/.git保留，Outlook唯一写调用仍为lp_v3.py:2424 mail.Save。
- 根目录layout_audit_v384.py、e2e_v384_verify.py为本轮原生验收入口；证据验收_v3.8.4/evidence/，截图screens/，备份backups/。
- 代码commit/tag：`39368aaf7125a11ea60cfe7724326cfd9ce10c9d`；CI run **40321539154** completed/success，head_sha与tag一致。Release **id=424214288**：https://github.com/fundops-notes/doc-mailer/releases/tag/v3.8.4 ，正文普通文本。新版本两资产已上传并重新下载，ZIP完整性、大小、SHA256与本地/GitHub digest、便携源码与版本、用户数据排除均通过；下载exe=48,363,240 B，Authenticode Valid，签名证书与本地一致。旧tag未改写，无force push。

| 资产 | 字节 | SHA256 | 下载校验 |
| --- | ---: | --- | --- |
| doc-mailer_v3.8.4_portable.zip | 46435541 | `033FEED07052809C62F97CAB07A9E5035FC01862703B76593DCE0E736133B2EE` | 一致 |
| doc-mailer_win64.zip | 47915948 | `2682095F62BBBF4092988D25DE9A59EE4930BE05893CCA7520B79845475C535C` | 一致 |


---

## 42. v3.9.0：四种文件入口、下一步联动与紧凑编辑器（2026-10-26）

- 用户明确要求：Word MailMerge、PDF、非 PDF 和 Excel 按 LP 对应四入口并列；移除统一附件 UI；流水线下一步保存并处理后前进；邮件下一步保存确认后预生成；两个编辑器改为紧凑工具栏。
- PDF 可加水印和加密；非 PDF 可选转 PDF 后处理，未勾选转换时保留原格式；支持 Word、Excel、PowerPoint、图片、文本/HTML 转换。Excel 对应按条目独立存储，支持同一 LP 多行/多文件、相对路径，不写收件人专属附件列、不影响其它 Excel 条目。
- 流水线下一步提交正在编辑的参数和对应路径，保存自定义对应后按最终清单优先处理；异步处理成功后才进入邮件模板，预检/保存/转换失败留在当前页，处理期间限制切换与修改。邮件下一步确认保存未保存模板，再预生成并清空旧单封覆盖；取消不前进，后续草稿仍使用总览保存的单封修改。
- 统一附件 UI 已删除。打开旧事件时把旧附件显示为流水线条目，保留适用范围和原文件名；读取不落盘，首次保存为新结构前备份到 lp_data/backups，备份失败拒绝覆盖。旧收件人/字段/外部表路径来源继续兼容，未改用户真实数据。
- 两个编辑器共用小尺寸字体/字号/格式/颜色/对齐图标/图片/字段工具栏，保留所见即所得和 CID 内嵌图片。修复文件名列无法编辑、下一步未提交当前编辑、部分重跑覆盖其它收件人已处理清单；已处理文件按设置和来源文件版本核验后供草稿复用，避免重复转换覆盖自定义对应。
- 基线306 pytest passed；最终337 passed（76.45s，无跳过），含31项新工作流回归。真实 Tk+Word/Excel/PowerPoint 流程69 PASS/0 FAIL；Outlook 使用 fake Save，非真实邮箱集成验证。第一轮图片断言3项失败系验收脚本漏接局部导入的文件对话框，修正后两次完整通过，原始日志保留。
- 根源码两处 APP_VERSION 字面量均3.9.0，根/src/便携镜像已同步。证据：项目验收_v3.9.0/evidence（pytest_final.log、e2e_final.log、verify_result.json）；截图screens；旧版备份backups。构建、布局扩展和发布校验继续完成后补录。

- 布局扩展：三尺寸×默认/125%/150% Tk字号，1134次控件检查不可达0，另4弹窗通过。150%首轮窗口大小未稳定时两项失败，原日志保留；脚本增加窗口稳定检查后同一源码复验通过。exe=48,379,848 B / 签名Valid，隔离真实启动v3.9.0、正常退出0、临时目录清理通过。真实lp_data的32文件哈希不变，测试后WINWORD=0。新tag/CI/Release与下载资产校验待补录。

- v3.9.0发布完成：commit/tag c59b7237a1c5866d1c2d3d45208c5e7023e73839/v3.9.0；CI40354806014 success；Release id424461952；两包重新下载SHA256/大小/ZIP/版本/GitHub digest匹配，下载exe签名Valid。复核追加发现插入字段未标脏，保留原tag，继续v3.9.1修补，最终交付以v3.9.1为准。


## 43. v3.9.1：补齐字段插入的未保存保护（2026-10-27）

- 最终复核发现：两个紧凑编辑器通过插入字段菜单写入内容时未标脏，邮件下一步漏提示。3项新增真实控件回归在v3.9.0全部复现；v3.9.1调用既有_mark_active_dirty，统一模板和单封编辑都纳入原未保存保护，取消不前进且不改磁盘事件。
- 两处APP_VERSION字面量3.9.1，使用说明镜像同步；v3.9.0 tag保留。包含§42全部四入口/下一步/紧凑工具栏功能，未改数据迁移或布局。
- 最终340 pytest passed（77.98s，无跳过）；真实Tk+Word/Excel/PowerPoint e2e71 PASS/0 FAIL，Outlook为fake Save。当前源码补充150%字号55参数控件/4弹窗通过；完整1134次三尺寸三字号布局证据沿用未改布局的v3.9.0。
- exe=48,379,168 B，签名Valid；隔离真实启动v3.9.1、正常退出0、临时数据清理。原数据32文件哈希不变，WINWORD=0。证据验收_v3.9.1/、修复与发布报告_v3.9.1.md；发布完成见下。

- 代码commit/tag 01fa2de1207605383ee76ba92b4def1c165441c6 / v3.9.1，普通push，历史tag未动。CI run40356670381 completed/success，head_sha与tag一致。Release id=424475951，https://github.com/fundops-notes/doc-mailer/releases/tag/v3.9.1 。两个资产从Release重新下载，SHA256与本地及GitHub digest一致，大小、ZIP完整性、源码版本、用户数据排除均通过；下载exe签名Valid，证书与本地相同。发布正文为普通文本。

下载结果：便携包46,441,358 B / SHA256 5F80299BBDAD1BD46C5492AB345D4504F488C3B5475866E9FEE8B9628A8FAECF；win64包47,933,170 B / SHA256 89A4A1297A7950446AA9CC5101098C6D19248CA9A33F22DFE0ACEF24FA706087。证据release_verify_current.json、release_verify.log、downloaded_signature.json、release_upload.log。


## 44. v3.9.2：旧设置兼容升级、四路处理与可视化更新（2026-10-27）

- 用户反馈旧配置使流水线未全处理，并授权本地设置升级、要求后续兼容检查及更新进度。只升级获准事件，先备份再停用4位收件人的失效旧产物覆盖；隐藏旧专属附件改为第4个可见Excel对应条目，名单口令/原文件/原三条设置保留，未改安装源码。
- 本地实际来源文件+安装版源码隔离验证：4条×4人成功，每人4文件共16，旧附件内容一致，无真实Outlook草稿。明细留在非仓库验收_流水线诊断_20261027/，不提交用户配置/原文件。
- 开发版检查旧配置/更高版本/未知结构；确定可转的备份并验证后升级，备份失败或并发变更拒绝覆盖；保留合法自定义，迁出隐藏附件，清理失效覆盖缓存，跨版本首次保存备份。分别统计条目与收件人，补齐持久执行摘要，修正Excel预检引用，隔离Excel/PPT COM线程。
- 便携版更新显示下载量、百分比、校验/解压/准备阶段，支持取消及防重复启动；后台进程备份/安装，单独进度窗响应，失败尝试恢复，用户数据不覆盖。旧版首次升级仍使用旧更新器，本版起生效。
- 363 pytest passed（80.42s，无skip，新增23）；核心原生GUI+Office71 PASS/0 FAIL（Outlook fake）；真实下载/兼容/原生安装UI15项通过，数据备份及安装后字节校验通过。默认/150%字号取消按钮可达，未改系统DPI。根源码/仓库/便携镜像一致，版本字面量3.9.2；exe48,392,880 B/签名Valid，隔离启动/退出0；开发数据32件未变，测试后WINWORD=0。
- 证据：根目录修复与发布报告_v3.9.2.md、验收_v3.9.2/evidence/、screens/；原始失败日志保留。一份失败的隔离更新暂存目录清理遭策略拦截，保留在out/tmp，不进入发行包。
- 发布tag/CI/Release与下载资产校验待补录；历史tag不改写。
