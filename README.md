# Press CSV Merger

> 把压机每分钟导出的海量小 CSV，合并成「日期 + 机台 + 类型」大文件，并安全清理原始文件的 PowerShell 生产工具。
> A PowerShell production tool that merges thousands of tiny per-minute CSV exports from industrial presses into one file per [date + machine + type] — with a delete path designed to never lose a row.

**零依赖**：只用 Windows 自带的 PowerShell 5.1，不装任何东西、不联网、不写注册表。
**为现场而生**：无 Excel、无网络的车间电脑，U 盘拷入即用。

---

## 解决的痛点

某大型制造企业的压机每台每分钟导出 2 个 CSV 文件（Utility 工艺数据 + WORKINFO 批次信息）：

| 场景 | 文件数 |
|---|---|
| 单台每天 | 2,880 个 |
| 3 台每天 | 8,640 个 |
| 12 台满产一天 | **34,560 个** |
| 两周积压 | **≈ 48 万个** |

卡顿的元凶不是数据量（实际只有 33 MB），而是**文件数量**：

- 48 万个小文件 × 4KB 簇 ≈ **1.9 GB** 磁盘占用（实际内容仅 33 MB）
- Windows 单目录文件数上万后，枚举/复制速度**指数级劣化**，资源管理器卡死
- U 盘拷贝几十万个小文件，以小时计

**本工具把这 48 万个 2 行小文件合并成约 336 个按「日期+机台+类型」组织的大文件**，实测压缩比 **1,208 : 1**（文件数）/ **21 : 1**（磁盘占用），U 盘秒传。

## 核心特性

### 数据安全（11 项，全部在 `Code.ps1` 中实现）

| # | 机制 | 说明 |
|---|---|---|
| S1 | 合并阶段严格只读 | 只用 `File.ReadAllText` 读取源文件，绝不修改、绝不移动 |
| S2 | 安全时间窗 | 默认只处理文件名时间戳早于今天 00:00 的文件，当天正在写的一律不碰 |
| S3 | 行数校验才转正 | 先写 `.tmp`，回读校验「合并行数 == 参与合并的源文件数」才转正；失败保留 .tmp 并告警，绝不静默产出残缺文件 |
| S4 | 逐组原子操作 | 合并一组 → 校验 → 只删这一组 → 下一组；中断时已删的必已安全合并，重跑即续 |
| S5 | 默认不删除 | 删除必须显式进入 MergeAndClean 模式并人工输入 Y 确认 |
| S6 | 预览模式 | 只统计、不写不删，先看规模再动手 |
| S7 | 单点失败不中断 | 单个文件读取失败 → 记日志跳过、**绝不删除该文件**，整体继续 |
| S8 | 完整日志 | 每次运行写 `Logs\RUNLOG_*.txt` + `DELETED_*.txt`（删除清单），失败另出 `FAILURE_*.txt` 报告 |
| S9 | 不做的事 | 不碰非 CSV、不写注册表、不联网、不装任何东西、不递归删除 |
| S10 | 删除白名单复核 | 删除前逐个校验：文件名匹配 `^\d{14}_.+_(Utility|WORKINFO)\.CSV$` + 必须位于源目录树下 + 再过一次冷却期 |
| S11 | 幂等增量 | 重跑只追加新文件（读回已合并文件的 SrcFile 列去重），不重复、不覆盖 |

### 系统保护（6 项，优先保证不打扰压机生产）

| # | 机制 | 说明 |
|---|---|---|
| P1 | 进程降优 | 自动降到 BelowNormal，系统忙时让出 CPU |
| P2 | 磁盘节流 | 每 200 个文件主动休眠（默认 30 ms），绝不占满磁盘 |
| P3 | 单实例锁 | `_running.lock` 防止排程与手动运行撞车，第二个实例直接退出 |
| P4 | 全局时限 | 默认 30 分钟自动停在安全点，此时**什么都不删**，下次继续 |
| P5 | 紧急停止开关 | 在工具目录放一个 `STOP.txt`，下一个安全点即停，不删任何文件 |
| P6 | 排程静默运行 | 隐藏窗口、不弹问题，操作员无感 |

### 其他实用设计

- **SrcFile 溯源列**：合并后每一行都记录来自哪个原始文件，零丢失可追溯，也为下游合成器预留了反推时间戳的接口
- **智能编码识别**：配置文件自动识别 UTF-8 BOM / 系统 ANSI / 无 BOM UTF-8，简体繁体路径都不会乱码
- **跨目录同名自动加前缀**：不同子目录的同名文件不会互相覆盖
- **`cmd del` 批量删除**：比逐文件 `File.Delete` 快约 25 倍（实测 54 ms/个 → 约 2 ms/个）
- **内置环境自检**：`0_CheckEnv.bat` 只读检查 PowerShell 版本、编码页、目录读写权限、排程能力，输出报告给工程师

## 实测数据

全部来自 `docs/feasibility-report.md` 记录的真实验收（2026-10-14）：

| 指标 | 数值 |
|---|---|
| 全量合并 | 114,769 个文件 → 95 个文件 / 114,769 行，**零丢失、零失败** |
| 核心耗时 | 枚举 6.0 s + 解析分组 1.2 s + 读写合并 139.4 s ≈ **147 s** |
| 日常外推 | 一天 34,560 个文件 ≈ **45 s/天** |
| 文件数压缩比 | **1,208 : 1**（114,769 → 95） |
| 磁盘压缩比 | **21 : 1**（≈448 MB → ≈21 MB） |
| 合并+删除 | 隔离环境 8,860 个：合并 11.8 s，删后原件清零 |
| 排程端到端 | 960 个文件 → 5 个文件 + 原件清零，**2.4 s，返回码 0** |
| 安全边界 | 当天的 CSV、命名不符的 CSV、`.txt`/`.docx` —— 一个都没被删 |
| 幂等性 | 重跑只追加新行（8,860 → 8,863），零重复 |

## 安装与使用

**环境要求**：Windows 10 / 11，PowerShell 5.1（系统自带）。无需安装任何东西。

### 3 步上手

```text
1. 用记事本打开 0_Settings_EDIT_ME.txt，设置唯一一行：
       Src=D:\PressData\Export      ← 压机导出 CSV 的目录
2. 双击 1_Preview.bat              ← 预演：看看会处理多少文件（不写不删）
3. 双击 3_MergeClean.bat           ← 合并 + 删除源文件（会要求输入 Y 确认）
```

然后把 `_Merged` 文件夹拷走（U 盘/共享盘均可）。

### 每日自动运行（Windows 排程）

```text
1. 先跑一次 2_Merge.bat（确保 Src= 已配置）
2. 双击 5_InstallSchedule.bat
   - 默认每天 08:00 后台静默执行
   - 回车接受，或输入其他时间（如 07:30）
3. 取消：双击 6_RemoveSchedule.bat
```

> ⚠️ 排程要求电脑**开机且已登录 Windows**。错过时间后开机会自动补跑（StartWhenAvailable）。
> 若现场每天关机，请设置自动登录，或请 IT 将任务改为「不管用户是否登录都执行」。

### 文件说明

| 文件 | 作用 |
|---|---|
| `0_Settings_EDIT_ME.txt` | **唯一需要编辑的文件**：`Src=` / `Out=` / `SafeDays=` |
| `0_CheckEnv.bat` | 环境自检（只读） |
| `1_Preview.bat` | 预演：只统计，不写不删 |
| `2_Merge.bat` | 合并，保留源文件 |
| `3_MergeClean.bat` | 合并 + 删除源文件（输入 Y 确认） |
| `4_AutoRun.bat` | 仅供排程调用（静默、无交互） |
| `5_InstallSchedule.bat` / `6_RemoveSchedule.bat` | 注册 / 取消每日排程 |
| `Code.ps1` | 主程序（请勿编辑） |
| `Schedule.ps1` | 排程注册程序（请勿编辑） |
| `CheckEnv.ps1` | 环境自检程序（请勿编辑） |
| `README.txt` | 现场操作卡（一页纸） |

运行时会产生：`Logs\`（运行日志/删除清单/失败报告）、`_Merged\`（合并结果）、`_running.lock`（运行锁）、`_StartupError.txt`（仅启动失败时出现）。详见 `.gitignore`。

### 合并规则

```text
输入：{YYYYMMDDHHmmss}_{机台}_{Utility|WORKINFO}.CSV   （每文件 = 表头 1 行 + 数据 1 行）
输出：{YYYYMMDD}_{机台}_{Utility|WORKINFO}.csv         （表头 1 行 + 当天全部数据行 + SrcFile 列）

例：20260817_M01_Utility.csv    ← 该机台当天所有 Utility 行（约 1440 行）
   20260817_M01_WORKINFO.csv   ← 该机台当天所有 WORKINFO 行
```

Utility 与 WORKINFO **永不混在一个文件**。行内数据按时间戳升序。

### 高级参数（一般不用）

加在 `.bat` 里 `Code.ps1` 之后：

```bat
-Src "D:\path"       覆盖源目录
-Out "E:\path"       覆盖输出目录（可指向 U 盘）
-Mode Preview        预览模式
-SafeDays 0          连同当天的文件一起处理（仅限导出已停止时）
-MinAgeMinutes 30    冷却 30 分钟（默认 10）
-MaxFiles 20000      本次只处理 2 万个（手动分批）
-TimeoutMinutes 60   时限 60 分钟（默认 30）
-ThrottleMs 100      每 200 个文件歇 100 ms（默认 30）
-DeleteSource        合并后删除源文件
-Yes                 删除不再询问
-NoPause            结束不等待按键
```

## 目录结构

```text
press-csv-merger/
├── Code.ps1                  # 主程序（~1000 行，含全部安全机制）
├── CheckEnv.ps1              # 环境自检
├── Schedule.ps1              # 排程注册/取消
├── 0_CheckEnv.bat
├── 0_Settings_EDIT_ME.txt    # 唯一配置文件
├── 1_Preview.bat
├── 2_Merge.bat
├── 3_MergeClean.bat
├── 4_AutoRun.bat
├── 5_InstallSchedule.bat
├── 6_RemoveSchedule.bat
├── README.txt                # 现场操作卡
├── docs/
│   └── feasibility-report.md # 可行性分析 + 验收实测全记录
├── LICENSE                   # MIT
└── .gitignore
```

## 已知限制

1. **排程依赖开机登录**：电脑关机或注销时排程不运行（开机后自动补跑）。见上文说明。
2. **只认两种数据类型**：文件名必须匹配 `^\d{14}_.+_(Utility|WORKINFO)\.CSV$`（时间戳_机台_类型），其他命名会被跳过并记入日志。改 `Code.ps1` 里的正则即可扩展。
3. **CSV 解析从简**：数据为纯数字/字母、无引号包裹字段，用简单逗号切分（`Split-CsvLine`）。若源数据含引号包裹的逗号字段，需改用真正的 CSV 解析器。
4. **默认不碰当天文件**：`SafeDays=1` 是防竞态设计。导出已停止想清当天，用 `-SafeDays 0`。
5. **首次面对历史积压需 10~30 分钟**：30 分钟时限会自动停在安全点，分多次跑完即可，无需手动分批。
6. **合并结果喂旧合成器需适配**：旧合成器假设「一个文件 = 一个时间戳」，需要约 30~50 行的「合并模式」适配层才能消化合并后的大文件（用 `SrcFile` 列反推时间戳）。这是已记录的待办，见 `docs/feasibility-report.md` 第 8.5 节。
7. **Windows only**：依赖 `cmd del` 批量删除与 Windows 工作排程器。
8. **控制台界面**：本期只有 bat + 控制台，无 GUI。

---

## Quick Start (English)

```text
1. Open 0_Settings_EDIT_ME.txt in Notepad and set one line:
       Src=D:\PressData\Export
2. Double-click 1_Preview.bat      (statistics only, no write, no delete)
3. Double-click 3_MergeClean.bat   (merge + delete sources, asks for Y)
```

Requirements: Windows 10/11 with built-in PowerShell 5.1. Nothing to install, no network, no Excel.

Optional daily automation: double-click `5_InstallSchedule.bat` (default 08:00, hidden background run), remove with `6_RemoveSchedule.bat`. The PC must be powered on and logged in for the schedule to run; missed runs catch up on next login.

Safety highlights: read-only merge phase, safe time window (today's files untouched by default), `.tmp` + row-count verification before promotion, per-group atomic merge→verify→delete, delete whitelist (name pattern + inside source + cooling re-check), idempotent append-only re-runs, single-instance lock, 30-minute safe stop, `STOP.txt` emergency switch, full run logs.

Verified on a real 114,769-file sample: 95 merged files / 114,769 rows in ~147 s, zero loss, zero failures. Compression: 1,208:1 by file count, 21:1 by disk usage.

## License

MIT — see [LICENSE](LICENSE).
