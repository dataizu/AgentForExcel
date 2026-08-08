# Agent for Excel 验收记录

日期：2026-07-31

| 能力 | 验收证据 | 状态 |
|---|---|---|
| VSTO 加载 | 隔离 Excel 16.0 中 `COMAddIns("AgentForExcel").Connect = true` | 通过 |
| 工具注册 | 加载项内部 COM 诊断返回 `24/24`，无缺失工具 | 通过 |
| 任务执行闭环 | Stage1Smoke 已验证计划状态、提前结束续跑、写入回读守卫 | 通过 |
| 表格读取 | Release 加载项通过真实派发器读取 `A1:B3`，返回结构化预览 | 通过 |
| Power Query | Stage1Smoke 已验证创建、兼容加载、清洗和源数据变化后刷新 | 通过 |
| Power Pivot / DAX | `--powerpivot-only` 与完整 Stage1Smoke 均返回 2 表、1 关系、1 度量值和 1 个模型透视表 | 通过 |
| 受控 VBA 预览 | Release 加载项通过真实派发器生成白名单代码、一次性令牌和安全措施 | 通过 |
| 受控 VBA 安全边界 | 仍限定白名单、一次性令牌、执行前确认、备份和审计，不接受任意代码注入 | 通过 |
| 受控 VBA 完整执行 | `--vba-only` 在隔离 `.xlsm` 中验证预览、一次性令牌、执行前备份、真实执行、临时模块清理和隐藏审计表 | 通过 |
| VBA 工程访问 | 用户已开启“信任对 VBA 工程对象模型的访问”，Agent 环境自检全部通过 | 通过 |
| 功能中心 | 15 个能力按分析数据、生成报告、数据工程和自动化集中注册；目录冒烟测试通过 | 通过 |
| 分层设置 | 模型、安全、工作簿、隐私、外观和诊断页面可加载；新设置持久化与运行时提示词测试通过 | 通过 |
| 工具权限 | Power Query、Power Pivot 与受控 VBA 关闭后由执行层直接阻止调用 | 通过 |
| 实时选区 | Excel 选区事件实时更新工作簿、工作表、地址和行列规模；任务级锁定与 `@当前选区` 已接入对话上下文 | 通过 |
| 安全自动化 | `--permission-only` 已验证锁定选区内空白写入无需确认，覆盖已有内容会强制确认且拒绝后不改值 | 通过 |
| 权限设置持久化 | 自动化模式、三项低风险白名单和单次写入上限已通过 `--settings-only` 持久化与 WPF 加载测试 | 通过 |
| 原生切片器联动 | 完整 Stage1Smoke 验证 2 个切片器均连接 7 张透视表，并实际改变 KPI 与图表数据点 | 通过 |
| 下拉兼容联动 | 完整 Stage1Smoke 验证组合筛选、公式卡片、动态图表和重启恢复 | 通过 |
| 大范围性能路径 | `--performance-only` 验证 9,999 行 × 5 列的结构化预览与数据体检 | 通过 |
| 测试宿主位宽 | Stage1Smoke 已关闭 `Prefer32Bit`；本机 64 位 Excel 下 `32BITPREF=0`，避免跨位宽 Mashup 假表 | 通过 |
| Release 构建 | `AgentForExcel.csproj /t:Rebuild /p:Configuration=Release /p:AgentEdition=Automation`，0 警告、0 错误 | 通过 |
| 内部交付包 | `AgentForExcel-1.1.0.10-Automation.zip`，包含 `setup.exe`、VSTO 清单、应用文件、交付说明和 5 个验收/签名工具 | 通过 |
| 交付验收器 | 普通模式返回 15 PASS / 3 WARN / 0 FAIL；正式发货模式按预期因开发证书和未签名 `setup.exe` 返回 2 个 FAIL | 通过 |
| 安装后验收器 | 本机 x64 Excel 16.0 加载部署清单 1.1.0.10，验证注册、连接、24/24 工具、真实区域读取和环境自检，9 PASS / 0 WARN / 0 FAIL | 通过 |
| 正式签名流水线 | 支持隔离输出、证书信任与 Code Signing 用途门禁、双清单重签、EXE SHA-256/RFC 3161 签名以及签名后自动复核；开发证书演练已通过 | 通过 |
| 单机盲测编排 | 安装前、安装后和卸载后三阶段生成结构化机器报告；脚本不自动安装、卸载或关闭用户 Excel | 通过 |
| 双位宽证据门禁 | 严格汇总正式签名、包哈希、版本、时效、不同机器和 x86/x64 Excel；缺失项会 fail closed 并写入 `gaps` | 通过 |

## 本轮回归命令

```powershell
.\bin\Debug\Stage1Smoke\AgentForExcel.Stage1Smoke.exe
.\bin\Debug\Stage1Smoke\AgentForExcel.Stage1Smoke.exe --powerpivot-only
.\bin\Debug\Stage1Smoke\AgentForExcel.Stage1Smoke.exe --vba-only
.\bin\Debug\Stage1Smoke\AgentForExcel.Stage1Smoke.exe --settings-only
.\bin\Debug\Stage1Smoke\AgentForExcel.Stage1Smoke.exe --permission-only
.\bin\Debug\Stage1Smoke\AgentForExcel.Stage1Smoke.exe --performance-only
```

上述命令在 2026-07-31 均返回 `PASS`。最新内部 Automation 包：

- 路径：`artifacts\sellable-release\AgentForExcel-1.1.0.10-Automation\AgentForExcel-1.1.0.10-Automation.zip`
- SHA-256：`E5D70A5F85E242B433732FE914666125B0F23335C6D06F5B3D3C05A60D97E6A0`
- 构建验收报告：`artifacts\sellable-release\AgentForExcel-1.1.0.10-Automation\AcceptanceReport-build.json`
- 正式门禁报告：`artifacts\sellable-release\AgentForExcel-1.1.0.10-Automation\AcceptanceReport-formal-gate.json`
- 本机安装后验收报告：`artifacts\sellable-release\InstalledAcceptance-local.json`
- 本机 x64 机器报告：`artifacts\clean-machine-evidence\DATAIZU-x64-postinstall-1.1.0.10.json`
- 当前发布证据门禁：`artifacts\clean-machine-evidence\ReleaseEvidence-1.1.0.10.json`

## 尚未完成的正式发布门槛

- 当前开发机的 Smart App Control 为关闭状态，不能用它证明正式安装包能通过强制模式；应在另一台处于评估或强制模式的干净设备上验证全部安装、卸载和 Office 加载路径。
- 当前 VSTO 发布使用项目临时开发证书，仅适合本机测试；正式对外分发需替换为受信任的代码签名证书。
- 当前开发机可找到 Mage.exe，但没有 SignTool.exe；正式签名机还需安装包含 SignTool 的 Windows SDK。
- 尚未在两台干净设备上完成 32 位 Excel 与 64 位 Excel 的安装、首次启动、模型配置和卸载盲测。
- 当前 ZIP 是内部候选包，不得在完成正式签名和双位宽盲测前作为公开售卖版交付。
