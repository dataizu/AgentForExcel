# Agent for Excel 售卖版交付规范

## 结论

买家不需要安装 Visual Studio。售卖包应由 `setup.exe` 负责检查和安装 .NET Framework 4.8、VSTO Runtime 与 VSTO 加载项；买家只需要 Windows 桌面版 Excel、网络连接和安装权限。

当前 `Build-SellableRelease.ps1` 生成的是“在线前置组件”一键安装包：缺少的 Microsoft 运行库会由安装程序下载。它不是离线全量安装包，也不包含在线订阅、支付、账号或激活服务。

## 支持范围

- Windows 10/11。
- Microsoft 365、Office 2019、Office 2021 或更高版本的桌面版 Excel；32/64 位均须单独盲测。
- 不支持 Excel 网页版、macOS Excel 或受企业策略禁止 VSTO 的设备。
- 首次安装需要管理员权限来安装缺少的 Microsoft 运行库。
- 买家自行配置模型 API Key；不得把卖家的 API Key、Token 或第三方账号写进安装包。

## 版本分级

| 构建参数 | 面向买家的名称 | 可用能力 | 不包含 |
|---|---|---|---|
| `Trial` | 体验版 | 公式、图表、普通透视表和基础读写 | 数据体检、看板、Power Query、Power Pivot、VBA |
| `Standard` | 标准分析版 | 体验版 + 数据体检、趋势/对比分析、安全分析视图 | 看板、数据工程、VBA |
| `Professional` | 专业自动化版 | 标准分析版 + 看板、Power Query、Power Pivot / DAX | 受控 VBA |
| `Automation` | 自动化交付版 | 专业自动化版 + 受控 VBA 白名单配方 | 任意 VBA 注入、无人审查的自动化 |

能力分级同时在功能中心和工具执行层生效。它是用于售卖验证的产品边界，不是 DRM：在未接入签名许可证或在线授权前，不应把它宣传为“不可破解的订阅防复制”。

## 构建命令

在开发机 PowerShell 中运行：

```powershell
.\scripts\Build-SellableRelease.ps1 -Edition Trial -ApplicationVersion 1.1.0.0
.\scripts\Build-SellableRelease.ps1 -Edition Professional -ApplicationVersion 1.1.0.0
```

脚本会在 `artifacts\sellable-release` 创建一个全新的版本目录，拒绝覆盖旧交付物；若没有生成 `setup.exe`、`.vsto` 或 `Application Files`，会直接失败，不会把不完整安装包误当成售卖版。

构建完成后会自动运行只读验收器，检查包结构、清单版本、清单签名、敏感文件、Excel 位数和本机前置环境，并在版本目录生成 `AcceptanceReport-build.json`。验收器不会安装或卸载加载项，也不会修改注册表。

## 正式签名

正式签名前先把代码签名证书安装到当前用户或本机的“个人”证书存储，并安装包含 `SignTool.exe` 的 Windows SDK。证书必须在有效期内、具有私钥、包含 Code Signing 用途，并能构建到受信任根。时间戳地址应使用证书颁发机构提供的服务。

```powershell
.\scripts\Sign-SellableRelease.ps1 `
  -ReleaseDirectory .\artifacts\sellable-release\AgentForExcel-1.1.0.10-Automation `
  -OutputDirectory .\artifacts\signed\AgentForExcel-1.1.0.10-Automation `
  -CertificateThumbprint <正式证书指纹> `
  -Edition Automation `
  -Mode Formal `
  -TimestampUri <证书颁发机构提供的时间戳地址>
```

脚本始终复制到全新的输出目录，不覆盖原候选包；依次重签应用清单、更新并重签部署清单、用 SHA-256 和 RFC 3161 时间戳签署 `setup.exe`，再运行 Mage、SignTool 和正式交付验收器。自签名证书、未受信任证书、缺少 Code Signing 用途、缺少时间戳或缺少 SignTool 时，正式模式会在发货前失败。

`-Mode Rehearsal` 只用于开发证书演练，不会把 `externalDeliveryReady` 提升为 `true`。

## 干净机验收命令

内部候选包检查：

```powershell
.\scripts\Test-SellableRelease.ps1 `
  -Package .\artifacts\sellable-release\AgentForExcel-1.1.0.10-Automation\AgentForExcel-1.1.0.10-Automation.zip `
  -ExpectedVersion 1.1.0.10 `
  -ExpectedEdition Automation
```

正式发货门禁：

```powershell
.\scripts\Test-SellableRelease.ps1 `
  -Package .\AgentForExcel-1.1.0.10-Automation.zip `
  -ExpectedVersion 1.1.0.10 `
  -ExpectedEdition Automation `
  -RequireTrustedPublisher
```

结果分为三个层次：

- `packageStructureReady = true`：包结构、版本、文件引用和敏感文件检查通过。
- `machineEnvironmentReady = true`：当前电脑检测到桌面版 Excel 和 .NET Framework 4.8；缺少 VSTO Runtime 时由在线安装程序补齐。
- `externalDeliveryReady = true`：除上述条件外，ClickOnce 清单和 `setup.exe` 还必须使用当前机器信任的正式发布证书。

每台 32 位和 64 位 Excel 干净机都应保留一份 JSON 报告，报告中不得填写或附带模型 API Key。

发布 ZIP 内置 `AcceptanceTools`，解压后可以直接从该目录运行两套验收脚本，不依赖开发仓库。

安装完成后，先关闭所有 Excel 窗口，再运行真实加载验收：

```powershell
.\AcceptanceTools\Test-InstalledAddIn.ps1 `
  -ExpectedDeploymentVersion 1.1.0.10 `
  -ExpectedAssemblyVersion 1.1.0.0 `
  -ReportPath .\InstalledAcceptance.json
```

该脚本不会关闭用户正在使用的 Excel；发现已有 Excel 进程时会直接停止。它会启动自己的隔离 Excel，验证加载项注册、部署版本、VSTO 连接、24 项工具注册、真实区域读取和环境自检，随后只关闭自己创建的临时工作簿和 Excel 进程。`installedAddInReady = true` 才代表安装后的核心加载链路通过。

### 单台干净机完整证据

每台测试机都保留原始 ZIP，并从解压目录的 `AcceptanceTools` 运行以下命令。脚本不会自动安装、卸载或关闭用户的 Excel；安装和卸载仍由测试人员明确执行。

```powershell
$package = '..\AgentForExcel-1.1.0.10-Automation.zip'

.\AcceptanceTools\Invoke-CleanMachineAcceptance.ps1 `
  -Package $package -ExpectedVersion 1.1.0.10 -ExpectedEdition Automation `
  -Phase PreInstall -RequireTrustedPublisher `
  -ReportPath .\Evidence\preinstall.json

.\setup.exe

.\AcceptanceTools\Invoke-CleanMachineAcceptance.ps1 `
  -Package $package -ExpectedVersion 1.1.0.10 -ExpectedEdition Automation `
  -Phase PostInstall -ReportPath .\Evidence\postinstall.json

# Manually uninstall Agent for Excel, close Excel, then run:
.\AcceptanceTools\Invoke-CleanMachineAcceptance.ps1 `
  -Package $package -ExpectedVersion 1.1.0.10 -ExpectedEdition Automation `
  -Phase PostUninstall -ReportPath .\Evidence\postuninstall.json
```

### 双位宽最终门禁

收齐正式签名报告和两台不同电脑的报告后执行：

```powershell
.\AcceptanceTools\Test-ReleaseEvidence.ps1 `
  -FormalGateReport .\Evidence\FormalGate.json `
  -X86MachineReport .\Evidence\x86-postinstall.json `
  -X64MachineReport .\Evidence\x64-postinstall.json `
  -X86PostUninstallReport .\Evidence\x86-postuninstall.json `
  -X64PostUninstallReport .\Evidence\x64-postuninstall.json `
  -ExpectedVersion 1.1.0.10 `
  -ReportPath .\Evidence\ReleaseEvidence.json
```

只有正式签名、版本、包哈希、证据时效、两台不同电脑、x86 Excel、x64 Excel 和安装后验收全部通过时，`releaseReady` 才会为 `true`。缺失证据不会按零问题处理，而会明确列入 `gaps`。

## 发货门槛

1. 用受信任的正式代码签名证书重签 ClickOnce/VSTO 清单；开发证书只允许内部测试。
2. 在至少两台干净 Windows 设备上盲装：一台 32 位 Excel、一台 64 位 Excel；安装前运行 `Test-SellableRelease.ps1`，安装后运行 `Test-InstalledAddIn.ps1`，分别保存 JSON 报告。
3. 安装后启动 Excel，运行“Agent 环境自检”，确认版本、.NET、VSTO、Excel 位数和模型配置状态。
4. 记录安装成功率、首次自检通过率、单客支持时长和退款原因。首批 10 个创始用户的目标是：安装成功率 >= 80%，单客支持 <= 15 分钟。

## 订阅边界

首期采用“一次性授权 + 12 个月更新”验证需求。只有在已有稳定付费用户后，才接入签名许可证、激活服务和 AI 用量网关；支付、定价、续费和正式发布仍需要人工批准。
