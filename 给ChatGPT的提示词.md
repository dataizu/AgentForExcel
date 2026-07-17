# 复制下面所有内容发给 ChatGPT

---

我正在开发一个 Excel VSTO 加载项(C# / .NET Framework 4.8),目标是做一个 AI 数据分析插件,右侧挂一个 AI 对话面板(CustomTaskPane)。

但现在遇到一个棘手的问题:**VSTO 任务窗格(CustomTaskPane)在代码层面创建成功、Visible=True,但 Excel 界面上完全不显示。** 我已经排查了很多方向,需要你帮忙定位根因。

## 开发环境

- Visual Studio: **2026 (18.8.0)**
- Excel: **Microsoft 365,版本 16.0.20131.20126,x64**(更新通道疑似 Beta/Insider)
- .NET Framework: **4.8**
- VSTO: 运行时已正确加载
- 操作系统: Windows 11 (10.0.26200)

## 项目类型与配置

- 项目模板:VS 的 "Excel VSTO 外接程序"(Excel VSTO Add-in)
- csproj 的 ProjectTypeGuids: `{BAA0C2D2-18E2-41B9-852F-F413020CAA33};{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}`
- `<OfficeApplication>Excel</OfficeApplication>`
- `<LoadBehavior>3</LoadBehavior>`
- `<DefineConstants>VSTO40;UseOfficeInterop</DefineConstants>`(我加了 UseOfficeInterop 来解决嵌入互操作类型的枚举问题)
- 目标框架:.NET Framework 4.8

## 问题现象

1. 加载项**已被 Excel 加载**(COMAddIns 列表里 AgentForExcel 显示 Connect=True)
2. ThisAddIn 的 Startup 事件**正常触发**,代码全部执行(我写了文件日志,确认每一步都跑到了,无异常)
3. `CustomTaskPanes.Add(...)` 调用成功,返回的对象非 null
4. `_taskPane.Visible = true` 设置成功,读取回来确实是 True
5. **但 Excel 界面上完全不显示这个任务窗格** —— 右侧没有任何面板,浮动模式也看不到

## 关键代码(ThisAddIn.cs 的 Startup)

```csharp
private void ThisAddIn_Startup(object sender, EventArgs e)
{
    // 初始化应用上下文
    _appContext = new AppContext(Application, Globals.Factory);
    _appContext.Initialize();

    // 创建面板宿主控件(纯 WinForms UserControl,内含一个 Label)
    _host = new TaskPaneHost { Dock = DockStyle.Fill };
    var testLabel = new System.Windows.Forms.Label
    {
        Text = "测试面板",
        Dock = System.Windows.Forms.DockStyle.Fill,
        BackColor = System.Drawing.Color.LightYellow,
    };
    _host.Controls.Add(testLabel);

    // 添加任务窗格
    _taskPane = CustomTaskPanes.Add(_host, "Agent for Excel");
    _taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
    _taskPane.Width = 380;
    _taskPane.Visible = true;   // 设置成功,读回来是 True
    TaskPane = _taskPane;
}
```

## 我已经做过的诊断(请避免重复建议这些)

我做了大量排查,**以下都已经确认正常,不需要再建议**:

| 检查项 | 结果 |
|--------|------|
| 注册表 `HKCU\...\Excel\Addins\AgentForExcel` 的 LoadBehavior | = 3(正确,启动时加载) |
| Manifest 路径 | 正确指向 `AgentForExcel.vsto\|vstolocal` |
| Excel 的 Resiliency\DisabledItems(禁用项列表) | **空的**,加载项没被禁用 |
| Excel COMAddIns 列表 | AgentForExcel 在,Connect=True |
| VSTO 运行时 DLL | vstoee.dll、VSTOLoader.dll、所有 Microsoft.Office.Tools.*.ni.dll 都在 Excel 进程模块里 |
| Startup 代码是否执行 | **是的**,文件日志每次都完整产生,无异常 |
| CustomTaskPanes.Add 返回值 | 非 null,成功 |
| Visible 属性 | 设置 True 成功,读回来是 True |
| 编译 | 成功,生成了 AgentForExcel.dll 和 .vsto |
| WPF 导致的问题 | **已排除** —— 我改成纯 WinForms(只有一个 Label),还是不显示 |

## 我还尝试过的(都无效)

1. **停靠位置改成浮动**(`msoCTPDockPositionFloating`)—— 仍然看不到
2. **延迟强制显示** —— 用 WinForms Timer 在启动 1.5 秒后重新设置 `Visible=false` 再 `Visible=true`,日志显示成功,但界面仍不显示
3. **加 Ribbon 按钮**来手动切换 Visible —— Ribbon 选项卡本身也没出现在 Excel 功能区
4. **用 EnumChildWindows 枚举 Excel 主窗口的所有子窗口** —— 找不到任何 Task/Pane 相关的子窗口,**说明 Excel 根本没创建任务窗格的实际 UI 窗口**

## 核心矛盾

```
.NET 对象层面:CustomTaskPanes.Add 成功,对象存在,Visible=True
        ↓
Excel 原生 UI 层面:完全不创建对应的窗口,界面什么都不显示
```

## 我的疑问

请帮我分析:

1. 在代码成功、注册表正确、加载项已连接的情况下,为什么 Excel 不渲染任务窗格?最可能的根因是什么?
2. 这是否是 Excel 16.0.20131(疑似 Beta 通道)这个特定版本的已知 bug?你是否了解相关情况?
3. VSTO 创建 CustomTaskPane 有没有我可能忽略的**前置条件**(比如必须先有某个 Workbook 打开、必须在某个特定事件里创建、必须设置某个我没设的属性)?
4. 创建 CustomTaskPane 的最佳时机是什么?在 ThisAddIn_Startup 里创建是否有问题?应该改到 WorkbookOpen 或其他时机吗?
5. Ribbon 选项卡也没出现(RequestService 方法没被调用)—— 这和任务窗格不显示是否是同一个根因?
6. 如果确实是这个 Excel 版本的兼容性问题,有哪些**可靠的替代方案**可以实现"右侧挂一个 AI 对话面板"?(请考虑:独立 WinForms 窗口停靠在 Excel 旁、Windows API SetParent 嵌入 Excel 窗口、或其他思路)

请给出具体的、可操作的建议,最好附上代码示例。谢谢!
