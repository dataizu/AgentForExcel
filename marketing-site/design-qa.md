# Design QA — Agent for Excel 推广页

## Comparison target

- Source visual truth: `C:\Users\71824\.codex\generated_images\019f9d6b-fcda-7fc1-9c90-4d4db5d101e0\exec-6c71b1ea-97e6-42de-8c99-d91f7f796997.png`（已选方案 1）。
- Implementation screenshot: `D:\Solo_Work\AgentSpace\Zcode\VsTemplate\AgentForExcel\marketing-site\implementation-desktop-final.png`。
- Browser state: 首页顶部，所有弹窗关闭。
- Browser-rendered viewport: 1442 × 695 CSS px, devicePixelRatio 1.75；实现截图为 2524 × 1217 px。
- Source pixels: 1024 × 1536 px。为便于对照，英雄区被裁为 1024 × 493 px；实现截图等比缩放为相同的 1024 × 493 px。合成对照图：`qa-comparison-final.png`。
- Agent Window 的“Agent 正在控制”浮层属于浏览器自动化工具，不属于页面，已从页面本身的判断中排除。

## Findings

无 P0、P1、P2 问题。

- [P3] 英雄标题的自然换行与原图略有差异。
  - Location: `src/styles.css` 的 `.hero h1`。
  - Evidence: 原图第二行从“用 AI”开始；当前 1442 px 视口中“用”仍在首行。
  - Impact: 不影响主视觉层级、可读性或转化路径。
  - Follow-up: 如要针对固定 1440 px 截图追求像素级一致，可把标题最大字号再下调 1–2 px；当前保留更适合长标题的响应式换行。

## Required fidelity surfaces

- Fonts and typography: 使用系统中文字体栈；超大标题、辅助说明、导航和表格层级与原图的克制企业软件感一致，没有截断。
- Spacing and layout rhythm: 顶部导航、左右英雄区、真实 Excel 截图、信任说明条与后续白底信息区按原图的留白逻辑实现；桌面首屏无重叠或溢出。
- Colors and tokens: 以暖白、深墨绿和 Excel 绿为主；主按钮、边框、辅助文字与真实插件截图协调，未使用霓虹或夸张渐变。
- Image quality and asset fidelity: 使用项目内真实 `agent-logo.png` 和真实 Excel 侧栏截图；没有用 CSS/SVG/占位图替代产品图像。
- Copy and content: 文案明确 Windows 桌面版 Excel、API Key 自备、人工咨询交付和不支持范围；未宣传未接入的订阅、支付、账号或自动发货能力。
- Icons: 选定视觉不依赖额外功能图标；页面只使用真实产品 Logo，未加入自制 SVG 或 CSS 图标。
- Accessibility: 语义化导航、按钮、表格、对话框、表单标签与图片替代文本已具备；主交互有可见焦点样式；移动端在 600 px 以下改为单列布局与全宽操作按钮。

## Interaction checks

- 导航按钮可滚动至功能、场景和版本区域。
- “查看安装要求”可打开和关闭安装条件对话框。
- “咨询获取安装包”可打开咨询表单；已测试填写姓名、切换 Power Query / Power Pivot 场景、补充需求并生成咨询摘要。
- Vite 构建成功，`npm run test:sites` 4/4 通过。浏览器自动化工具未提供 Console 流；构建和浏览器交互中未出现运行时报错。

## Comparison history

1. 初次对照发现英雄标题在当前桌面宽度下换行偏密，调整 `.hero h1` 最大字号从 54 px 降至 50 px。
2. 重新捕获 `implementation-desktop-final.png`，并以 `qa-comparison-final.png` 进行等宽英雄区对照；未发现需要修复的 P0/P1/P2 差异。

## Implementation checklist

- [x] 真实产品图片与品牌资产。
- [x] 可用导航、安装要求弹窗和咨询表单。
- [x] 版本边界与兼容条件。
- [x] 桌面浏览器渲染、构建与静态站点测试。

## Follow-up polish

- [P3] 在用户提供闲鱼店铺二维码、客服方式或正式下载地址后，将“咨询信息”接入真实转化渠道。

final result: passed
