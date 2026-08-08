import { useState } from "react";
import agentLogo from "./assets/agent-logo.png";
import excelPanel from "./assets/excel-agent-panel.png";

const editions = [
  {
    name: "免费版",
    audience: "日常处理表格，想先体验 AI 协助的个人用户",
    highlights: "公式、图表、透视表与基础读写",
  },
  {
    name: "标准分析版",
    audience: "需要经常完成分析与汇报的办公人员",
    highlights: "免费版 + 数据体检、趋势与对比分析",
  },
  {
    name: "专业自动化版",
    audience: "需要搭建分析流程的运营、财务与数据人员",
    highlights: "标准版 + 看板、Power Query、Power Pivot / DAX",
  },
  {
    name: "自动化交付版",
    audience: "需要为客户交付标准化 Excel 流程的服务者",
    highlights: "专业版 + 受控 VBA 与交付支持",
  },
];

const requirements = [
  "Windows 10 或 Windows 11",
  "Microsoft 365、Office 2019、Office 2021 或更高版本的桌面版 Excel",
  "首次安装时的网络连接与安装权限，用于下载所需的 Microsoft 运行组件",
  "你自己的大模型 API Key；插件不会内置卖家的 Key、Token 或第三方账号",
];

function scrollToSection(id) {
  document.getElementById(id)?.scrollIntoView({ behavior: "smooth", block: "start" });
}

export function App() {
  const [dialog, setDialog] = useState(null);
  const [submitted, setSubmitted] = useState(false);
  const [form, setForm] = useState({ name: "", scenario: "数据分析与报表", detail: "" });

  const updateForm = (event) => {
    const { name, value } = event.target;
    setForm((current) => ({ ...current, [name]: value }));
  };

  const submitConsultation = (event) => {
    event.preventDefault();
    setSubmitted(true);
  };

  const closeDialog = () => {
    setDialog(null);
    setSubmitted(false);
  };

  return (
    <main>
      <header className="site-header">
        <button className="brand" onClick={() => scrollToSection("top")} aria-label="返回顶部">
          <img src={agentLogo} alt="Agent for Excel" />
          <span>Agent for Excel</span>
        </button>
        <nav aria-label="页面导航">
          <button onClick={() => scrollToSection("capabilities")}>功能概览</button>
          <button onClick={() => scrollToSection("scenarios")}>适用场景</button>
          <button onClick={() => scrollToSection("editions")}>版本选择</button>
          <button onClick={() => setDialog("requirements")}>使用说明</button>
        </nav>
      </header>

      <section className="hero section-frame" id="top">
        <div className="hero-copy">
          <p className="eyebrow">Windows 版 Excel AI 助手（VSTO 插件）</p>
          <h1>在你熟悉的 Excel 中，<strong>用 AI 快速搞定数据分析</strong></h1>
          <p className="hero-lead">不学复杂工具，不切换平台。在 Excel 里直接提问、写公式、做图表、做透视表，让分析更快、更准、更省心。</p>
          <div className="hero-actions">
            <button className="button button-primary" onClick={() => setDialog("consult")}>咨询获取安装包</button>
            <button className="button button-secondary" onClick={() => setDialog("requirements")}>查看安装要求</button>
          </div>
          <p className="microcopy">当前为人工咨询交付；不包含在线付款、账号或自动续订。</p>
        </div>
        <div className="hero-visual" aria-label="Excel 内的 Agent for Excel 助手界面">
          <img src={excelPanel} alt="Agent for Excel 在 Excel 右侧任务窗格中提供数据分析模板" />
        </div>
      </section>

      <section className="trust-strip" aria-label="使用前提">
        <p><b>仅支持 Windows 桌面版 Excel</b><span>不支持网页版 Excel 与 macOS Excel</span></p>
        <p><b>需要你自己的 API Key</b><span>可按自己的模型服务配置</span></p>
        <p><b>在 Excel 内完成工作</b><span>按权限读取当前工作簿上下文</span></p>
      </section>

      <section className="section-frame capabilities" id="capabilities">
        <div className="section-heading">
          <p className="eyebrow">从问题到洞察，快人一步</p>
          <h2>把重复的表格操作，变成一句清楚的需求</h2>
          <p>从当前工作簿出发，先理解你的问题，再给出可执行、可核验的分析建议。</p>
        </div>
        <div className="capability-list">
          <article>
            <span>01</span>
            <h3>先读懂数据</h3>
            <p>检查结构、字段与质量问题，确认分析范围。</p>
          </article>
          <article>
            <span>02</span>
            <h3>再拆解思路</h3>
            <p>把含糊需求变成公式、图表、透视表或分析视图。</p>
          </article>
          <article>
            <span>03</span>
            <h3>最后交付结果</h3>
            <p>在 Excel 内看到结果，并保留下一步可复用的方法。</p>
          </article>
        </div>
      </section>

      <section className="section-frame scenario-section" id="scenarios">
        <div className="scenario-copy">
          <p className="eyebrow">适用场景</p>
          <h2>不是替你“盲目自动化”，而是帮你更快完成关键一步。</h2>
          <p>适合销售明细整理、经营数据对比、财务报表分析、周报月报制作，以及需要反复处理 Excel 的数据服务工作。</p>
        </div>
        <dl className="scenario-list">
          <div><dt>表格又乱又难用</dt><dd>先识别字段、缺失值和格式问题，再决定怎么处理。</dd></div>
          <div><dt>公式会写但效率低</dt><dd>用自然语言说明目标，获得可解释的公式建议。</dd></div>
          <div><dt>图表和透视表总差一步</dt><dd>从分析目标出发，而不是从菜单里盲目试选项。</dd></div>
          <div><dt>分析做完难复用</dt><dd>把可复用的分析流程沉淀在你原来的 Excel 工作方式里。</dd></div>
        </dl>
      </section>

      <section className="section-frame editions" id="editions">
        <div className="section-heading">
          <p className="eyebrow">版本选择</p>
          <h2>按你的工作重点，选择更合适的协作边界</h2>
          <p>先从最常用的任务开始。版本能力会同时体现在功能中心和工具执行层。</p>
        </div>
        <div className="edition-table" role="table" aria-label="Agent for Excel 版本能力">
          <div className="edition-row edition-labels" role="row">
            <span role="columnheader">版本</span><span role="columnheader">适合谁</span><span role="columnheader">核心能力</span>
          </div>
          {editions.map((edition) => (
            <div className="edition-row" role="row" key={edition.name}>
              <strong role="cell">{edition.name}</strong>
              <span role="cell">{edition.audience}</span>
              <span role="cell">{edition.highlights}</span>
            </div>
          ))}
        </div>
        <button className="button button-secondary centered-button" onClick={() => setDialog("consult")}>咨询适合我的版本</button>
      </section>

      <section className="section-frame final-cta">
        <div>
          <p className="eyebrow">开始之前</p>
          <h2>先确认你的 Excel 环境，再拿到适合你的安装包。</h2>
          <p>我们会先确认你的 Office 版本、使用场景和想要处理的数据任务，再给出合适的版本建议。</p>
        </div>
        <button className="button button-primary" onClick={() => setDialog("consult")}>咨询获取安装包</button>
      </section>

      <footer>
        <img src={agentLogo} alt="" />
        <p>Agent for Excel 是运行在 Windows 桌面版 Excel 中的 VSTO 加载项。</p>
      </footer>

      {dialog === "requirements" && (
        <section className="dialog-backdrop" role="presentation" onMouseDown={closeDialog}>
          <div className="dialog" role="dialog" aria-modal="true" aria-labelledby="requirements-title" onMouseDown={(event) => event.stopPropagation()}>
            <p className="eyebrow">安装要求</p>
            <h2 id="requirements-title">请在咨询前确认这些条件</h2>
            <ul>{requirements.map((item) => <li key={item}>{item}</li>)}</ul>
            <p className="dialog-note">不支持 Excel 网页版、macOS Excel，或受企业策略禁止 VSTO 的设备。</p>
            <div className="dialog-actions">
              <button className="button button-primary" onClick={() => setDialog("consult")}>我满足要求，继续咨询</button>
              <button className="button button-secondary" onClick={closeDialog}>关闭</button>
            </div>
          </div>
        </section>
      )}

      {dialog === "consult" && (
        <section className="dialog-backdrop" role="presentation" onMouseDown={closeDialog}>
          <div className="dialog consultation" role="dialog" aria-modal="true" aria-labelledby="consult-title" onMouseDown={(event) => event.stopPropagation()}>
            {!submitted ? <>
              <p className="eyebrow">获取安装包</p>
              <h2 id="consult-title">先告诉我，你想用 Excel 完成什么？</h2>
              <form onSubmit={submitConsultation}>
                <label>怎么称呼你<input name="name" value={form.name} onChange={updateForm} placeholder="例如：小林" required /></label>
                <label>主要场景<select name="scenario" value={form.scenario} onChange={updateForm}><option>数据分析与报表</option><option>公式、图表与透视表</option><option>Power Query / Power Pivot</option><option>受控 VBA 与自动化交付</option></select></label>
                <label>补充需求<textarea name="detail" value={form.detail} onChange={updateForm} placeholder="例如：每周要整理销售明细并做趋势分析" rows="4" /></label>
                <div className="dialog-actions"><button className="button button-primary" type="submit">生成咨询信息</button><button className="button button-secondary" type="button" onClick={closeDialog}>暂不咨询</button></div>
              </form>
            </> : <>
              <p className="eyebrow">咨询信息已整理</p>
              <h2>下一步，把这段信息发送给卖家即可。</h2>
              <p className="consult-summary">称呼：{form.name}<br />场景：{form.scenario}<br />需求：{form.detail || "暂无补充"}</p>
              <p className="dialog-note">这是静态推广页演示，不会上传或保存你的内容。</p>
              <div className="dialog-actions"><button className="button button-primary" onClick={() => navigator.clipboard?.writeText(`称呼：${form.name}\n场景：${form.scenario}\n需求：${form.detail || "暂无补充"}`)}>复制咨询信息</button><button className="button button-secondary" onClick={closeDialog}>关闭</button></div>
            </>}
          </div>
        </section>
      )}
    </main>
  );
}
