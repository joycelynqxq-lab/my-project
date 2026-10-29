// 制作1 负责页面的演讲稿（第4版大纲）：A 讲开场至 P5，B 讲 PART02 过渡，C 讲 PART03、P14、P15
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle } = require("docx");

const FONT = "Microsoft YaHei";
const RED = "7A1426", GREY = "6E6266";
const RATE = 230; // 字/分钟
const secs = (ps) => Math.round(ps.join("").replace(/\*\*/g, "").replace(/[，。、：；？！“”‘’（）—\s]/g, "").length / RATE * 60);

const runs = (s, base = {}) => s.split(/(\*\*[^*]+\*\*|【[^】]*】)/).filter(Boolean).map((t) =>
  t.startsWith("**") ? new TextRun({ text: t.slice(2, -2), bold: true, color: RED, ...base })
  : t.startsWith("【") ? new TextRun({ text: t, highlight: "yellow", ...base }) : new TextRun({ text: t, ...base }));
const P = (s) => new Paragraph({ children: runs(s), spacing: { after: 160, line: 400 } });
const note = (s) => new Paragraph({ children: [new TextRun({ text: "提示：", bold: true, color: GREY, size: 20 }), ...runs(s, { color: GREY, size: 20 })],
  spacing: { before: 60, after: 240 }, shading: { type: ShadingType.CLEAR, fill: "F5EEF0" }, indent: { left: 120, right: 120 } });

const CORE = "如何让 18–25 岁年轻人在量贩零食店第一次尝鲜美枣王，并通过首购体验带动复购？";
const slides = [
  { who: "A", t: "封面", p: [`老师好，我们组要回答的核心问题是：**${CORE}**`], n: "核心问题全组统一，一个字都不要改。说完停一拍再翻页。" },
  { who: "A", t: "目录", p: ["我们分三部分讲：先找问题，再给方案，最后讲怎么落地。"], n: "带过即可，不要逐条念。" },
  { who: "A", t: "PART 01 背景与问题", p: ["先看背景与问题。"], n: "一句话即翻页。" },
  { who: "A", t: "P2 量贩零食正在爆发，年轻人“尝鲜大于品牌”", p: [
    "先看两个事实：行业龙头门店超过 **2 万家**，头部两家占了 **七成半** 份额——品牌想找年轻人，就得进量贩店。",
    "再看人：15 个同龄人里，**9 个吃完会查配料或含糖量**。年轻人尝鲜大于品牌，也越来越在意健康——**机会很大，但卖法要变。**",
  ], n: "讲到 75.5% 和 9/15 两个大数字时各停半拍。“糖必标”新国标在页面上，不用念；老师问起再说：2027 年 3 月起包装要强制标示糖。数字出处在页面右下角：龙头业绩公告、万辰集团招股书、GB 28050-2025。" },
  { who: "A", t: "P3 年轻人四步买零食，美枣王步步掉队", p: [
    "【换成你自己的真实经历】上周我刷小红书看到一款新零食，第二天就去校门口买了一小包，吃完先翻配料表。",
    "这就是年轻人买零食的四步：**种草、尝鲜、验证、复购**。",
    "而美枣王**步步掉队**：线上没人种草，店里找不到，枣像养生品，包装不回答健康问题，也没有再买的理由。",
  ], n: "第一句必须是演讲人自己真实发生过的事，老师很可能追问。讲四步时手从左往右指着路径图走。" },
  { who: "A", t: "P4 难的不是首购，而是让首购变成复购", p: [
    `所以我们把问题重新定义为——**${CORE}**`,
    "只追求第一次买意义不大，要在第一次体验里就埋好再买的理由。",
    "本质有两条：一是**线上种草、线下货架、产品体验三段断开**；二是**品牌、门店、内容创作者各做各的**，没有一起创造第一次尝鲜。",
  ], n: "全篇转折点，语速放慢，核心问题原句念出。时间富余可以加一句：丁其磊老师也说，买零食更多的是尝鲜，口感不差就会继续再买。" },
  { who: "A", t: "P5 我们的调研：4 种方法、1 条证据链", p: [
    "这些判断来自四种方法：同龄人访谈、小红书笔记分析、企业导师访谈和门店走访；问卷留到第二步前做验证。",
    "**问题找到了，下面看我们设计的闭环。**",
  ], n: "最后一句是交接词，说完把翻页笔交给 B。" },
  { who: "B", t: "PART 02 解决方案", p: ["下面看我们的解决方案。"], n: "B 接过翻页笔后的第一句，一句话即翻页。" },
  { who: "C", t: "PART 03 执行与收尾", p: ["方案讲完了，接下来讲怎么一步步落地。"], n: "C 接过翻页笔后的第一句，一句话即翻页。" },
  { who: "C", t: "P14 总结：从一次尝鲜到长期复购", p: [
    "回到我们的核心问题。问题是**首购没有变成复购**；方案是**种草、尝鲜、验证、复购的闭环**；落地从一所学校开始，**4 周、约 2600 元**，跑通就能复制。",
    "从一所学校，到全国量贩渠道。",
  ], n: "2600 元以制作 4 的 P12 最终数字为准，改了这里也要改。" },
  { who: "C", t: "P15 封底", p: ["**让年轻人第一口就爱上，第二包自然会来。**谢谢老师！"], n: "说完停在这页，等老师提问。" },
];
slides.forEach((s) => (s.sec = secs(s.p)));
const fmt = (x) => (x >= 60 ? `${Math.floor(x / 60)} 分 ${x % 60} 秒` : `${x} 秒`);
const tot = (w) => slides.filter((s) => s.who === w).reduce((a, s) => a + s.sec, 0);

const W = [800, 6026, 1100, 1100];
const bd = { style: BorderStyle.SINGLE, size: 4, color: "D9C7CB" };
const cell = (t, w, head, bold) => new TableCell({ width: { size: w, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: head ? RED : bold ? "F5EEF0" : "FFFFFF" },
  borders: { top: bd, bottom: bd, left: bd, right: bd }, margins: { top: 70, bottom: 70, left: 110, right: 110 },
  children: [new Paragraph({ children: [new TextRun({ text: t, bold: head || bold, color: head ? "FFFFFF" : "3A3536", size: 21 })] })] });
const rows = [["演讲者", "页面", "约时长"].length && ["序", "页面", "演讲者", "约时长"]];
slides.forEach((s, i) => rows.push([String(i + 1), s.t, s.who, `${s.sec} 秒`]));
const sumRows = [["", "演讲 A 合计（封面—P5）", "A", fmt(tot("A"))], ["", "演讲 B（PART 02）", "B", fmt(tot("B"))], ["", "演讲 C（PART 03、P14、P15）", "C", fmt(tot("C"))]];
const table = new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: W, rows: [
  ...rows.map((r, k) => new TableRow({ tableHeader: k === 0, children: r.map((t, i) => cell(t, W[i], k === 0)) })),
  ...sumRows.map((r) => new TableRow({ children: r.map((t, i) => cell(t, W[i], false, true)) })),
] });

const qa = [
  ["15 个人的访谈够不够？", "现阶段是找问题，用的是访谈和经验萃取，看的是“为什么”。丁其磊老师也提醒，面对面问出的十份比发出去的一百份更可靠；比例要靠第二步前约 200 份问卷验证，问卷的每道题都对应一个方案决定。"],
  ["行业数据是哪来的？", "龙头 2025 年度业绩公告、万辰集团招股书，营养标签新国标 GB 28050-2025；页面右下角都标了来源。"],
  ["为什么不只研究首购？", "年轻人本来就爱尝鲜，只尝不留，首购提升没有意义；所以要在第一次体验里就埋好复购的理由。"],
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: { ascii: FONT, eastAsia: FONT, hAnsi: FONT }, size: 24, color: "2A1A1D" } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal", run: { size: 40, bold: true, color: RED }, paragraph: { spacing: { after: 120 } } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 30, bold: true, color: RED }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, color: RED }, paragraph: { spacing: { before: 320, after: 60 }, outlineLevel: 1 } },
    ],
  },
  sections: [{ properties: { page: { margin: { top: 1300, bottom: 1300, left: 1440, right: 1440 } } }, children: [
    new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun("一口尝鲜，一路复购｜演讲稿（制作1负责页面·最终版）")] }),
    new Paragraph({ children: [new TextRun({ text: "按第 4 版 PPT 大纲；封面、目录、三张分隔页、P2–P5、P14、P15", color: GREY })], spacing: { after: 200 } }),
    P("按每分钟约 230 字估算时长。讲稿是口语，不要逐字念 PPT；**红色加粗**的句子放慢、加重；黄色【】处要换成演讲人自己的内容。"),
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("时长总表")] }),
    table,
    new Paragraph({ children: [new TextRun({ text: "大纲给演讲 A 的 P1–P5 是 2 分钟，目录和分隔页是新加的，所以 A 合计控制在 2 分钟左右。彩排超时先压 P3 的个人经历。", color: GREY, size: 20 })], spacing: { before: 120 } }),
    ...["A", "B", "C"].flatMap((w) => [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(w === "A" ? "演讲 A：开场到 P5" : w === "B" ? "演讲 B：方案部分的开头" : "演讲 C：执行部分的开头与收尾")] }),
      ...slides.filter((s) => s.who === w).flatMap((s) => [
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(s.t)] }),
        new Paragraph({ children: [new TextRun({ text: `约 ${s.sec} 秒`, color: GREY, size: 20 })], spacing: { after: 120 } }),
        ...s.p.map(P), note(s.n),
      ]),
    ]),
    new Paragraph({ heading: HeadingLevel.HEADING_1, keepNext: true, children: [new TextRun("可能被问到的问题（A 负责的部分）")] }),
    ...qa.flatMap(([q, a]) => [
      new Paragraph({ keepNext: true, children: [new TextRun({ text: "问：" + q, bold: true })], spacing: { before: 120, after: 60 } }),
      new Paragraph({ children: [new TextRun({ text: "答：" + a })], spacing: { after: 160, line: 360 } }),
    ]),
  ] }],
});
fs.writeFileSync(__dirname + "/notes.json", JSON.stringify(slides.map((x) => `【${x.who}｜约 ${x.sec} 秒】` + x.p.join("").replace(/\*\*/g, "") + "\n提示：" + x.n)));
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(process.argv[2], b); console.log("A", tot("A"), "B", tot("B"), "C", tot("C")); });
