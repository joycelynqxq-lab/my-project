const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType, BorderStyle } = require("docx");

const FONT = "Microsoft YaHei";
const RED = "7A1426", GREY = "6E6266";

// "**加粗**" → 粗体红字
const runs = (s, base = {}) => s.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((t) =>
  t.startsWith("**") ? new TextRun({ text: t.slice(2, -2), bold: true, color: RED, ...base }) : new TextRun({ text: t, ...base }));
const P = (s, o = {}) => new Paragraph({ children: runs(s), spacing: { after: 160, line: 380 }, ...o });
const note = (s) => new Paragraph({ children: [new TextRun({ text: "提示：", bold: true, color: GREY, size: 20 }), ...runs(s, { color: GREY, size: 20 })], spacing: { before: 60, after: 240 },
  shading: { type: ShadingType.CLEAR, fill: "F5EEF0" }, indent: { left: 120, right: 120 } });
const H = (t, meta) => [
  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t })], spacing: { before: 360, after: 80 } }),
  new Paragraph({ children: [new TextRun({ text: meta, color: GREY, size: 20 })], spacing: { after: 160 } }),
];

const slides = [
  { t: "第 1 页｜封面", m: "约 10 秒", p: [
    "各位老师、同学好，我们组的题目是**“3 秒首购”**——帮美枣王，让 18 到 25 岁的年轻人在量贩零食店里，第一次把它拿起来。",
    "先问大家一个问题：**你上次站在零食很忙的货架前，停了几秒？**",
  ], n: "问完停一拍，看向台下，再翻页。" },
  { t: "第 2 页｜目录", m: "约 5 秒", p: [
    "我们按**提出问题、分析问题、解决问题**三步来讲，重点是第三部分的“3 秒首购计划”。",
  ], n: "带过即可，不要逐条念四个章节。" },
  { t: "第 3 页｜Part 01 背景与问题", m: "约 3 秒", p: ["先看第一部分，背景与问题。"], n: "一句话即翻页。" },
  { t: "第 4 页｜渠道、人群、技术同时改变", m: "约 30 秒", p: [
    "这个问题为什么现在变得重要？因为有三股力量叠在了一起。",
    "**渠道在狂奔。**零食很忙和赵一鸣合并成的鸣鸣很忙，2025 年底门店已经有 **2.19 万家**，一年营收 661.7 亿元，同比增长 68%。平价、散称、上新快，是它的规则。",
    "**人群变了。**18 到 25 岁的年轻人进店就是“逛”，不带清单，靠货架当场决定；很多人进店前，就已经被小红书、抖音的“量贩必买榜”种草了。",
    "**技术也变了。**AI 让做包装、做文案的成本大幅下降，门店的收银和会员数据，又让小范围的 A/B 测试变得可行。副作用是，人人都能快速出新品，货架越来越挤。",
    "渠道、人群、技术都变了——**但传统枣类品牌的打法，还没变。**",
  ], n: "讲到三个大字（2.19 万家、尝鲜 > 品牌、成本↓测试↑）时各停半拍。数据来自鸣鸣很忙 2025 年度业绩（虎嗅、观察者网报道）。" },
  { t: "第 5 页｜年轻人为什么不拿起第一包", m: "约 35 秒", p: [
    "企业家给我们的原问题是：**怎么提高美枣王新零食在量贩店里，18 到 25 岁人群的首次购买率。**",
    "我们先看到两个现象。",
    "第一，**拿起率较低。**一家量贩店有近 2000 个商品，每个月还有一到两成在替换；新品如果没人拿起来，很快就会被末位淘汰。",
    "第二，在年轻人心里，**枣 = 养生、送礼、食药同源。**尼尔森的研究里，红枣和枸杞、山药一样，是大家最认定的“食药同源”食材——它更像补品，而不是零食。",
    "这个问题有三层来源：企业提出的首购问题，量贩行业新品曝光窗口短，还有 AI 让人人都能快速出新品、货架更拥挤。正如企业导师丁其磊老师说的，零食消费“更多是一个尝鲜”，尝了不错才会继续买——**首购，是复购的前提。**",
  ], n: "如果门店观察和面访数据出来了，可以在“拿起率较低”后加一句“我们观察到经过的 x 个年轻人里只有 y 个拿起来”。注意说“年轻人对枣的第一印象”，不要说“年轻人不买枣”，线上数据显示 85–95 后其实是枣类零食的主力买家。SKU 数据来自华创证券零食量贩报告，红枣认知来自尼尔森《食药同源白皮书》。" },
  { t: "第 6 页｜现有做法分析", m: "约 30 秒", p: [
    "那市面上常见的做法管用吗？我们比较了四种：降价买赠、明星代言、线上投流、门店堆头陈列。",
    "看四个维度：成本高不高、能不能打到货架前那一刻、能不能改变“枣 = 养生”的印象、效果能不能测。",
    "表里红色的格子很多。降价和堆头能打到货架，但改不了印象，降价还会碰量贩店的平价规则；代言和投流又只打到线上。**没有一种能同时改变第一印象。**",
    "所以我们的切入点是：**不加预算打广告，而是降低“第一次试”的风险和门槛。**",
  ], n: "讲“红色的格子”时可用激光笔划一下“触达货架那一刻”和“转变枣=养生”两列。超时就删掉第三段中间逐项解释的那句。" },
  { t: "第 7 页｜本质：第一次尝鲜体验没有被共同创造", m: "约 45 秒", p: [
    "那问题的本质是什么？我们画了美枣王的价值网络。",
    "美枣王供货给零食很忙，零食很忙作为买手决定怎么陈列，年轻人在货架前决定拿不拿。进店前，小红书和抖音的必买榜已经在种草；AI 工具又让新品越来越多、货架越来越挤。",
    "关键在这个**红圈**。货架前只有几秒，年轻人卡在三个地方：看不懂口味、一包太大怕踩雷、觉得枣是长辈吃的。**而品牌在这一刻是缺席的。**",
    "第一，从**价值网络**看，量贩店是“买手”，选什么、摆哪里、卖多少钱都由渠道决定，品牌在货架前说不上话，所以首购必须和渠道一起解决。",
    "第二，从**服务主导逻辑**看，价值在吃第一口时才产生。首购低，是因为感知风险大于预期价值，不是产品不好。",
    "所以我们的定义是：**首购率低，是美枣王、零食很忙、年轻人三方没有共同创造“第一次尝鲜体验”。**",
    "这个判断对不对？我们去门店和同龄人那里验证了一下。",
  ], n: "讲价值网络时按箭头方向从左往右指，讲到“红圈”时停在圆上。最后一句是交接词，说完交给下一位讲调研部分。" },
];

const timing = [["1", "封面", "10 秒"], ["2", "目录", "5 秒"], ["3", "Part 01 背景与问题", "3 秒"], ["4", "渠道、人群、技术同时改变", "30 秒"],
  ["5", "年轻人为什么不拿起第一包", "35 秒"], ["6", "现有做法分析", "30 秒"], ["7", "本质：第一次尝鲜体验没有被共同创造", "45 秒"], ["", "合计", "约 2 分 40 秒"]];
const W = [1000, 5826, 2200];
const bd = { style: BorderStyle.SINGLE, size: 4, color: "D9C7CB" };
const cell = (t, w, head, last) => new TableCell({ width: { size: w, type: WidthType.DXA },
  shading: { type: ShadingType.CLEAR, fill: head ? RED : last ? "F5EEF0" : "FFFFFF" },
  borders: { top: bd, bottom: bd, left: bd, right: bd }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
  children: [new Paragraph({ children: [new TextRun({ text: t, bold: head || last, color: head ? "FFFFFF" : "2A1A1D" })] })] });
const table = new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: W, rows: [
  new TableRow({ tableHeader: true, children: ["页", "内容", "时长"].map((t, i) => cell(t, W[i], true)) }),
  ...timing.map((r, k) => new TableRow({ children: r.map((t, i) => cell(t, W[i], false, k === timing.length - 1)) })),
] });

const qa = [
  ["为什么只盯量贩零食店，不做电商？", "作业要求聚焦一个价值创造活动；量贩店是年轻人“逛”出来的首购场景，鸣鸣很忙已有 2.19 万家店，场景足够大。"],
  ["年轻人不是也在线上买枣吗？", "对，所以我们说的是“第一印象”问题：线上有人买，到了量贩货架前却没被拿起，卡点在货架那一刻。"],
  ["对比表里的“能 / 不能”是怎么判断的？", "是小组根据公开资料做的定性判断，标准是能否同时覆盖货架那一刻和认知转变两个条件。"],
  ["价值网络和服务主导逻辑分别解决什么？", "价值网络解释“为什么品牌一方解决不了”，服务主导逻辑解释“为什么问题出在第一次体验，而不是产品本身”。"],
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: { ascii: FONT, eastAsia: FONT, hAnsi: FONT }, size: 24, color: "2A1A1D" } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal", run: { size: 40, bold: true, color: RED }, paragraph: { spacing: { after: 120 } } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 30, bold: true, color: RED }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, color: RED }, paragraph: { outlineLevel: 1 } },
    ],
  },
  sections: [{ properties: { page: { margin: { top: 1300, bottom: 1300, left: 1440, right: 1440 } } }, children: [
    new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun("3 秒首购｜演讲稿")] }),
    new Paragraph({ children: [new TextRun({ text: "数智商业创新 小组作业 · 第 1–7 页（背景与问题）", color: GREY })], spacing: { after: 240 } }),
    P("按每分钟 220–240 字的语速写，合计约 2 分 40 秒。讲稿是口语，不要逐字念 PPT 上的字；**红色加粗**的句子要放慢、加重。"),
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("时长总表")] }),
    table,
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("逐页讲稿")], spacing: { before: 480, after: 80 } }),
    ...slides.flatMap((s) => [...H(s.t, `${s.m}`), ...s.p.map((x) => P(x)), note(s.n)]),
    new Paragraph({ heading: HeadingLevel.HEADING_1, keepNext: true, children: [new TextRun("可能被问到的问题")] }),
    ...qa.flatMap(([q, a]) => [
      new Paragraph({ children: [new TextRun({ text: "问：" + q, bold: true })], spacing: { before: 120, after: 60 }, keepNext: true }),
      new Paragraph({ children: [new TextRun({ text: "答：" + a })], spacing: { after: 160, line: 360 } }),
    ]),
  ] }],
});

Packer.toBuffer(doc).then((b) => { fs.writeFileSync(process.argv[2], b); console.log("ok"); });
