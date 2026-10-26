const path = require("path");
const { newPres, THEME, HEX, icon } = require("./master");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
const { FaStore, FaUsers, FaRobot, FaEye, FaQuoteLeft, FaIndustry, FaUserFriends, FaMobileAlt, FaMagic } = require("react-icons/fa");

const OUT = process.argv[2] || "P1-P5.pptx";

(async () => {
  const { pres, C } = newPres();
  const sh = () => ({ type: "outer", color: "000000", opacity: 0.12, blur: 6, offset: 2, angle: 90 });
  const pend = (t) => ({ text: t, options: { color: C.accent1, bold: true, highlight: HEX.accent2 } }); // 【待填】高亮

  // ---------------- P1 封面 ----------------
  pres.addSection({ title: "开场与背景" });
  let s = pres.addSlide({ masterName: "封面", sectionTitle: "开场与背景" });
  s.addText("3秒首购：让 18-25 岁年轻人在量贩零食店第一次拿起美枣王", { placeholder: "title" });
  s.addText("基于服务生态系统的货架尝鲜体验共创方案", { placeholder: "body" });
  s.addText([
    { text: "组长：【姓名】    组员：【姓名】【姓名】【姓名】", options: { breakLine: true } },
    { text: "课程：数智商业创新    2026 年【】月【】日" },
  ], { placeholder: "members" });
  // “3 秒”视觉：同心圆 + 大数字
  s.addShape(pres.shapes.OVAL, { x: 9.15, y: 1.55, w: 3.6, h: 3.6, fill: { color: C.accent1 }, objectName: "3s-ring" });
  s.addShape(pres.shapes.OVAL, { x: 9.55, y: 1.95, w: 2.8, h: 2.8, fill: { color: C.text2 }, line: { color: C.accent2, width: 3 }, objectName: "3s-core" });
  s.addText([
    { text: "3", options: { fontSize: 88, bold: true, color: C.accent2, breakLine: true } },
    { text: "秒决定拿不拿", options: { fontSize: 16, color: C.background1 } },
  ], { x: 9.55, y: 2.05, w: 2.8, h: 2.6, align: "center", valign: "middle", margin: 0, isTextBox: true, paraSpaceAfter: 0, objectName: "3s-text" });
  s.addNotes("【A｜约 10 秒】各位老师好，我们组的题目是“3秒首购”——帮美枣王让 18 到 25 岁的年轻人，在量贩零食店第一次把它拿起来。先问大家一个问题：你上次在零食很忙货架前，停了几秒？");

  // ---------------- P2 背景 ----------------
  s = pres.addSlide({ masterName: "内容页", sectionTitle: "开场与背景" });
  s.addText("渠道、人群、技术三股力量同时变了", { placeholder: "title" });
  s.addText("服务生态系统", { placeholder: "theory" });
  s.addText("鸣鸣很忙 2025 年度业绩（虎嗅、观察者网报道）；观研报告网", { placeholder: "source" });
  const cards = [
    { ic: FaStore, head: "行业｜量贩零食崛起", big: "2.19 万家", lab: "鸣鸣很忙门店数（2025 年末）",
      body: ["2025 年营收 661.7 亿元，同比 +68.2%", "平价、散称、SKU 多、上新快", "2024 年量贩零食市场约千亿元规模"] },
    { ic: FaUsers, head: "社会｜Z 世代爱“逛”", big: "尝鲜 > 品牌", lab: "18-25 岁：情绪价值优先",
      body: ["不带清单，边逛边挑，靠货架当场决定", "小红书 / 抖音“量贩必买榜”在进店前种草", [pend("【待填：面访中“愿为尝鲜买新品”占比】")]] },
    { ic: FaRobot, head: "技术｜AI + 门店数据", big: "成本↓ 测试↑", lab: "内容更便宜，验证更快",
      body: ["AI 让包装、文案、内容生成成本大幅降低", "POS / 会员数据让小范围 A/B 测试可行", "副作用：人人都能快速出新品，货架更挤"] },
  ];
  const cw = 3.81, gap = 0.35;
  for (let i = 0; i < 3; i++) {
    const c = cards[i], x = 0.6 + i * (cw + gap), y = 1.45;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: 4.5, rectRadius: 0.12, fill: { color: C.background2 }, objectName: `card-${i + 1}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.3, w: 0.62, h: 0.62, fill: { color: C.accent1 } });
    s.addImage({ data: await icon(c.ic, "FFFFFF"), x: x + 0.45, y: y + 0.45, w: 0.32, h: 0.32 });
    s.addText(c.head, { x: x + 1.05, y: y + 0.3, w: cw - 1.25, h: 0.62, fontSize: 16, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true });
    s.addText(c.big, { x: x + 0.3, y: y + 1.1, w: cw - 0.6, h: 0.8, fontSize: 34, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true });
    s.addText(c.lab, { x: x + 0.3, y: y + 1.9, w: cw - 0.6, h: 0.35, fontSize: 12, color: C.accent4, margin: 0, isTextBox: true });
    s.addText(c.body.map((b, j) => {
      const runs = Array.isArray(b) ? b : [{ text: b, options: {} }];
      return runs.map((r) => ({ text: r.text, options: { ...r.options, bullet: true, breakLine: j < c.body.length - 1 } }));
    }).flat(), { x: x + 0.3, y: y + 2.4, w: cw - 0.5, h: 1.95, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true });
  }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 6.13, w: 12.13, h: 0.62, rectRadius: 0.1, fill: { color: C.text2 }, objectName: "summary" });
  s.addText([
    { text: "小结  ", options: { bold: true, color: C.accent2 } },
    { text: "渠道、人群、技术都变了——传统枣类品牌的打法还没变", options: { color: C.background1 } },
  ], { x: 0.85, y: 6.13, w: 11.6, h: 0.62, fontSize: 16, valign: "middle", margin: 0, isTextBox: true });
  s.addNotes("【A｜约 30 秒】先看背景，三股力量叠在一起。行业上，量贩零食店在狂奔——零食很忙和赵一鸣合并后的鸣鸣很忙，2025 年底门店已经超过 2.1 万家，营收六百多亿。人群上，18 到 25 岁的年轻人进店就是“逛”，尝鲜和情绪价值排第一，很多人进店前已经被小红书的必买榜种草。技术上，AI 让做包装、做内容的成本大幅下降，门店数据又让小范围测试变得可行。渠道、人群、技术都变了，但传统枣类品牌的打法没变。");

  // ---------------- P3 问题 ----------------
  pres.addSection({ title: "提出问题" });
  s = pres.addSlide({ masterName: "内容页", sectionTitle: "提出问题" });
  s.addText("年轻人为什么不拿起第一包", { placeholder: "title" });
  s.addText("问题界定", { placeholder: "theory" });
  s.addText("企业导师访谈逐字稿；门店观察【n=】；面访【n=】", { placeholder: "source" });
  // 左：企业家原问题
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.45, w: 5.5, h: 3.05, rectRadius: 0.12, fill: { color: C.text2 }, objectName: "origin-q" });
  s.addImage({ data: await icon(FaQuoteLeft, HEX.accent2), x: 0.9, y: 1.72, w: 0.45, h: 0.45 });
  s.addText("企业家原问题", { x: 1.5, y: 1.72, w: 3.5, h: 0.45, fontSize: 14, bold: true, color: C.accent2, valign: "middle", margin: 0, isTextBox: true });
  s.addText("如何提高美枣王新零食在量贩零食店 18-25 岁消费者的首次购买率？", { x: 0.9, y: 2.35, w: 4.9, h: 1.9, fontSize: 26, bold: true, color: C.background1, valign: "top", margin: 0, isTextBox: true });
  // 左下：导师引语
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 4.75, w: 5.5, h: 1.95, rectRadius: 0.12, fill: { color: C.background2 }, objectName: "mentor-q" });
  s.addText([
    { text: "零食消费“更多是一个尝鲜”，尝了不错才会继续买。", options: { italic: true, color: C.text1, breakLine: true } },
    { text: "——企业导师 丁其磊老师", options: { fontSize: 12, color: C.accent4, breakLine: true } },
    { text: "→ 首购是复购的前提", options: { bold: true, color: C.accent1 } },
  ], { x: 0.9, y: 4.9, w: 4.95, h: 1.65, fontSize: 16, valign: "middle", margin: 0, paraSpaceAfter: 4, isTextBox: true });
  // 右上：现象层
  s.addText("现象层", { x: 6.5, y: 1.45, w: 3, h: 0.4, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true });
  const phen = [
    { ic: FaEye, h: "拿起率低，拿起又放下", t: [{ text: "门店货架观察：" }, pend("【拿起率 / 放下率】")] },
    { ic: FaQuoteLeft, h: "“枣”= 养生、送礼、长辈", t: [{ text: "面访原话：" }, pend("【贴 1 句受访者原话】")] },
  ];
  for (let i = 0; i < 2; i++) {
    const y = 1.95 + i * 1.12;
    s.addShape(pres.shapes.OVAL, { x: 6.5, y: y + 0.1, w: 0.6, h: 0.6, fill: { color: C.accent5 } });
    s.addImage({ data: await icon(phen[i].ic, HEX.accent1), x: 6.65, y: y + 0.25, w: 0.3, h: 0.3 });
    s.addText([
      { text: phen[i].h, options: { bold: true, fontSize: 16, color: C.text1, breakLine: true } },
      ...phen[i].t.map((r) => ({ text: r.text, options: { fontSize: 14, color: C.accent4, ...(r.options || {}) } })),
    ], { x: 7.3, y, w: 5.4, h: 0.85, valign: "middle", margin: 0, isTextBox: true });
  }
  // 右下：三层来源
  s.addText("问题的三层来源", { x: 6.5, y: 4.3, w: 3, h: 0.4, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true });
  const src = [
    { h: "企业 / 专家", t: "企业家提出首购率问题；导师强调“尝鲜”" },
    { h: "行业背景", t: "量贩店 SKU 多，新品曝光窗口短" },
    { h: "技术变化", t: "AI 让人人快速出新品，货架更拥挤" },
  ];
  const tw = 1.95, tg = 0.15;
  src.forEach((d, i) => {
    const x = 6.5 + i * (tw + tg);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.8, w: tw, h: 1.9, rectRadius: 0.1, fill: { color: C.background2 }, objectName: `source-${i + 1}` });
    s.addText([
      { text: d.h, options: { bold: true, fontSize: 15, color: C.accent1, breakLine: true } },
      { text: d.t, options: { fontSize: 13, color: C.text1 } },
    ], { x: x + 0.15, y: 4.95, w: tw - 0.3, h: 1.6, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true });
  });
  s.addNotes("【A｜约 35 秒】企业家给我们的原问题是：怎么提高美枣王新零食在量贩店里 18 到 25 岁人群的首次购买率。我们在门店观察到，年轻人拿起率很低，拿起来又放回去【报观察数字】；问他们“枣”让你想到什么，答案是养生、送礼、长辈【引一句原话】。这个问题有三层来源：企业提出的、量贩行业新品窗口短带来的、AI 让货架更拥挤带来的。丁其磊老师说，零食消费更多是尝鲜——首购，是复购的前提。");

  // ---------------- P4 现有方案 ----------------
  s = pres.addSlide({ masterName: "内容页", sectionTitle: "提出问题" });
  s.addText("现有做法打不到货架前那一刻", { placeholder: "title" });
  s.addText("现有方案对比", { placeholder: "theory" });
  s.addText("公开资料与组内分析；评价为组内定性判断", { placeholder: "source" });
  const GOOD = { fill: { color: HEX.accent6 }, color: HEX.accent3, bold: true };
  const BAD = { fill: { color: HEX.accent5 }, color: HEX.accent1, bold: true };
  const MID = { fill: { color: "FFF1CC" }, color: "8A5A00", bold: true };
  const cell = (t, st) => ({ text: t, options: { ...st, fill: { ...st.fill }, align: "center" } });
  const hdr = (t) => ({ text: t, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, align: "center" } });
  const rows = [
    [hdr("现有做法"), hdr("成本"), hdr("触达货架那一刻"), hdr("转变“枣=养生”"), hdr("可测量")],
    [{ text: "降价 / 买赠", options: { bold: true } }, cell("中", MID), cell("能", GOOD), cell("不能", BAD), cell("能", GOOD)],
    [{ text: "明星代言", options: { bold: true } }, cell("很高", BAD), cell("不能", BAD), cell("部分", MID), cell("难", BAD)],
    [{ text: "线上投流", options: { bold: true } }, cell("高", BAD), cell("不能", BAD), cell("部分", MID), cell("仅线上", MID)],
    [{ text: "门店堆头陈列", options: { bold: true } }, cell("中", MID), cell("能", GOOD), cell("不能", BAD), cell("部分", MID)],
  ];
  s.addTable(rows, { x: 0.6, y: 1.45, w: 12.13, colW: [2.93, 2.3, 2.3, 2.3, 2.3], rowH: 0.56, fontSize: 15, color: HEX.dk1, valign: "middle",
    border: { type: "solid", pt: 1.5, color: HEX.lt1 }, fill: { color: HEX.lt2 }, objectName: "compare-table" });
  // 结论 → 切入点
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 4.55, w: 5.6, h: 1.75, rectRadius: 0.12, fill: { color: C.background2 }, objectName: "conclusion" });
  s.addText([
    { text: "结论", options: { bold: true, fontSize: 16, color: C.text2, breakLine: true } },
    { text: "要么破坏量贩店的平价规则，要么只打到线上；没有一种能同时改变“枣=养生”的第一印象。", options: { fontSize: 16, color: C.text1 } },
  ], { x: 0.85, y: 4.55, w: 5.1, h: 1.75, valign: "middle", margin: 0, paraSpaceAfter: 6, isTextBox: true });
  s.addShape(pres.shapes.RIGHT_ARROW, { x: 6.4, y: 5.15, w: 0.6, h: 0.55, fill: { color: C.accent2 }, objectName: "arrow" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.2, y: 4.55, w: 5.53, h: 1.75, rectRadius: 0.12, fill: { color: C.accent1 }, objectName: "entry-point" });
  s.addText([
    { text: "我们的切入点", options: { bold: true, fontSize: 16, color: C.accent2, breakLine: true } },
    { text: "不加预算打广告，而是降低“第一次试”的风险和门槛。", options: { fontSize: 18, bold: true, color: C.background1 } },
  ], { x: 7.45, y: 4.55, w: 5.05, h: 1.75, valign: "middle", margin: 0, paraSpaceAfter: 6, isTextBox: true });
  s.addNotes("【A｜约 30 秒】市面上常见的做法有四种：降价买赠、明星代言、线上投流、门店堆头。我们从四个维度比较：成本、能不能打到货架前那一刻、能不能改变“枣等于养生”的印象、效果能不能测。结果是，它们要么破坏量贩店的平价规则，要么只打到线上，没有一个能同时改变第一印象。所以我们的切入点不是加预算打广告，而是降低年轻人“第一次试”的风险和门槛。");

  // ---------------- P5 问题本质：价值网络图 ----------------
  s = pres.addSlide({ masterName: "内容页", sectionTitle: "提出问题" });
  s.addText("本质：第一次尝鲜体验没有被共同创造", { placeholder: "title" });
  s.addText("价值网络＋服务主导逻辑", { placeholder: "theory" });
  s.addText("企业导师访谈逐字稿；组内分析", { placeholder: "source" });
  // 价值网络图区域
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.45, w: 7.55, h: 4.45, rectRadius: 0.12, fill: { color: C.background2 }, objectName: "network-bg" });
  s.addText("美枣王价值网络", { x: 0.85, y: 1.6, w: 3, h: 0.35, fontSize: 13, bold: true, color: C.accent4, margin: 0, isTextBox: true });
  const cy = 3.45; // 主链中心线
  const node = (txt, sub, x, w, fill, fc, name) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy - 0.48, w, h: 0.96, rectRadius: 0.1, fill: { color: fill }, shadow: sh(), objectName: name });
    s.addText([
      { text: txt, options: { bold: true, fontSize: 15, color: fc, breakLine: true } },
      { text: sub, options: { fontSize: 11, color: fc } },
    ], { x, y: cy - 0.48, w, h: 0.96, align: "center", valign: "middle", margin: 0, isTextBox: true });
  };
  const arrow = (x1, y1, x2, y2, dash) => s.addShape(pres.shapes.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1) || 0.001, h: Math.abs(y2 - y1) || 0.001,
    flipH: x2 < x1, flipV: y2 < y1,
    line: { color: dash ? HEX.accent4 : HEX.dk2, width: dash ? 1.25 : 2, dashType: dash ? "dash" : "solid", endArrowType: "triangle" },
  });
  node("美枣王", "品牌 / 工厂", 0.85, 1.45, HEX.lt1, HEX.dk2, "n-brand");
  node("零食很忙", "买手 / 门店", 2.75, 1.5, HEX.lt1, HEX.dk2, "n-store");
  // 货架决策节点（红圈）
  s.addShape(pres.shapes.OVAL, { x: 4.6, y: cy - 0.8, w: 1.6, h: 1.6, fill: { color: C.background1 }, line: { color: C.accent1, width: 3.5 }, objectName: "n-shelf" });
  s.addText([
    { text: "货架", options: { bold: true, fontSize: 15, color: C.accent1, breakLine: true } },
    { text: "决策节点", options: { bold: true, fontSize: 15, color: C.accent1 } },
  ], { x: 4.6, y: cy - 0.8, w: 1.6, h: 1.6, align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addText("品牌在这里缺席", { x: 4.3, y: cy + 0.88, w: 2.2, h: 0.35, fontSize: 12, bold: true, color: C.accent1, align: "center", margin: 0, isTextBox: true });
  node("年轻人", "18-25 岁", 6.55, 1.4, HEX.dk2, HEX.lt1, "n-youth");
  arrow(2.3, cy, 2.75, cy); arrow(4.25, cy, 4.6, cy); arrow(6.2, cy, 6.55, cy);
  // 上：内容平台 / KOC；下：AI 工具商
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.0, y: 1.6, w: 2.9, h: 0.7, rectRadius: 0.1, fill: { color: C.background1 }, line: { color: C.accent4, width: 1, dashType: "dash" }, objectName: "n-content" });
  s.addImage({ data: await icon(FaMobileAlt, HEX.accent4), x: 5.15, y: 1.8, w: 0.3, h: 0.3 });
  s.addText("内容平台 / 校园 KOC", { x: 5.5, y: 1.6, w: 2.35, h: 0.7, fontSize: 13, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  arrow(7.25, 2.3, 7.25, cy - 0.48, true);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.35, y: 4.85, w: 2.6, h: 0.7, rectRadius: 0.1, fill: { color: C.background1 }, line: { color: C.accent4, width: 1, dashType: "dash" }, objectName: "n-ai" });
  s.addImage({ data: await icon(FaMagic, HEX.accent4), x: 1.5, y: 5.05, w: 0.3, h: 0.3 });
  s.addText("AI 工具商", { x: 1.85, y: 4.85, w: 2.0, h: 0.7, fontSize: 13, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  arrow(1.75, 4.85, 1.75, cy + 0.48, true); arrow(3.55, 4.85, 3.55, cy + 0.48, true);
  // 右：两条本质
  const ess = [
    { tag: "本质一｜价值网络", t: "量贩店是“买手”，品牌对终端控制力弱——首购问题不是品牌一方能解决的。" },
    { tag: "本质二｜服务主导逻辑", t: "价值在“吃第一口”时才产生。首购低 = 感知风险 > 预期价值，而不是产品不好。" },
  ];
  ess.forEach((e, i) => {
    const y = 1.45 + i * 2.3;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.45, y, w: 4.28, h: 2.1, rectRadius: 0.12, fill: { color: C.background1 }, line: { color: C.accent5, width: 1.5 }, shadow: sh(), objectName: `essence-${i + 1}` });
    s.addText([
      { text: e.tag, options: { bold: true, fontSize: 15, color: C.accent1, breakLine: true } },
      { text: e.t, options: { fontSize: 17, color: C.text1 } },
    ], { x: 8.7, y: y + 0.15, w: 3.8, h: 1.8, valign: "middle", margin: 0, paraSpaceAfter: 6, isTextBox: true });
  });
  // 一句话定义
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 6.08, w: 12.13, h: 0.7, rectRadius: 0.1, fill: { color: C.text2 }, objectName: "definition" });
  s.addText([
    { text: "一句话定义  ", options: { bold: true, color: C.accent2 } },
    { text: "首购率低，是美枣王、零食很忙、年轻人之间没有共同创造“第一次尝鲜体验”的问题", options: { color: C.background1 } },
  ], { x: 0.85, y: 6.08, w: 11.7, h: 0.7, fontSize: 16, valign: "middle", margin: 0, isTextBox: true });
  s.addNotes("【A｜约 40 秒】那问题的本质是什么？我们画了美枣王的价值网络：美枣王把货卖给零食很忙，零食很忙把它摆上货架，年轻人在货架前做决定。旁边还有内容平台和 AI 工具商。关键在这个红圈——货架前那几秒，品牌是缺席的。第一，从价值网络看，量贩店是“买手”，品牌对终端控制力很弱，首购不是品牌一家能解决的。第二，从服务主导逻辑看，价值在吃第一口时才产生，首购低是因为感知风险大于预期价值，不是产品不好。所以我们的定义是：首购率低，是三方没有共同创造“第一次尝鲜体验”。这个判断对不对？我们去门店和同龄人那里验证了一下。");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
