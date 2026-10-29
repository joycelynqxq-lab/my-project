// 制作1 负责的页面：封面、目录、三张分隔页、P2-P5、P14、P15
const { newPres, THEME, HEX, icon } = require("./master3");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
const { FaBalanceScale, FaStore, FaUsers, FaRobot, FaComments, FaHashtag, FaUserTie, FaWalking, FaExclamationTriangle, FaSyncAlt, FaFlag, FaQuoteLeft, FaIndustry, FaPenNib } = require("react-icons/fa");
const OUT = process.argv[2] || "mine.pptx";
const CORE = "如何让 18–25 岁年轻人在量贩零食店第一次尝鲜美枣王，并通过首购体验带动复购？";

(async () => {
  const { pres, C } = newPres();
  const T = (s, txt, o) => s.addText(txt, { margin: 0, isTextBox: true, ...o });
  const card = (s, x, y, w, h, fill, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.1, fill: { color: fill || HEX.lt1 }, line: { color: HEX.accent4, width: 0.75 }, objectName: name });
  const dot = async (s, ic, x, y, d, fill, fg) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill } });
    s.addImage({ data: await icon(ic, fg), x: x + d * 0.27, y: y + d * 0.27, w: d * 0.46, h: d * 0.46 });
  };

  // ---------- P1 封面 ----------
  pres.addSection({ title: "开场" });
  let s = pres.addSlide({ masterName: "封面", sectionTitle: "开场" });
  s.addText([{ text: "一口尝鲜，一路复购", options: { breakLine: true } }, { text: "美枣王在量贩零食店的年轻用户增长方案", options: { fontSize: 22 } }], { placeholder: "title" });
  s.addText(CORE, { placeholder: "body" });
  s.addText([
    { text: "《数智商业模式创新》小组汇报", options: { bold: true, breakLine: true } },
    { text: "第【】组　组长：【】　组员：【】【】【】【】【】【】", options: { breakLine: true } },
    { text: "2026 年【】月【】日" },
  ], { placeholder: "members" });
  s.addNotes("【A｜约 10 秒】老师好，我们组要回答的核心问题是：如何让 18–25 岁年轻人在量贩零食店第一次尝鲜美枣王，并通过首购体验带动复购？");

  // ---------- 目录 ----------
  s = pres.addSlide({ masterName: "分隔页", sectionTitle: "开场" });
  s.addText("目录", { placeholder: "title" });
  s.addText("CONTENTS", { placeholder: "part" });
  s.addText("提出问题 · 解决方案 · 落地执行", { placeholder: "body" });
  const parts = [["01", "背景与问题", "行业趋势 · 购买路径 · 核心问题 · 调研"], ["02", "解决方案", "种草—尝鲜—验证—复购闭环"], ["03", "执行与收尾", "三步走 · 首轮成本 · 老板选项"]];
  parts.forEach(([n, h, d], i) => {
    const y = 5.5 + i * 0.6, x = 0.7;
    T(s, n, { x, y, w: 0.7, h: 0.5, fontSize: 22, bold: true, color: C.accent5, valign: "middle" });
    T(s, h, { x: x + 0.75, y, w: 1.9, h: 0.5, fontSize: 16, bold: true, color: C.background1, valign: "middle" });
    T(s, d, { x: x + 2.7, y, w: 3.9, h: 0.5, fontSize: 12, color: C.background2, valign: "middle" });
  });
  s.addNotes("【A｜约 5 秒，带过】我们分三部分讲：先找问题，再给方案，最后讲怎么落地。");

  // ---------- PART 01 ----------
  pres.addSection({ title: "背景与问题" });
  s = pres.addSlide({ masterName: "分隔页", sectionTitle: "背景与问题" });
  s.addText("PART 01", { placeholder: "part" });
  s.addText("背景与问题", { placeholder: "title" });
  s.addText("Background & Problem", { placeholder: "body" });

  // ---------- P2 PEST ----------
  s = pres.addSlide({ masterName: "内容页", sectionTitle: "背景与问题" });
  s.addText("量贩零食正在爆发，年轻人“尝鲜大于品牌”", { placeholder: "title" });
  s.addText("理论：PEST", { placeholder: "theory" });
  s.addText("来源：鸣鸣很忙 2025 年度业绩；万辰集团招股书；GB 28050-2025；同龄人访谈，n=15", { placeholder: "source" });
  const pest = [
    { ic: FaBalanceScale, h: "政策与经济", big: "糖必标", lab: "营养标签新国标，2027.3 起强制标示糖", b: "平价消费成主流，量贩靠低价、散称、多 SKU 吸引年轻人" },
    { ic: FaStore, h: "行业", big: "75.5%", lab: "头部两家量贩品牌合计份额（2025）", b: "龙头门店超 2 万家，GMV 同比 +68.5%；品牌要触达年轻人，必须进量贩店" },
    { ic: FaUsers, h: "社会", big: "9 / 15", lab: "同龄人吃完会查配料或含糖量", b: "尝鲜大于品牌，重视情绪价值，也越来越看重健康" },
    { ic: FaRobot, h: "技术", big: "成本↓", lab: "AI 让内容制作和小范围测试更便宜", b: "品牌可以低成本、小步快跑地反复试错" },
  ];
  for (let i = 0; i < 4; i++) {
    const p = pest[i], x = 0.6 + (i % 2) * 6.2, y = 1.4 + Math.floor(i / 2) * 2.4, w = 5.93, h = 2.2;
    card(s, x, y, w, h, HEX.lt1, `pest-${i + 1}`);
    await dot(s, p.ic, x + 0.3, y + 0.3, 0.55, HEX.dk2, HEX.lt2);
    T(s, p.h, { x: x + 1.0, y: y + 0.3, w: 2.2, h: 0.55, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
    T(s, p.big, { x: x + 0.3, y: y + 1.0, w: 2.3, h: 0.7, fontSize: 30, bold: true, color: C.accent1, valign: "middle" });
    T(s, p.lab, { x: x + 0.3, y: y + 1.7, w: 2.5, h: 0.4, fontSize: 10, color: C.accent3, valign: "top" });
    T(s, p.b, { x: x + 2.9, y: y + 0.95, w: 2.75, h: 1.1, fontSize: 14, color: C.text1, valign: "middle" });
  }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 6.2, w: 12.13, h: 0.6, rectRadius: 0.08, fill: { color: C.text2 }, objectName: "conclusion" });
  T(s, "渠道变了、人群变了、技术变了——枣类零食的卖法还没变", { x: 0.85, y: 6.2, w: 11.6, h: 0.6, fontSize: 16, bold: true, color: C.background2, valign: "middle" });
  s.addNotes("【A｜约 25 秒】先看两个事实：行业龙头 2025 年底门店已经超过 2 万家，头部两家合计占了七成半的份额——品牌要找年轻人，就得进量贩店。再看人：我们访谈的 15 个同龄人里，9 个吃完会去查配料或含糖量；2027 年起，营养标签新国标还要求把“糖”强制标出来。年轻人尝鲜大于品牌，也越来越在意健康。机会很大，但卖法要变。");

  // ---------- P3 购买路径 ----------
  s = pres.addSlide({ masterName: "内容页", sectionTitle: "背景与问题" });
  s.addText("年轻人四步买零食，美枣王步步掉队", { placeholder: "title" });
  s.addText("理论：购买决策路径", { placeholder: "theory" });
  s.addText("来源：同龄人访谈，n=15（组内经验萃取 7 人 + 身边同龄人 8 人）", { placeholder: "source" });
  const steps = [
    { h: "种草", d: "刷小红书、抖音看到测评", k: "线上几乎没人种草美枣王", ev: ["10/15", "最近一次在小红书被种草"] },
    { h: "尝鲜", d: "去最近的量贩店买一小包，第一口不好吃就不再买", k: "店里找不到；“枣”像养生品，不想试", ev: ["15/15", "第一次都在附近量贩店买"] },
    { h: "验证", d: "吃完查配料表、含糖量", k: "包装没回答健康问题", ev: ["9/15", "吃完会查配料或含糖量"] },
    { h: "复购", d: "认准一个品牌或一家店", k: "没有再买的理由", ev: ["“同一家店一直有货”", "F1，大三"] },
  ];
  const nw = 2.85, ng = 0.243;
  for (let i = 0; i < 4; i++) {
    const st = steps[i], x = 0.6 + i * (nw + ng), cx = x + nw / 2;
    s.addShape(pres.shapes.OVAL, { x: cx - 0.3, y: 1.4, w: 0.6, h: 0.6, fill: { color: C.text2 }, objectName: `step-${i + 1}` });
    T(s, String(i + 1), { x: cx - 0.3, y: 1.4, w: 0.6, h: 0.6, fontSize: 18, bold: true, color: C.background2, align: "center", valign: "middle" });
    if (i < 3) s.addShape(pres.shapes.LINE, { x: cx + 0.4, y: 1.7, w: nw + ng - 0.8, h: 0, line: { color: HEX.dk2, width: 1.5, endArrowType: "triangle" } });
    T(s, st.h, { x, y: 2.1, w: nw, h: 0.4, fontSize: 16, bold: true, color: C.text2, align: "center", valign: "middle" });
    T(s, st.d, { x, y: 2.5, w: nw, h: 0.55, fontSize: 13, color: C.text1, align: "center", valign: "top" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 3.15, w: nw, h: 1.1, rectRadius: 0.08, fill: { color: C.accent5 }, objectName: `stuck-${i + 1}` });
    T(s, [{ text: "美枣王卡在这", options: { fontSize: 10, bold: true, color: C.accent1, breakLine: true } }, { text: st.k, options: { fontSize: 14, bold: true, color: C.text2 } }],
      { x: x + 0.15, y: 3.2, w: nw - 0.3, h: 1.0, align: "center", valign: "middle", paraSpaceAfter: 3 });
    const quote = st.ev[0].startsWith("“");
    card(s, x, 4.45, nw, 1.45, HEX.lt1, `evidence-${i + 1}`);
    T(s, [{ text: st.ev[0], options: { fontSize: quote ? 14 : 26, bold: true, color: quote ? C.text2 : C.accent1, breakLine: true } }, { text: st.ev[1], options: { fontSize: 11, color: C.accent3 } }],
      { x: x + 0.12, y: 4.5, w: nw - 0.24, h: 1.35, align: "center", valign: "middle", paraSpaceAfter: 2 });
  }
  T(s, "同龄人证据", { x: 0.6, y: 6.0, w: 2, h: 0.3, fontSize: 10, color: C.accent3 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 6.3, w: 12.13, h: 0.5, rectRadius: 0.08, fill: { color: C.text2 }, objectName: "path-conclusion" });
  T(s, "种草→尝鲜→验证→复购：问题不只在“第一次买”，而在买完之后没有被记住", { x: 0.85, y: 6.3, w: 11.6, h: 0.5, fontSize: 15, bold: true, color: C.background2, valign: "middle" });
  s.addNotes("【A｜约 30 秒】用自己的经历讲：我上周刷小红书看到一款新零食，第二天就去了学校门口的零食店，先买一小包；吃完第一件事是翻过来看配料表。这就是年轻人买零食的四步：种草、尝鲜、验证、复购。我们访谈的 15 个人里，10 个最近一次是在小红书被种草，9 个吃完会查配料或含糖量。而美枣王在每一步都掉队：线上没人种草，店里不显眼，枣像养生品，包装没回答健康问题，也没有再买的理由。");

  // ---------- P4 核心问题 ----------
  s = pres.addSlide({ masterName: "内容页", sectionTitle: "背景与问题" });
  s.addText("难的不是首购，而是让首购变成复购", { placeholder: "title" });
  s.addText("理论：购买决策路径 × 价值共创", { placeholder: "theory" });
  s.addText("来源：丁其磊访谈；鸣鸣很忙 2025 年度业绩", { placeholder: "source" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.4, w: 12.13, h: 1.55, rectRadius: 0.1, fill: { color: C.text2 }, objectName: "core-question" });
  T(s, "核心问题", { x: 0.95, y: 1.52, w: 3, h: 0.35, fontSize: 12, bold: true, color: C.accent5 });
  T(s, CORE, { x: 0.95, y: 1.85, w: 11.4, h: 0.95, fontSize: 24, bold: true, color: C.background1, valign: "middle" });
  T(s, [{ text: "为什么这样定义　", options: { bold: true, color: C.accent1 } }, { text: "年轻人本来就爱尝鲜，单纯提高首购意义不大——要在第一次体验里，就设计好让他们再买的理由。", options: { color: C.text1 } }],
    { x: 0.6, y: 3.1, w: 12.13, h: 0.45, fontSize: 15, valign: "middle" });
  // 本质一：三段断开
  card(s, 0.6, 3.7, 5.93, 2.2, HEX.lt1, "essence-1");
  T(s, "本质一｜三段路径是断开的", { x: 0.9, y: 3.85, w: 5.4, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
  ["线上种草", "线下货架", "产品体验"].forEach((t, i) => {
    const x = 0.9 + i * 1.85;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.4, w: 1.5, h: 0.5, rectRadius: 0.08, fill: { color: C.accent2 } });
    T(s, t, { x, y: 4.4, w: 1.5, h: 0.5, fontSize: 13, bold: true, color: C.text2, align: "center", valign: "middle" });
    if (i < 2) T(s, "✕", { x: x + 1.5, y: 4.4, w: 0.35, h: 0.5, fontSize: 16, bold: true, color: C.accent1, align: "center", valign: "middle" });
  });
  T(s, "被种草了找不到店，买到了不知道健不健康", { x: 0.9, y: 5.1, w: 5.4, h: 0.6, fontSize: 14, color: C.text1, valign: "middle" });
  // 本质二：共创缺位
  card(s, 6.8, 3.7, 5.93, 2.2, HEX.lt1, "essence-2");
  T(s, "本质二｜价值共创缺位", { x: 7.1, y: 3.85, w: 5.4, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
  const co = [["品牌", FaIndustry], ["门店", FaStore], ["内容创作者", FaPenNib]];
  for (let i = 0; i < 3; i++) {
    const x = 7.1 + i * 1.8;
    await dot(s, co[i][1], x + 0.3, 4.35, 0.6, HEX.accent2, HEX.dk2);
    T(s, co[i][0], { x: x - 0.1, y: 4.95, w: 1.4, h: 0.3, fontSize: 12, bold: true, color: C.text2, align: "center" });
  }
  T(s, "各做各的，没有一起创造“第一次尝鲜”的体验", { x: 7.1, y: 5.3, w: 5.4, h: 0.45, fontSize: 14, color: C.text1, valign: "middle" });
  // 证据条
  await dot(s, FaQuoteLeft, 0.6, 6.1, 0.5, HEX.accent5, HEX.accent1);
  T(s, [{ text: "买零食“更多的是”尝鲜，口感不差“就会继续再买”", options: { bold: true, color: C.text2, breakLine: true } }, { text: "企业导师 丁其磊老师；渠道是“买手”，会到工厂来选货", options: { fontSize: 10, color: C.accent3 } }],
    { x: 1.2, y: 6.05, w: 5.6, h: 0.65, fontSize: 13, valign: "middle" });
  T(s, "76%", { x: 7.0, y: 6.05, w: 1.15, h: 0.65, fontSize: 28, bold: true, color: C.accent1, valign: "middle" });
  T(s, [{ text: "行业龙头会员复购率（2025）", options: { bold: true, color: C.text2, breakLine: true } }, { text: "渠道已经留得住人，品牌要让自己被记住", options: { fontSize: 10, color: C.accent3 } }],
    { x: 8.2, y: 6.05, w: 4.5, h: 0.65, fontSize: 13, valign: "middle" });
  s.addNotes("【A｜约 35 秒，全篇转折点，语速放慢】所以我们把问题重新定义为：如何让 18–25 岁年轻人在量贩零食店第一次尝鲜美枣王，并通过首购体验带动复购？年轻人本来就爱尝鲜，只追求第一次买意义不大。本质有两条：第一，线上种草、线下货架、产品体验三段是断开的，被种草了找不到店，买到了不知道健不健康；第二，品牌、门店、内容创作者各做各的，没有一起创造第一次尝鲜的体验。丁其磊老师说，零食消费更多是尝鲜，尝了不错才会继续买。");

  // ---------- P5 调研 ----------
  s = pres.addSlide({ masterName: "内容页", sectionTitle: "背景与问题" });
  s.addText("我们的调研：4 种方法、1 条证据链", { placeholder: "title" });
  s.addText("方法：定性优先 · 交叉验证", { placeholder: "theory" });
  s.addText("来源：同龄人访谈，n=15；丁其磊访谈逐字稿；小红书笔记 AI 分析（人工复核）；门店走访", { placeholder: "source" });
  const ms = [
    { ic: FaComments, h: "同龄人访谈", d: "组内经验萃取 7 人 + 身边同龄人 8 人，同样 4 个问题", u: "画出四步购买路径" },
    { ic: FaHashtag, h: "小红书内容分析", d: "AI 归纳“量贩零食”“枣零食”笔记的热门卖点和评论关注点，人工复核", u: "定种草卖点（方案①）" },
    { ic: FaUserTie, h: "专家访谈", d: "企业导师丁其磊老师访谈，整理成逐字稿", u: "判断渠道规则和 AI 边界（方案②④）" },
    { ic: FaWalking, h: "门店走访", d: "走访学校周边量贩零食店，记录新品陈列位置", u: "设计货架方案（方案②）" },
  ];
  const mw = 2.85, mg = 0.243;
  for (let i = 0; i < 4; i++) {
    const m = ms[i], x = 0.6 + i * (mw + mg);
    card(s, x, 1.4, mw, 3.85, HEX.lt1, `method-${i + 1}`);
    await dot(s, m.ic, x + 0.3, 1.65, 0.65, HEX.dk2, HEX.lt2);
    T(s, `0${i + 1}`, { x: x + mw - 0.9, y: 1.65, w: 0.6, h: 0.65, fontSize: 20, bold: true, color: C.accent4, align: "right", valign: "middle" });
    T(s, m.h, { x: x + 0.3, y: 2.45, w: mw - 0.6, h: 0.45, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
    T(s, m.d, { x: x + 0.3, y: 2.95, w: mw - 0.6, h: 1.15, fontSize: 14, color: C.text1, valign: "top" });
    s.addShape(pres.shapes.LINE, { x: x + 0.3, y: 4.2, w: mw - 0.6, h: 0, line: { color: HEX.accent4, width: 0.75 } });
    T(s, [{ text: "用来 → ", options: { bold: true, color: C.accent1 } }, { text: m.u, options: { bold: true, color: C.text2 } }],
      { x: x + 0.3, y: 4.3, w: mw - 0.6, h: 0.8, fontSize: 13, valign: "middle" });
  }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.45, w: 12.13, h: 0.6, rectRadius: 0.08, fill: { color: C.text2 }, objectName: "workload" });
  T(s, "15 位同龄人　｜　AI 分析小红书相关笔记　｜　1 次企业导师访谈　｜　走访学校周边门店", { x: 0.6, y: 5.45, w: 12.13, h: 0.6, fontSize: 15, bold: true, color: C.background2, align: "center", valign: "middle" });
  T(s, [{ text: "问卷怎么用　", options: { bold: true, color: C.accent1 } }, { text: "问卷只适合问题已经明确、要验证比例的时候用，所以放在第二步开始前，验证卖点和复购钩子（约 200 份，设计见执行部分）", options: { color: C.accent6 } }],
    { x: 0.6, y: 6.2, w: 12.13, h: 0.5, fontSize: 12, valign: "middle" });
  s.addNotes("【A｜约 20 秒，带过】这些判断来自四种方法：访谈了 15 位同龄人，用 AI 分析小红书笔记再人工复核，做了 1 次企业导师访谈，走访了学校周边门店。问卷留到第二步开始前做验证。问题找到了，下面看我们设计的闭环。");

  // ---------- PART 02 / 03 分隔页（放到各部分开头） ----------
  pres.addSection({ title: "分隔页（插到方案、执行部分前）" });
  s = pres.addSlide({ masterName: "分隔页", sectionTitle: "分隔页（插到方案、执行部分前）" });
  s.addText("PART 02", { placeholder: "part" });
  s.addText("解决方案", { placeholder: "title" });
  s.addText("Solution", { placeholder: "body" });
  s.addNotes("插在方案总览（原 P6）之前。");
  s = pres.addSlide({ masterName: "分隔页", sectionTitle: "分隔页（插到方案、执行部分前）" });
  s.addText("PART 03", { placeholder: "part" });
  s.addText("执行与收尾", { placeholder: "title" });
  s.addText("Execution", { placeholder: "body" });
  s.addNotes("插在“分三步走”（原 P11）之前。");

  // ---------- P14 总结 ----------
  pres.addSection({ title: "总结与封底" });
  s = pres.addSlide({ masterName: "内容页", sectionTitle: "总结与封底" });
  s.addText("总结：从一次尝鲜到长期复购", { placeholder: "title" });
  s.addText("回扣核心问题", { placeholder: "theory" });
  s.addText("来源：成本为第一步测算（假设值），见执行部分", { placeholder: "source" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.4, w: 12.13, h: 1.05, rectRadius: 0.1, fill: { color: C.text2 }, objectName: "core-question" });
  T(s, CORE, { x: 0.95, y: 1.4, w: 11.4, h: 1.05, fontSize: 20, bold: true, color: C.background1, valign: "middle" });
  const sum = [
    { ic: FaExclamationTriangle, h: "问题", t: "首购没有变成复购" },
    { ic: FaSyncAlt, h: "方案", t: "种草—尝鲜—验证—复购闭环" },
    { ic: FaFlag, h: "落地", t: "校园起步，4 周、约 2600 元，跑通即可复制" },
  ];
  for (let i = 0; i < 3; i++) {
    const y = 2.7 + i * 1.05;
    card(s, 0.6, y, 12.13, 0.85, HEX.lt1, `summary-${i + 1}`);
    await dot(s, sum[i].ic, 0.85, y + 0.14, 0.57, HEX.dk2, HEX.lt2);
    T(s, sum[i].h, { x: 1.65, y, w: 1.2, h: 0.85, fontSize: 18, bold: true, color: C.accent1, valign: "middle" });
    T(s, sum[i].t, { x: 2.85, y, w: 9.6, h: 0.85, fontSize: 22, bold: true, color: C.text2, valign: "middle" });
  }
  T(s, [{ text: "展望　", options: { bold: true, color: C.accent1 } }, { text: "从一所学校跑通，变成美枣王在全国量贩渠道的年轻化增长模式", options: { color: C.text2 } }],
    { x: 0.6, y: 5.95, w: 12.13, h: 0.6, fontSize: 18, valign: "middle" });
  s.addNotes("【C｜约 20 秒】回到我们的核心问题。问题是首购没有变成复购；方案是一条种草、尝鲜、验证、复购的闭环；落地从一所学校开始，4 周、约 2600 元，跑通就能复制。我们希望从一所学校跑通，变成美枣王在全国量贩渠道的年轻化增长模式。");

  // ---------- P15 封底 ----------
  s = pres.addSlide({ masterName: "封面", sectionTitle: "总结与封底" });
  s.addText("谢谢！欢迎提问", { placeholder: "title" });
  s.addText("让年轻人第一口就爱上，第二包自然会来。", { placeholder: "body" });
  s.addText([{ text: "第【】组　组长：【】", options: { breakLine: true } }, { text: "组员：【】【】【】【】【】【】" }], { placeholder: "members" });
  s.addNotes("【C｜约 5 秒】让年轻人第一口就爱上，第二包自然会来。谢谢老师！");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
