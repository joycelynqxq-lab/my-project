// P5 问题本质（详细版），单独生成
const { newPres, THEME, HEX, icon } = require("./master");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
const { FaMobileAlt, FaMagic, FaIndustry, FaStore, FaUserFriends } = require("react-icons/fa");
const OUT = process.argv[2] || "P5.pptx";

(async () => {
  const { pres, C } = newPres();
  const sh = () => ({ type: "outer", color: "000000", opacity: 0.12, blur: 6, offset: 2, angle: 90 });
  const s = pres.addSlide({ masterName: "内容页" });
  s.addText("本质：第一次尝鲜体验没有被共同创造", { placeholder: "title" });
  s.addText("价值网络＋服务主导逻辑", { placeholder: "theory" });
  s.addText("企业导师访谈逐字稿；组内分析", { placeholder: "source" });

  const T = (txt, o) => s.addText(txt, { margin: 0, isTextBox: true, ...o });
  const arrow = (x1, y1, x2, y2, dash) => s.addShape(pres.shapes.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1) || 0.001, h: Math.abs(y2 - y1) || 0.001,
    flipH: x2 < x1, flipV: y2 < y1,
    line: { color: dash ? HEX.accent4 : HEX.dk2, width: dash ? 1.25 : 2, dashType: dash ? "dash" : "solid", endArrowType: "triangle" },
  });
  const lab = (t, x, y, w, color) => T(t, { x, y, w, h: 0.28, fontSize: 10, color: color || C.accent4, align: "center", valign: "middle" });

  // ---------- 左：价值网络图 ----------
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.4, w: 7.75, h: 4.55, rectRadius: 0.12, fill: { color: C.background2 }, objectName: "network-bg" });
  T("美枣王价值网络", { x: 0.85, y: 1.55, w: 2.5, h: 0.35, fontSize: 13, bold: true, color: C.accent4 });

  const cy = 3.3;
  const node = async (ic, title, sub, x, w, dark, name) => {
    const fill = dark ? HEX.dk2 : HEX.lt1, fc = dark ? HEX.lt1 : HEX.dk2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy - 0.6, w, h: 1.2, rectRadius: 0.1, fill: { color: fill }, shadow: sh(), objectName: name });
    s.addImage({ data: await icon(ic, dark ? HEX.accent2 : HEX.accent1), x: x + w / 2 - 0.13, y: cy - 0.5, w: 0.26, h: 0.26 });
    T([
      { text: title, options: { bold: true, fontSize: 14, color: fc, breakLine: true } },
      { text: sub, options: { fontSize: 10, color: dark ? HEX.lt1 : HEX.accent4 } },
    ], { x: x + 0.05, y: cy - 0.22, w: w - 0.1, h: 0.75, align: "center", valign: "top", paraSpaceAfter: 2 });
  };
  await node(FaIndustry, "美枣王", "提供产品\n终端控制力弱", 0.8, 1.45, false, "n-brand");
  await node(FaStore, "零食很忙", "买手：选品·陈列\n掌握 POS 数据", 2.75, 1.6, false, "n-store");
  // 货架决策节点
  s.addShape(pres.shapes.OVAL, { x: 4.75, y: cy - 0.78, w: 1.56, h: 1.56, fill: { color: C.background1 }, line: { color: C.accent1, width: 3.5 }, objectName: "n-shelf" });
  T([
    { text: "货架", options: { bold: true, fontSize: 15, color: C.accent1, breakLine: true } },
    { text: "决策节点", options: { bold: true, fontSize: 15, color: C.accent1, breakLine: true } },
    { text: "只有几秒", options: { fontSize: 10, color: C.accent4 } },
  ], { x: 4.75, y: cy - 0.78, w: 1.56, h: 1.56, align: "center", valign: "middle" });
  await node(FaUserFriends, "年轻人", "18-25 岁\n逛店、爱尝鲜", 6.85, 1.35, true, "n-youth");
  arrow(2.25, cy, 2.75, cy); lab("供货", 2.2, cy - 0.32, 0.6);
  arrow(4.35, cy, 4.75, cy); lab("陈列", 4.25, cy - 0.32, 0.6);
  arrow(6.31, cy, 6.85, cy); lab("拿不拿", 6.31, cy - 0.32, 0.54, C.accent1);

  // 上：内容平台
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.55, y: 1.55, w: 3.6, h: 0.72, rectRadius: 0.1, fill: { color: C.background1 }, line: { color: C.accent4, width: 1, dashType: "dash" }, objectName: "n-content" });
  s.addImage({ data: await icon(FaMobileAlt, HEX.accent4), x: 4.7, y: 1.76, w: 0.3, h: 0.3 });
  T([{ text: "内容平台 / 校园 KOC", options: { bold: true, fontSize: 13, color: C.text1, breakLine: true } }, { text: "小红书、抖音“量贩必买榜”", options: { fontSize: 10, color: C.accent4 } }],
    { x: 5.1, y: 1.55, w: 3.0, h: 0.72, valign: "middle" });
  arrow(7.52, 2.27, 7.52, cy - 0.6, true); lab("进店前种草", 6.6, 2.33, 0.85);

  // 下：AI 工具商
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 4.65, w: 3.55, h: 0.75, rectRadius: 0.1, fill: { color: C.background1 }, line: { color: C.accent4, width: 1, dashType: "dash" }, objectName: "n-ai" });
  s.addImage({ data: await icon(FaMagic, HEX.accent4), x: 0.95, y: 4.87, w: 0.3, h: 0.3 });
  T([{ text: "AI 工具商", options: { bold: true, fontSize: 13, color: C.text1, breakLine: true } }, { text: "包装/文案快速生成，新品更多、货架更挤", options: { fontSize: 10, color: C.accent4 } }],
    { x: 1.35, y: 4.65, w: 2.95, h: 0.75, valign: "middle" });
  arrow(1.5, 4.65, 1.5, cy + 0.6, true); arrow(3.55, 4.65, 3.55, cy + 0.6, true);

  // 红圈下方：缺席说明
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.6, y: 4.3, w: 3.55, h: 1.38, rectRadius: 0.1, fill: { color: C.background1 }, line: { color: C.accent1, width: 1.25 }, objectName: "absence" });
  T([
    { text: "品牌在这里缺席", options: { bold: true, fontSize: 13, color: C.accent1, breakLine: true } },
    { text: "年轻人卡在这三处：", options: { fontSize: 11, color: C.text1, breakLine: true } },
    { text: "看不懂口味", options: { bullet: true, fontSize: 11, color: C.text1, breakLine: true } },
    { text: "一包太大、怕踩雷", options: { bullet: true, fontSize: 11, color: C.text1, breakLine: true } },
    { text: "“枣 = 养生、长辈”", options: { bullet: true, fontSize: 11, color: C.text1 } },
  ], { x: 4.8, y: 4.38, w: 3.25, h: 1.35, valign: "top", paraSpaceAfter: 1 });
  arrow(5.53, 4.3, 5.53, cy + 0.78);

  // 图例
  s.addShape(pres.shapes.LINE, { x: 0.85, y: 5.72, w: 0.4, h: 0, line: { color: HEX.dk2, width: 2 } });
  T("产品 / 交易流", { x: 1.32, y: 5.6, w: 1.3, h: 0.25, fontSize: 10, color: C.accent4 });
  s.addShape(pres.shapes.LINE, { x: 2.65, y: 5.72, w: 0.4, h: 0, line: { color: HEX.accent4, width: 1.25, dashType: "dash" } });
  T("信息 / 内容流", { x: 3.12, y: 5.6, w: 1.3, h: 0.25, fontSize: 10, color: C.accent4 });

  // ---------- 右：两条本质 ----------
  const card = (y, h, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.6, y, w: 4.13, h, rectRadius: 0.12, fill: { color: C.background1 }, line: { color: C.accent5, width: 1.5 }, shadow: sh(), objectName: name });
  card(1.4, 2.2, "essence-1");
  T([
    { text: "本质一｜价值网络", options: { bold: true, fontSize: 15, color: C.accent1, breakLine: true } },
    { text: "量贩店是“买手”：选什么、摆哪里、卖多少钱，由渠道决定", options: { bullet: true, fontSize: 13, color: C.text1, breakLine: true } },
    { text: "品牌控制力弱，在货架前“说不上话”", options: { bullet: true, fontSize: 13, color: C.text1, breakLine: true } },
    { text: "首购不是品牌一方能解决的，必须和渠道共创", options: { bold: true, fontSize: 14, color: C.text2 } },
  ], { x: 8.82, y: 1.5, w: 3.7, h: 2.0, valign: "top", paraSpaceAfter: 8 });

  card(3.75, 2.2, "essence-2");
  T("本质二｜服务主导逻辑", { x: 8.82, y: 3.85, w: 3.7, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
  T("价值在“吃第一口”时才产生（使用价值），不在出厂时", { x: 8.82, y: 4.22, w: 3.7, h: 0.5, fontSize: 13, color: C.text1, valign: "top" });
  // 天平式对比：风险 > 价值
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.82, y: 4.8, w: 1.55, h: 0.55, rectRadius: 0.08, fill: { color: C.accent1 } });
  T("感知风险", { x: 8.82, y: 4.8, w: 1.55, h: 0.55, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle" });
  T(">", { x: 10.4, y: 4.8, w: 0.4, h: 0.55, fontSize: 22, bold: true, color: C.accent1, align: "center", valign: "middle" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 10.83, y: 4.88, w: 1.4, h: 0.4, rectRadius: 0.08, fill: { color: C.accent5 } });
  T("预期价值", { x: 10.83, y: 4.88, w: 1.4, h: 0.4, fontSize: 12, bold: true, color: C.accent1, align: "center", valign: "middle" });
  T("问题不在产品好不好，在“第一次试”的门槛太高", { x: 8.82, y: 5.45, w: 3.7, h: 0.4, fontSize: 12, bold: true, color: C.text2, valign: "middle" });

  // ---------- 底：一句话定义 ----------
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 6.1, w: 12.13, h: 0.68, rectRadius: 0.1, fill: { color: C.text2 }, objectName: "definition" });
  T([
    { text: "一句话定义  ", options: { bold: true, color: C.accent2 } },
    { text: "首购率低，是美枣王、零食很忙、年轻人之间没有共同创造“第一次尝鲜体验”的问题", options: { color: C.background1 } },
  ], { x: 0.85, y: 6.1, w: 11.7, h: 0.68, fontSize: 16, valign: "middle" });

  s.addNotes("【A｜约 45 秒】那问题的本质是什么？我们画了美枣王的价值网络：美枣王把货供给零食很忙，零食很忙作为买手决定陈列，年轻人在货架前决定拿不拿；进店前，小红书和抖音的必买榜已经在种草，AI 工具又让新品越来越多、货架越来越挤。关键在这个红圈——货架前只有几秒，年轻人卡在看不懂口味、一包太大、觉得枣是长辈吃的，而品牌在这一刻是缺席的。第一，从价值网络看，量贩店是买手，品牌在终端说不上话，首购必须和渠道一起解决。第二，从服务主导逻辑看，价值在吃第一口时才产生，首购低是因为感知风险大于预期价值，不是产品不好。所以我们的定义是：首购率低，是三方没有共同创造“第一次尝鲜体验”。这个判断对不对？我们去门店和同龄人那里验证了一下。");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
