// 目录页（插在 P1 封面和 P2 背景之间），单独生成
const { newPres, THEME, HEX } = require("./master");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
const OUT = process.argv[2] || "toc.pptx";

(async () => {
  const { pres, C } = newPres();
  const s = pres.addSlide();
  s.background = { color: HEX.lt1 };

  // 左侧深色块
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 4.4, h: 7.5, fill: { color: C.text2 }, objectName: "toc-panel" });
  s.addText("目录", { x: 0.7, y: 2.45, w: 3.3, h: 1.0, fontSize: 48, bold: true, color: C.background1, margin: 0, isTextBox: true, objectName: "toc-title" });
  s.addText("CONTENTS", { x: 0.7, y: 3.45, w: 3.3, h: 0.45, fontSize: 18, bold: true, color: C.accent2, charSpacing: 4, margin: 0, isTextBox: true });
  s.addText("提出问题 → 分析问题 → 解决问题", { x: 0.7, y: 4.2, w: 3.4, h: 0.4, fontSize: 13, color: C.background1, margin: 0, isTextBox: true });

  // 右侧四个章节
  const parts = [
    { n: "01", h: "背景与问题", d: "时代背景、企业家原问题、现有方案、问题本质", p: "P3-P6" },
    { n: "02", h: "调研设计与发现", d: "价值链环节、利益相关者、面访 / 观察 / 人机对比", p: "P7-P9" },
    { n: "03", h: "方案：3 秒首购计划", d: "货架尝鲜体验共创：总览 + 四个分方案", p: "P10-P14" },
    { n: "04", h: "实施与总结", d: "分阶段路线、资源与 KPI、回扣问题", p: "P15-P17" },
  ];
  const x0 = 5.1, rowH = 1.2, y0 = 1.15;
  parts.forEach((t, i) => {
    const y = y0 + i * (rowH + 0.2);
    const hi = i === 2; // 方案是重点
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x0, y, w: 7.63, h: rowH, rectRadius: 0.12, fill: { color: hi ? HEX.accent1 : HEX.lt2 }, objectName: `toc-row-${i + 1}` });
    s.addText(t.n, { x: x0 + 0.3, y, w: 1.1, h: rowH, fontSize: 40, bold: true, color: hi ? C.accent2 : C.accent1, valign: "middle", margin: 0, isTextBox: true });
    s.addText([
      { text: t.h, options: { fontSize: 20, bold: true, color: hi ? C.background1 : C.text2, breakLine: true } },
      { text: t.d, options: { fontSize: 13, color: hi ? C.background1 : C.accent4 } },
    ], { x: x0 + 1.5, y, w: 4.75, h: rowH, valign: "middle", margin: 0, paraSpaceAfter: 2, isTextBox: true });
    s.addText(t.p, { x: x0 + 6.3, y, w: 1.1, h: rowH, fontSize: 13, bold: true, color: hi ? C.background1 : C.accent4, align: "right", valign: "middle", margin: 0, isTextBox: true });
  });

  s.addNotes("【A｜约 5 秒，带过即可】我们按“提出问题、分析问题、解决问题”三步来讲，重点是第三部分的“3秒首购计划”。");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
