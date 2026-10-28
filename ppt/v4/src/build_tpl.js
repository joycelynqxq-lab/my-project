// 发给制作2/3/4的母版文件：三个版式 + 一页规范 + 一页空白示例
const { newPres, THEME, HEX } = require("./master3");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
const OUT = process.argv[2] || "template.pptx";

(async () => {
  const { pres, C } = newPres();
  const T = (s, txt, o) => s.addText(txt, { margin: 0, isTextBox: true, ...o });

  let s = pres.addSlide({ masterName: "内容页" });
  s.addText("母版规范（正式稿删掉本页）", { placeholder: "title" });
  s.addText("理论：本页用到的理论", { placeholder: "theory" });
  s.addText("来源：联商网，2026.7；同龄人访谈，n=15；丁其磊访谈", { placeholder: "source" });
  const sw = [["7A1426", "枣红 主色", "标题、深色块"], ["B3203A", "亮枣红", "大数字、角标、强调"], ["FAF6F0", "米白", "页面背景"], ["FFFFFF", "白", "卡片底"], ["3A3536", "深灰", "正文"], ["8A8083", "中灰", "注释、来源"]];
  sw.forEach((c, i) => {
    const x = 0.6 + (i % 3) * 2.05, y = 1.45 + Math.floor(i / 3) * 1.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 0.75, h: 0.75, rectRadius: 0.08, fill: { color: c[0] }, line: { color: HEX.accent4, width: 0.75 } });
    T(s, [{ text: c[1], options: { bold: true, breakLine: true } }, { text: "#" + c[0], options: { fontSize: 10, color: C.accent3, breakLine: true } }, { text: c[2], options: { fontSize: 10 } }],
      { x: x + 0.85, y, w: 1.15, h: 0.9, fontSize: 12, color: C.text1, valign: "top" });
  });
  T(s, [
    { text: "字体：全套微软雅黑（一种黑体），不要混用宋体、楷体", options: { bullet: true, breakLine: true } },
    { text: "字号：标题 30 加粗；正文 14–16；注释 10–12", options: { bullet: true, breakLine: true } },
    { text: "标题写成一句结论；正文不超过 80 字，能画图就不写段落", options: { bullet: true, breakLine: true } },
    { text: "新建页：开始 → 新建幻灯片 → 选「内容页」；右上「理论」、右下「来源」是版式自带的框，每页都填", options: { bullet: true, breakLine: true } },
    { text: "证据放在它支撑的观点旁边，不做单独的数据页", options: { bullet: true, breakLine: true } },
    { text: "页码自动生成；卡片用白底圆角 + 浅灰细边，深色结论条用枣红", options: { bullet: true } },
  ], { x: 6.95, y: 1.45, w: 5.78, h: 5.2, fontSize: 14, color: C.text1, valign: "top", paraSpaceAfter: 8 });

  s = pres.addSlide({ masterName: "内容页" });
  s.addText("示例：标题写结论，比如“附近门店接住流量”", { placeholder: "title" });
  s.addText("理论：价值共创", { placeholder: "theory" });
  s.addText("来源：门店走访；丁其磊访谈", { placeholder: "source" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.4, w: 12.13, h: 5.3, rectRadius: 0.1, fill: { color: C.background1 }, line: { color: HEX.accent4, width: 0.75 } });
  T(s, "内容区：x 0.6～12.73 英寸，y 1.4～6.8 英寸；卡片间距 0.25 英寸", { x: 0.9, y: 3.8, w: 11.5, h: 0.5, fontSize: 16, color: C.accent3, align: "center" });

  s = pres.addSlide({ masterName: "分隔页" });
  s.addText("PART 0X", { placeholder: "part" });
  s.addText("分隔页版式", { placeholder: "title" });
  s.addText("English subtitle", { placeholder: "body" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
