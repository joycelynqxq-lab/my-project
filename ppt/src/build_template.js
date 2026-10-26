const { newPres, THEME, HEX } = require("./master");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
const OUT = process.argv[2] || "template.pptx";

(async () => {
  const { pres, C } = newPres();

  // 1. 使用说明
  let s = pres.addSlide({ masterName: "内容页" });
  s.addText("母版使用说明（正式稿删掉本页）", { placeholder: "title" });
  s.addText("理论名写这里", { placeholder: "theory" });
  s.addText("证据来源写这里，如：面访 n=12", { placeholder: "source" });
  const sw = [["7A1426", "枣红 主色", "标题、深色块"], ["B3203A", "亮枣红", "理论角标、重点数字"], ["FFB61E", "活力黄", "深色块上的强调字、箭头"],
    ["2F7D68", "枣叶绿", "“好/能”"], ["F5EEF0", "浅枣粉", "卡片底色"], ["6E6266", "中灰", "注释、来源"]];
  sw.forEach((c, i) => {
    const x = 0.6 + (i % 3) * 2.1, y = 1.5 + Math.floor(i / 3) * 1.3;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 0.8, h: 0.8, rectRadius: 0.08, fill: { color: c[0] }, line: { color: HEX.accent4, width: 0.5 } });
    s.addText([{ text: c[1], options: { bold: true, breakLine: true } }, { text: "#" + c[0], options: { color: C.accent4, fontSize: 11, breakLine: true } }, { text: c[2], options: { fontSize: 11 } }],
      { x: x + 0.9, y, w: 1.15, h: 0.95, fontSize: 12, color: C.text1, valign: "top", margin: 0, isTextBox: true });
  });
  s.addText([
    { text: "字体：微软雅黑（全套统一）", options: { bullet: true, breakLine: true } },
    { text: "字号：页标题 30 加粗｜卡片小标题 16-18 加粗｜正文 14-16｜注释 10-12", options: { bullet: true, breakLine: true } },
    { text: "新建页面：开始 → 新建幻灯片 → 选「内容页」或「方案页」版式", options: { bullet: true, breakLine: true } },
    { text: "右上角「理论」、右下角「来源」是版式自带占位符，每页都填，没用理论就写本页主题", options: { bullet: true, breakLine: true } },
    { text: "没拿到的数据写成【待填：……】并用黄色底高亮，统稿前全部替换，不要编数", options: { bullet: true, breakLine: true } },
    { text: "页码由母版自动生成，合并后会自动连续", options: { bullet: true } },
  ], { x: 7.0, y: 1.5, w: 5.73, h: 5.1, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 8, isTextBox: true });

  // 2. 封面版式示例
  s = pres.addSlide({ masterName: "封面" });
  s.addText("封面 / 封底版式（制作 4 做 P17 用）", { placeholder: "title" });
  s.addText("谢谢老师 + 一句 slogan", { placeholder: "body" });
  s.addText("欢迎提问", { placeholder: "members" });

  // 3. 内容页空白示例
  s = pres.addSlide({ masterName: "内容页" });
  s.addText("内容页示例：标题写一句结论", { placeholder: "title" });
  s.addText("价值链", { placeholder: "theory" });
  s.addText("面访 n=【】；门店观察【】店", { placeholder: "source" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.45, w: 12.13, h: 5.2, rectRadius: 0.12, fill: { color: C.background2 } });
  s.addText("内容区：x 0.6～12.73 英寸，y 1.45～6.75 英寸。卡片用浅枣粉圆角矩形，间距 0.35 英寸", { x: 0.9, y: 3.6, w: 11.5, h: 0.6, fontSize: 16, color: C.accent4, align: "center", margin: 0, isTextBox: true });

  // 4. 方案页示例
  s = pres.addSlide({ masterName: "方案页" });
  s.addText("① 产品尝鲜化：一口就能试", { placeholder: "title" });
  s.addText("模块化", { placeholder: "theory" });
  s.addText([{ text: "具体动作 1", options: { bullet: true, breakLine: true } }, { text: "具体动作 2", options: { bullet: true, breakLine: true } }, { text: "具体动作 3", options: { bullet: true } }], { placeholder: "body" });
  s.addText([{ text: "P7 “一包太大不敢试” 【待填数据】", options: { bullet: true, breakLine: true } }, { text: "丁其磊老师：“安全第一、口感第二”", options: { bullet: true } }], { placeholder: "evidence" });
  s.addText("面访 n=【】；企业导师访谈", { placeholder: "source" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
