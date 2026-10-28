// 第 4 版统一母版：枣红主色 + 米白 + 深灰，全套微软雅黑；理论角标右上，来源右下
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

const FONT = "Microsoft YaHei";
const BG = path.join(__dirname, "bg.jpg");
const THEME = {
  name: "MeizaoWang v4",
  headFontFace: FONT,
  bodyFontFace: FONT,
  colors: {
    dk1: "3A3536", // 深灰 正文
    lt1: "FFFFFF",
    dk2: "7A1426", // 枣红 主色
    lt2: "FAF6F0", // 米白 背景
    accent1: "B3203A", // 亮枣红 强调
    accent2: "F2E6DC", // 米白深一档 卡片底
    accent3: "8A8083", // 中灰 注释
    accent4: "D9CBC4", // 浅灰 分隔线
    accent5: "F6DADE", // 浅红 卡点底
    accent6: "5A5255", // 深灰二
    hlink: "B3203A",
    folHlink: "7A1426",
  },
};
const HEX = THEME.colors;

function newPres() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "一口尝鲜，一路复购";
  pres.subject = "数智商业模式创新 小组汇报";
  const C = pres.SchemeColor;
  const footer = { text: { text: "美枣王 · 一口尝鲜，一路复购", options: { x: 1.05, y: 7.0, w: 3.5, h: 0.3, fontSize: 10, color: C.accent3, margin: 0, valign: "middle" } } };
  const num = { x: 0.6, y: 7.0, w: 0.45, h: 0.3, fontSize: 10, color: C.accent3, fontFace: FONT };

  pres.defineSlideMaster({
    title: "封面",
    background: { path: BG },
    objects: [
      { rect: { x: 0, y: 0, w: 7.6, h: 7.5, fill: { color: "1E1416", transparency: 35 } } },
      { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 1.5, w: 6.6, h: 1.9, fontSize: 40, bold: true, color: C.background1, align: "left", valign: "bottom", margin: 0 }, text: "主标题" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.7, y: 3.6, w: 6.4, h: 1.1, fontSize: 16, color: C.background2, align: "left", valign: "top", margin: 0 }, text: "副标题：核心问题" } },
      { placeholder: { options: { name: "members", type: "body", x: 0.7, y: 5.55, w: 6.4, h: 1.0, fontSize: 13, color: C.background2, align: "left", valign: "top", margin: 0 }, text: "课程 / 组长 / 组员 / 日期" } },
    ],
  });

  pres.defineSlideMaster({
    title: "分隔页",
    background: { path: BG },
    objects: [
      { rect: { x: 0, y: 0, w: 7.6, h: 7.5, fill: { color: "1E1416", transparency: 35 } } },
      { rect: { x: 0, y: 2.7, w: 7.6, h: 2.6, fill: { color: C.text2, transparency: 12 } } },
      { placeholder: { options: { name: "part", type: "body", x: 0.7, y: 1.2, w: 5, h: 1.3, fontSize: 54, bold: true, color: C.background2, align: "left", valign: "bottom", margin: 0 }, text: "PART 01" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 2.95, w: 6.6, h: 1.3, fontSize: 48, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "章节标题" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.7, y: 4.25, w: 6.6, h: 0.8, fontSize: 16, color: C.background2, align: "left", valign: "top", margin: 0 }, text: "一句话说明" } },
    ],
  });

  pres.defineSlideMaster({
    title: "内容页",
    background: { color: C.background2 },
    slideNumber: num,
    objects: [
      footer,
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 9.3, h: 0.75, fontSize: 30, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "页标题：写成一句结论" } },
      { placeholder: { options: { name: "theory", type: "body", x: 10.0, y: 0.5, w: 2.73, h: 0.5, fontSize: 12, bold: true, color: C.accent1, align: "right", valign: "middle", margin: 0 }, text: "理论：如 PEST" } },
      { placeholder: { options: { name: "source", type: "body", x: 5.0, y: 6.98, w: 7.73, h: 0.34, fontSize: 10, color: C.accent3, align: "right", valign: "middle", margin: 0 }, text: "来源：如 联商网，2026.7；同龄人访谈，n=15" } },
    ],
  });
  return { pres, C };
}

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

module.exports = { newPres, THEME, HEX, FONT, icon };
