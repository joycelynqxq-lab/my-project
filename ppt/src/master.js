// 共享母版：主题色、字体、版式（layouts）。P1-P5 和给组员的模板都从这里生成。
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

const FONT = "Microsoft YaHei";
const THEME = {
  name: "MeizaoWang 3s",
  headFontFace: FONT,
  bodyFontFace: FONT,
  colors: {
    dk1: "2A1A1D", // 正文深色
    lt1: "FFFFFF",
    dk2: "7A1426", // 枣红（主色）
    lt2: "F5EEF0", // 浅枣粉（卡片底）
    accent1: "B3203A", // 亮枣红
    accent2: "FFB61E", // 活力黄（强调/年轻）
    accent3: "2F7D68", // 枣叶绿（“好/能”）
    accent4: "6E6266", // 中灰（来源/注释）
    accent5: "F7D9DE", // 浅红（“差/不能”底色）
    accent6: "DCEFE8", // 浅绿（“好/能”底色）
    hlink: "B3203A",
    folHlink: "7A1426",
  },
};
const HEX = THEME.colors;

function newPres() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "3秒首购：美枣王首购率方案";
  pres.subject = "数智商业创新 小组作业";
  const C = pres.SchemeColor;

  // 封面 / 封底（深色）
  pres.defineSlideMaster({
    title: "封面",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 1.7, w: 8.2, h: 2.0, fontSize: 36, align: "left", bold: true, color: C.background1, valign: "top", margin: 0 }, text: "标题" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.7, y: 3.85, w: 8.2, h: 0.6, fontSize: 18, color: C.accent2, valign: "top", margin: 0 }, text: "副标题" } },
      { placeholder: { options: { name: "members", type: "body", x: 0.7, y: 5.3, w: 8.2, h: 1.3, fontSize: 14, color: C.background1, valign: "top", margin: 0 }, text: "组长 / 组员 / 课程 / 日期" } },
    ],
  });

  // 内容页（浅色）：标题左上，理论角标右上，证据角标右下
  pres.defineSlideMaster({
    title: "内容页",
    background: { color: C.background1 },
    margin: [0.5, 0.6, 0.6, 0.6],
    slideNumber: { x: 0.6, y: 6.98, w: 0.5, h: 0.3, fontSize: 10, color: C.accent4, fontFace: FONT },
    objects: [
      { text: { text: "美枣王 · 3秒首购计划", options: { x: 1.05, y: 6.98, w: 3, h: 0.3, fontSize: 10, color: C.accent4, margin: 0, valign: "middle" } } },
      { rect: { x: 9.75, y: 0.58, w: 0.62, h: 0.34, fill: { color: C.accent1 }, rectRadius: 0.08 } },
      { text: { text: "理论", options: { x: 9.75, y: 0.58, w: 0.62, h: 0.34, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 } } },
      { text: { text: "来源", options: { x: 7.55, y: 6.98, w: 0.5, h: 0.3, fontSize: 10, bold: true, color: C.accent1, margin: 0, valign: "middle" } } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 9.0, h: 0.7, fontSize: 30, align: "left", bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "页面标题：一句话结论" } },
      { placeholder: { options: { name: "theory", type: "body", x: 10.45, y: 0.52, w: 2.3, h: 0.46, fontSize: 12, bold: true, color: C.accent1, valign: "middle", margin: 0 }, text: "如：价值网络" } },
      { placeholder: { options: { name: "source", type: "body", x: 8.05, y: 6.93, w: 4.7, h: 0.4, fontSize: 10, color: C.accent4, valign: "middle", margin: 0 }, text: "如：面访 n=12；门店观察 3 店" } },
    ],
  });

  // 方案页：左“做什么”右“证据”（给制作 3 用）
  pres.defineSlideMaster({
    title: "方案页",
    background: { color: C.background1 },
    slideNumber: { x: 0.6, y: 6.98, w: 0.5, h: 0.3, fontSize: 10, color: C.accent4, fontFace: FONT },
    objects: [
      { text: { text: "美枣王 · 3秒首购计划", options: { x: 1.05, y: 6.98, w: 3, h: 0.3, fontSize: 10, color: C.accent4, margin: 0, valign: "middle" } } },
      { rect: { x: 9.75, y: 0.58, w: 0.62, h: 0.34, fill: { color: C.accent1 }, rectRadius: 0.08 } },
      { text: { text: "理论", options: { x: 9.75, y: 0.58, w: 0.62, h: 0.34, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 } } },
      { text: { text: "来源", options: { x: 7.55, y: 6.98, w: 0.5, h: 0.3, fontSize: 10, bold: true, color: C.accent1, margin: 0, valign: "middle" } } },
      { rect: { x: 0.6, y: 1.4, w: 6.6, h: 5.3, fill: { color: C.background2 } } },
      { rect: { x: 7.5, y: 1.4, w: 5.23, h: 5.3, fill: { color: C.background1 }, line: { color: C.accent1, width: 1.25, dashType: "dash" } } },
      { text: { text: "做什么", options: { x: 0.9, y: 1.6, w: 3, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0 } } },
      { text: { text: "调研证据", options: { x: 7.8, y: 1.6, w: 3, h: 0.4, fontSize: 16, bold: true, color: C.accent1, margin: 0 } } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 9.0, h: 0.7, fontSize: 30, align: "left", bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "① 分方案标题" } },
      { placeholder: { options: { name: "theory", type: "body", x: 10.45, y: 0.52, w: 2.3, h: 0.46, fontSize: 12, bold: true, color: C.accent1, valign: "middle", margin: 0 }, text: "如：模块化" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.9, y: 2.15, w: 6.0, h: 4.3, fontSize: 16, color: C.text1, valign: "top", margin: 0 }, text: "具体动作（3-4 条）" } },
      { placeholder: { options: { name: "evidence", type: "body", x: 7.8, y: 2.15, w: 4.65, h: 4.3, fontSize: 16, color: C.text1, valign: "top", margin: 0 }, text: "对应调研数据 / 受访者原话" } },
      { placeholder: { options: { name: "source", type: "body", x: 8.05, y: 6.93, w: 4.7, h: 0.4, fontSize: 10, color: C.accent4, valign: "middle", margin: 0 }, text: "如：面访 n=12" } },
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
