// 制作1：Word 第5章（定稿版）+ 个人作业
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType, BorderStyle, PageBreak, VerticalAlign } = require("docx");

const SONG = { ascii: "Times New Roman", eastAsia: "宋体", hAnsi: "Times New Roman" };
const HEI = { ascii: "Times New Roman", eastAsia: "黑体", hAnsi: "Times New Roman" };
const XS = 24; // 小四
const LINE = 360; // 1.5 倍行距

const P = (parts, o = {}) => new Paragraph({
  children: (Array.isArray(parts) ? parts : [parts]).map((t) => typeof t === "string" ? new TextRun(t) : new TextRun(t)),
  indent: { firstLine: 480 }, spacing: { line: LINE, after: 0 }, alignment: AlignmentType.JUSTIFIED, ...o });
const LP = (label, text, o = {}) => P([{ text: label, bold: true }, text], o);
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const cap = (t) => new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120, after: 60, line: 300 }, children: [new TextRun({ text: t, font: HEI, size: 21 })] });

const bd = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
function table(widths, rows, center = []) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, alignment: AlignmentType.CENTER,
    rows: rows.map((r, ri) => new TableRow({ tableHeader: ri === 0, cantSplit: true, children: r.map((c, ci) => new TableCell({
      width: { size: widths[ci], type: WidthType.DXA }, borders: { top: bd, bottom: bd, left: bd, right: bd }, verticalAlign: VerticalAlign.CENTER,
      shading: ri === 0 ? { type: ShadingType.CLEAR, fill: "E7E6E6" } : undefined, margins: { top: 50, bottom: 50, left: 90, right: 90 },
      children: [new Paragraph({ alignment: ri === 0 || center.includes(ci) ? AlignmentType.CENTER : AlignmentType.LEFT, spacing: { line: 280 },
        children: [new TextRun({ text: c, size: 21, bold: ri === 0, font: ri === 0 ? HEI : SONG })] })] })) })) });
}
const styles = {
  default: { document: { run: { font: SONG, size: XS } } },
  paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: HEI, size: 32, bold: true }, paragraph: { spacing: { before: 240, after: 240, line: LINE }, alignment: AlignmentType.CENTER, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: HEI, size: 28, bold: true }, paragraph: { spacing: { before: 200, after: 80, line: LINE }, keepNext: true, outlineLevel: 1 } },
  ],
};
const page = { margin: { top: 1440, bottom: 1440, left: 1800, right: 1800 } };

// ================= 第5章 =================
const map = [
  ["调研方法", "种草", "尝鲜", "验证", "复购", "支撑的方案"],
  ["同龄人访谈（n=15）", "●", "●", "●", "●", "第3章购买路径；方案①③"],
  ["小红书内容分析", "●", "", "●", "", "方案①笔记卖点；方案②货架卡"],
  ["专家访谈（丁其磊老师）", "", "●", "", "●", "方案②④渠道合作条件"],
  ["门店走访", "", "●", "", "", "方案②货架与选店"],
];
const tA = [
  ["受访者", "最近一次在哪被种草", "第一次在哪买、为什么", "吃完会查什么", "什么会让你再买"],
  ["演讲A", "小红书首页推荐的测评笔记", "学校门口的量贩零食店，近，可以只买一小包", "配料表，有没有太多添加剂", "好吃，而且下次去同一家店还能买到"],
  ["演讲B", "室友分享的小红书链接", "顺路的量贩零食店，价格便宜", "含糖量，怕吃多了长胖", "口味有记忆点，价格不贵"],
  ["演讲C", "抖音短视频“宿舍零食开箱”", "地铁站旁边的量贩零食店，正好路过", "不太查，好吃就行", "有优惠券，或者出了新口味"],
  ["制作1", "小红书“学校附近隐藏好物”", "学校周边量贩店，散装可以先买一颗试", "配料和热量", "配料干净，觉得比较健康"],
  ["制作2", "小红书零食测评博主", "量贩零食店，品类多，可以一起买", "品牌是哪家、产地在哪", "朋友也说好吃，一起回购"],
  ["制作3", "刷到同校同学发的笔记", "同学推荐的那家店，就在宿舍楼附近", "查有没有不好的评价", "第二次买有折扣"],
  ["制作4", "小红书首页推荐", "家附近的量贩零食店", "含糖量和配料", "好吃，包装好看，可以分享给朋友"],
];
const tB = [
  ["受访者", "最近一次在哪被种草", "第一次在哪买、为什么", "吃完会查什么", "什么会让你再买"],
  ["F1 女，21岁，大三金融", "小红书测评笔记", "校门口的量贩零食店，散称可以只拿几颗", "配料表", "好吃，而且同一家店一直有货"],
  ["F2 男，20岁，大二计算机", "抖音短视频", "宿舍附近的量贩零食店，比便利店便宜", "不查", "第二件有优惠"],
  ["F3 女，23岁，研一新闻传播", "小红书“低卡零食”合集", "学校旁的量贩零食店", "含糖量和热量", "配料干净，吃着没负担"],
  ["F4 女，24岁，互联网运营", "小红书“办公室零食”推荐", "地铁口的量贩零食店，下班顺路", "配料，再看网上评价", "好吃，包装适合分给同事"],
  ["F5 女，19岁，大一视觉设计", "室友推荐", "宿舍楼下的量贩零食店", "配料表", "出了新口味"],
  ["F6 男，22岁，大四机械", "店里货架上看到", "常去的量贩零食店，便宜", "不查", "价格便宜、有活动"],
  ["F7 女，25岁，中学教师", "小红书测评", "小区楼下的量贩零食店", "含糖量和产地", "觉得健康，楼下就能买到"],
  ["F8 男，21岁，大三法学", "抖音达人试吃", "学校附近的量贩零食店", "网上评价", "好吃就会回购"],
];
const tSum = [
  ["问题", "汇总（n=15）"],
  ["在哪被种草", "小红书10人，抖音3人，同学推荐1人，店里看到1人"],
  ["在哪第一次买", "15人都在附近的量贩零食店，理由集中在“近、便宜、能少量买”"],
  ["吃完查什么", "查配料或含糖量9人，看网上评价2人，查产地1人，不查3人"],
  ["什么让你再买", "几乎所有人都把“好吃”当前提；其次是同一家店方便买到、有优惠、配料健康、新口味"],
];
const tStore = [
  ["编号", "门店位置", "走访时间", "新品陈列位置", "散称/小包装", "收银台旁空位", "照片"],
  ["S1", "", "", "", "", "", ""], ["S2", "", "", "", "", "", ""], ["S3", "", "", "", "", "", ""],
];

const ch5 = new Document({ styles, sections: [{ properties: { page }, children: [
  H1("第5章 数智创新调研"),
  P("我们用四种方法搭起了一条证据链：同龄人访谈找到问题，小红书内容分析确定卖点，专家访谈判断渠道可行性，门店走访设计货架方案。每个写进方案的结论，至少有两类来源相互印证。"),

  H2("5.1 调研思路"),
  P("我们先定问题，再选方法。年轻人“为什么第一次买、为什么不再买”是探索性问题，所以以定性方法为主；丁其磊老师也提醒，“收一百份问卷不如面对面问十份”。问卷只适合验证比例，因此放在第二步上海区域铺开前，用约200份验证卖点和复购钩子（见7.6节）。调研以第3章的四步购买路径为骨架，每一步至少有两种方法覆盖（表1）。"),
  cap("表1 调研方法与购买路径、方案的对应"),
  table([2300, 700, 700, 700, 700, 2600], map, [1, 2, 3, 4]),

  H2("5.2 四种方法"),
  LP("（1）同龄人访谈。", "组内7人做经验萃取，另访谈身边8位18–25岁同龄人，统一问4个问题（附录二）。15人中，10人最近一次在小红书被种草，15人都在附近量贩店首购，9人吃完会查配料或含糖量，再买首先看“好吃”。这些结果画出了四步购买路径，也决定了方案①以小红书为主阵地、方案③在包装正面写清配料和含糖量。"),
  LP("（2）小红书内容分析。", "AI归纳“量贩零食”“枣零食”相关笔记和评论的高频卖点与疑问，组员对照原文复核、剔除广告笔记。结果与访谈交叉后，形成方案①的候选卖点：新奇口味、一口一个、配料干净，也用于方案②的货架卡。"),
  LP("（3）专家访谈。", "访谈企业导师丁其磊老师并整理逐字稿（附录一）。他指出量贩渠道是“买手”，愿意要能卖得动的新品；零食消费“更多是一个尝鲜”，尝了不错才会继续买；产品要“安全第一、口感第二”。这些判断用于第3章的问题界定和第6章的门店合作条件。"),
  LP("（4）门店走访。", "走访学校周边量贩零食店，记录新品陈列位置、散称规则和收银台旁空位（附录三）。据此，方案②把尝鲜装放在收银台旁或新品区，正装放在常规货架。"),

  H2("5.3 AI在调研中的用法"),
  P("AI负责归纳和初稿，人负责判断和复核：AI批量整理笔记和评论，组员回到原文核对，再决定哪些卖点进入方案。丁其磊老师的比喻是，AI的结果“就像买豆腐自己要先尝一口”，必须人工验证；他也建议“一部分人做、一部分AI做，互相对比”。所以第二步会把问卷结论与AI归纳的卖点对照，一致才写进下一阶段方案。"),

  new Paragraph({ children: [new PageBreak()] }),
  H2("附录一 丁其磊老师访谈逐字稿"),
  P("（整理后的逐字稿按“问—答”顺序放入，删去语气词，不改原意。）"),
  H2("附录二 同龄人访谈记录（n=15）"),
  P("受访者包括组内7位成员和身边8位同龄人，均按同样4个问题访谈；受访者以编号代替姓名，只保留性别、年龄和身份。"),
  cap("表A 组内成员经验萃取（7人）"),
  table([900, 1900, 2100, 1600, 1800], tA),
  cap("表B 身边同龄人访谈（8人）"),
  table([1800, 1600, 2000, 1400, 1500], tB),
  cap("表C 四个问题汇总"),
  table([1700, 6600], tSum),
  H2("附录三 门店走访记录"),
  P("（由走访同学填写，照片按编号附在表后。）"),
  table([700, 1700, 1100, 1500, 1200, 1200, 900], tStore, [0]),
] }] });

// ================= 个人作业 =================
const me = new Document({ styles, sections: [{ properties: { page: { margin: { top: 1200, bottom: 1100, left: 1500, right: 1500 } } }, children: [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60, line: LINE }, children: [new TextRun({ text: "个人作业：从“3秒首购”到“一口尝鲜，一路复购”", font: HEI, size: 30, bold: true })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120, line: 300 }, children: [new TextRun({ text: "张紫祺　24901073　｜　小组分工：PPT 制作1、Word 第5章", size: 21 })] }),
  LP("概述。", "我们组为美枣王设计了一套让18–25岁年轻人在量贩零食店“一口尝鲜、一路复购”的增长方案。我负责PPT母版和“背景与问题”部分，画了年轻人的四步购买路径图，并撰写Word第5章“数智创新调研”。"),
  LP("进展和经验。", "第一版方案叫“3秒首购”，只盯着货架前的第一次购买，背景和调研占了一半篇幅。老师辅导时追问“只尝不留，首购提升有什么意义”，我们才把核心问题改为“通过首购体验带动复购”。我最大的收获有三点：一是理论是防漏清单而不是标签，做PEST四宫格时“政策”一格起初是空的，补查后才发现营养标签新国标要求2027年起强制标示糖，正好支撑了“包装提前回答健不健康”的方案；二是证据要贴着观点放，每页写清理论和来源，逼着我删掉了找不到出处的数据；三是方法要匹配问题，问卷从调研主角变成了试点之后的验证工具。"),
  LP("不足。", "同龄人访谈只有15人，其中7人是组员自己的经验萃取，样本偏向在校生；行业数据都来自二手公开资料；小红书分析依赖AI归纳，尚未量化；成本测算仍是假设值，需要试点数据修正。"),
  LP("执行关键点。", "①线上线下说同一句话：笔记、货架卡和包装用同一组卖点和图片；②每位KOC一个专属核销码，从曝光到首购、复购全程可追踪；③第一包只放最好吃的口味，并把配料和含糖量印在正面。"),
  LP("资源与能力保障。", "按价值共创的思路，方案需要四方各出一份资源：美枣王出小规格产线、样品和配料信息，门店出收银台旁的位置和销售数据，校园KOC出内容和同校影响力，年轻消费者出评价和分享。从资源基础观看，美枣王真正要培养的是“小范围测试—复盘—复制”的组织能力，而不是一次性的广告投放。"),
  LP("总结。", "项目从“如何提高首购率”出发，在老师和企业导师的反馈中，逐步变成一套可测试、可复制的闭环。我们希望它从一所学校跑通，成为美枣王在量贩渠道年轻化的增长模式。"),
] }] });

Promise.all([Packer.toBuffer(ch5), Packer.toBuffer(me)]).then(([a, b]) => {
  fs.writeFileSync(process.argv[2], a); fs.writeFileSync(process.argv[3], b); console.log("ok");
});
