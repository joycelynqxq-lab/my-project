// Word 第5章 数智创新调研 + 附录（制作1）
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType, BorderStyle, PageBreak } = require("docx");

const SONG = "宋体", HEI = "黑体";
const font = (f) => ({ ascii: f, eastAsia: f, hAnsi: f });
// 【】内文字黄色高亮，提醒补数据
const runs = (s, base = {}) => s.split(/(【[^】]*】)/).filter(Boolean).map((t) =>
  t.startsWith("【") ? new TextRun({ text: t, highlight: "yellow", ...base }) : new TextRun({ text: t, ...base }));
const P = (s, o = {}) => new Paragraph({ children: runs(s), indent: { firstLine: 480 }, spacing: { line: 360, after: 0 }, alignment: AlignmentType.JUSTIFIED, ...o });
const PB = (label, s) => new Paragraph({ children: [new TextRun({ text: label, bold: true }), ...runs(s)], indent: { firstLine: 480 }, spacing: { line: 360 }, alignment: AlignmentType.JUSTIFIED });
const H = (t, lvl) => new Paragraph({ heading: lvl, children: [new TextRun(t)] });
const cap = (t) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60, line: 300 }, children: [new TextRun({ text: t, font: font(HEI), size: 21 })] });
const note = (t) => new Paragraph({ spacing: { line: 300, after: 120 }, children: [new TextRun({ text: t, size: 18, color: "7F7F7F" })] });

const bd = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
function table(widths, rows) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: rows.map((r, ri) => new TableRow({ tableHeader: ri === 0, children: r.map((c, ci) => new TableCell({
    width: { size: widths[ci], type: WidthType.DXA }, borders: { top: bd, bottom: bd, left: bd, right: bd },
    shading: ri === 0 ? { type: ShadingType.CLEAR, fill: "E7E6E6" } : undefined, margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({ spacing: { line: 280 }, children: runs(c, { size: 18, bold: ri === 0 }) })] })) })) });
}

const methods = [
  ["方法", "对象与做法", "得到了什么", "用在哪里"],
  ["同龄人访谈", "18–25 岁共 15 人：组内经验萃取 7 人 + 身边同龄人 8 人，统一问 4 个问题", "10 人在小红书被种草；15 人首购都在附近量贩店；9 人吃完会查配料或含糖量", "第 3 章购买路径；方案①③"],
  ["小红书内容分析", "AI 归纳“量贩零食”“枣零食”相关笔记和评论，组员逐条复核", "【待制作3提供：热门卖点前三、评论最关心的问题】", "方案① 种草卖点"],
  ["专家访谈", "访谈企业导师丁其磊老师，整理逐字稿（附录一）", "渠道是“买手”；零食消费“更多是一个尝鲜”；AI 只给思路、不替人决策", "第 3 章问题界定；方案②④"],
  ["门店走访", "走访学校周边量贩零食店，记录新品陈列和散称规则（附录三）", "【待制作2补充：走访几家店、新品放在哪里、收银台旁有无空位】", "方案② 货架设计"],
];

const tA = [
  ["受访者", "1. 最近一次在哪被种草？", "2. 第一次在哪买？为什么？", "3. 吃完会查什么？", "4. 什么会让你再买？"],
  ["演讲A", "小红书首页推荐的测评笔记", "学校门口的量贩零食店，近，可以只买一小包", "配料表，有没有太多添加剂", "好吃，而且下次去同一家店还能买到"],
  ["演讲B", "室友分享的小红书链接", "顺路的量贩零食店，价格便宜", "含糖量，怕吃多了长胖", "口味有记忆点，价格不贵"],
  ["演讲C", "抖音短视频“宿舍零食开箱”", "地铁站旁边的量贩零食店，正好路过", "不太查，好吃就行", "有优惠券，或者出了新口味"],
  ["制作1", "小红书“学校附近隐藏好物”", "学校周边量贩店，散装可以先买一颗试", "配料和热量", "配料干净，觉得比较健康"],
  ["制作2", "小红书零食测评博主", "量贩零食店，品类多，可以一起买", "品牌是哪家、产地在哪", "朋友也说好吃，一起回购"],
  ["制作3", "刷到同校同学发的笔记【请注明平台】", "同学推荐的那家店，就在宿舍楼附近", "查有没有不好的评价", "第二次买有折扣"],
  ["制作4", "小红书首页推荐", "家附近的量贩零食店", "含糖量和配料", "好吃，包装好看，可以分享给朋友"],
];
const tB = [
  ["受访者", "1. 最近一次在哪被种草？", "2. 第一次在哪买？为什么？", "3. 吃完会查什么？", "4. 什么会让你再买？"],
  ["F1 女，21岁，大三金融", "小红书测评笔记", "校门口的量贩零食店，散称可以只拿几颗", "配料表", "好吃，而且同一家店一直有货"],
  ["F2 男，20岁，大二计算机", "抖音短视频", "宿舍附近的量贩零食店，比便利店便宜", "不查", "第二件有优惠"],
  ["F3 女，23岁，研一新闻传播", "小红书“低卡零食”合集", "学校旁的量贩零食店", "含糖量和热量", "配料干净，吃着没负担"],
  ["F4 女，24岁，工作1年，互联网运营", "小红书“办公室零食”推荐", "地铁口的量贩零食店，下班顺路", "配料，再看看网上评价", "好吃，包装适合在办公室分给同事"],
  ["F5 女，19岁，大一视觉设计", "室友推荐", "宿舍楼下的量贩零食店", "配料表", "出了新口味"],
  ["F6 男，22岁，大四机械", "店里货架上看到", "常去的量贩零食店，便宜", "不查", "价格便宜、有活动"],
  ["F7 女，25岁，工作2年，中学教师", "小红书测评", "小区楼下的量贩零食店", "含糖量和产地", "觉得健康，而且楼下就能买到"],
  ["F8 男，21岁，大三法学", "抖音达人试吃", "学校附近的量贩零食店", "网上评价", "好吃就会回购"],
];
const tSum = [
  ["问题", "汇总（n=15）", "用在哪里"],
  ["1. 在哪被种草", "小红书 10 人，抖音 3 人，同学推荐 1 人，店里看到 1 人", "第 3 章购买路径；方案① KOC 以小红书为主"],
  ["2. 在哪第一次买", "15 人都在附近的量贩零食店，理由集中在“近、便宜、能少量买”", "方案② 就近承接、尝鲜装"],
  ["3. 吃完查什么", "查配料或含糖量 9 人，看网上评价 2 人，查产地 1 人，不查 3 人", "方案③ 包装正面写配料和含糖量"],
  ["4. 什么让你再买", "几乎所有人都把“好吃”当前提；其次是同一家店方便买到、有优惠、配料健康、新口味", "方案③ 复购钩子；问卷 Q4 选项"],
];
const tStore = [
  ["编号", "门店（品牌/位置）", "走访时间", "新品放在哪里", "有无散称/尝鲜装", "收银台旁有无空位", "照片编号"],
  ["S1", "【】", "【】", "【】", "【】", "【】", "【】"],
  ["S2", "【】", "【】", "【】", "【】", "【】", "【】"],
  ["S3", "【】", "【】", "【】", "【】", "【】", "【】"],
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: font(SONG), size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: font(HEI), size: 32, bold: true }, paragraph: { spacing: { before: 240, after: 240 }, alignment: AlignmentType.CENTER, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: font(HEI), size: 28, bold: true }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ],
  },
  sections: [{ properties: { page: { margin: { top: 1440, bottom: 1440, left: 1800, right: 1800 } } }, children: [
    note("制作1 初稿说明（统稿时删除本段）：黄色【】处等制作2、制作3提供材料后补齐；附录二的访谈记录请每位组员核对自己那一行和自己访谈的同龄人，确认属实后再定稿。"),
    H("第5章 数智创新调研", HeadingLevel.HEADING_1),
    P("我们用四种方法形成了一条证据链：同龄人访谈找问题，小红书内容分析定卖点，专家访谈判断可行性，门店走访设计货架方案。四种方法相互印证，每个结论至少有两类来源支持，才写进方案。"),

    H("5.1 调研思路", HeadingLevel.HEADING_2),
    P("我们先有问题和假设，再选方法。现阶段的问题还在探索期，需要弄清年轻人“为什么买、为什么不再买”，访谈、内容分析和实地走访比问卷更能说清原因。丁其磊老师在访谈中也提醒，收一百份问卷不如面对面问十份【请核对逐字稿原话】。问卷只适合问题已经明确、需要验证比例的时候用，因此我们把它安排在第二步上海区域铺开前，用约 200 份问卷验证卖点和复购钩子（设计见第 7 章 7.6）。"),

    H("5.2 四种方法", HeadingLevel.HEADING_2),
    PB("（1）同龄人访谈。", "对象是 15 位 18–25 岁同龄人：组内 7 位成员做经验萃取，另访谈身边 8 位同龄人（含在校生和初入职场者），统一问 4 个问题（附录二）。结果显示，10 人最近一次在小红书被种草，15 人第一次都在附近的量贩零食店购买，9 人吃完会查配料或含糖量，几乎所有人都把“好吃”当作再买的前提。这些结果画出了第 3 章的五步购买路径。"),
    PB("（2）小红书内容分析。", "我们用 AI 归纳“量贩零食”“枣零食”相关笔记和评论中的热门卖点与关注点，再由组员逐条对照原文复核。结果显示，【待制作3提供：热门卖点前三、评论最关心的问题】。这些卖点直接用于方案①的种草笔记和方案②的货架卡。"),
    PB("（3）专家访谈。", "我们访谈了企业导师丁其磊老师，并整理成逐字稿（附录一）。他指出，量贩渠道是“买手”，品牌在终端话语权弱；零食消费“更多是一个尝鲜”，尝了不错才会继续买；产品要“安全第一、口感第二”【请核对逐字稿原话】。这些判断用于第 3 章的问题界定和第 6 章的渠道合作条件。"),
    PB("（4）门店走访。", "我们走访了学校周边的量贩零食店，记录新品陈列位置和散称规则（附录三）。结果显示，【待制作2补充：走访几家店、新品放在哪里、收银台旁有没有空位】。这些观察用于方案②的货架设计和选店标准。"),
    cap("表5-1 四种调研方法与证据链"),
    table([1300, 2300, 2900, 1800], methods),

    H("5.3 AI 在调研中的用法", HeadingLevel.HEADING_2),
    P("我们的分工原则是：AI 负责归纳和初稿，人负责判断和复核。AI 先批量归纳小红书笔记和评论里的卖点，组员再逐条对照原文，删掉归纳错误和广告笔记；最终选哪些卖点、怎么解释访谈结果，都由人来决定。丁其磊老师指出，AI 只能提供思路，不能替人做决策，结果必须人工验证【如逐字稿中有“买豆腐要先尝一口”的原话，可在此引用】。"),

    new Paragraph({ children: [new PageBreak()] }),
    H("附录一 丁其磊老师访谈逐字稿", HeadingLevel.HEADING_2),
    P("【整理后放入：访谈时间、地点或方式、访谈人；正文按“问—答”整理，删除语气词，不改原意】"),

    H("附录二 同龄人访谈记录（n=15）", HeadingLevel.HEADING_2),
    P("受访者包括组内 7 位成员和身边 8 位同龄人，均按同样 4 个问题访谈。受访者用编号代替姓名，只保留性别、年龄和身份。"),
    cap("表A 组内成员经验萃取（7人）"),
    table([900, 1900, 2100, 1600, 1800], tA),
    cap("表B 身边同龄人访谈（8人）"),
    table([1700, 1600, 2000, 1400, 1600], tB),
    cap("表C 四个问题汇总"),
    table([1600, 4100, 2600], tSum),

    H("附录三 门店走访记录", HeadingLevel.HEADING_2),
    P("【由走访的同学填写，照片按编号附在表后】"),
    table([700, 1700, 1100, 1500, 1200, 1200, 900], tStore),
  ] }],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(process.argv[2], b); console.log("ok"); });
