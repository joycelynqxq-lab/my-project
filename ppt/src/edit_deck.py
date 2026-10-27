import sys, os
U = sys.argv[1]  # unpacked dir
E = 914400
nid = [900]

def emu(v): return str(int(round(v * E)))

def tbox(x, y, w, h, runs, algn="r", anchor="t"):
    nid[0] += 1
    rs = ""
    for t, sz, b, col in runs:
        rs += (f'<a:r><a:rPr lang="zh-CN" altLang="en-US" sz="{sz}" b="{1 if b else 0}" dirty="0"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
               f'<a:latin typeface="微软雅黑"/><a:ea typeface="微软雅黑"/></a:rPr><a:t>{t}</a:t></a:r>')
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid[0]}" name="修订 文本 {nid[0]}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="{anchor}"><a:noAutofit/></a:bodyPr>'
            f'<a:p><a:pPr algn="{algn}"/>{rs}</a:p></p:txBody></p:sp>')

def rect(x, y, w, h, fill, prst="rect", text=None):
    nid[0] += 1
    body = '<p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:p><a:pPr algn="ctr"/>'
    if text:
        body += (f'<a:r><a:rPr lang="zh-CN" altLang="en-US" sz="1000" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
                 f'<a:latin typeface="微软雅黑"/><a:ea typeface="微软雅黑"/></a:rPr><a:t>{text}</a:t></a:r>')
    body += '</a:p></p:txBody>'
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid[0]}" name="修订 形状 {nid[0]}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>{body}</p:sp>')

RED, GREY = "B62F06", "6E6266"
def evidence(y, t, x=1.15): return tbox(x, y, 7.0, 0.24, [("证据｜", 1000, True, RED), (t, 1000, False, GREY)], algn="l")

def edit(n, reps=(), add=""):
    p = os.path.join(U, f"ppt/slides/slide{n}.xml")
    x = open(p, encoding="utf-8").read()
    for a, b in reps:
        assert x.count(a) == 1, (n, a, x.count(a)); x = x.replace(a, b)
    if add:
        i = x.rfind("</p:spTree>"); x = x[:i] + add + x[i:]
    open(p, "w", encoding="utf-8").write(x)

edit(2, [("调研设计与发现", "调研设计")])
edit(3, [("Marketing overview and objectives", "Background and Problem")])
edit(4, add=evidence(7.06, "二手资料：鸣鸣很忙 2025 年度业绩（虎嗅、观察者网报道）；观研报告网"))

# 第 5 页：现象层 → 初步判断（待验证），两条判断配二手资料
t1_old = '<a:off x="6949466" y="2244109"/><a:ext cx="1710055" cy="460375"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:p><a:pPr algn="r"/>'
t1_new = '<a:off x="6949466" y="2244109"/><a:ext cx="4572000" cy="460375"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:noAutofit/></a:bodyPr><a:p><a:pPr algn="l"/>'
t2_old = '<a:off x="6949466" y="3819307"/><a:ext cx="3944126" cy="461010"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"><a:noAutofit/></a:bodyPr><a:p><a:pPr algn="r"/>'
t2_new = t2_old.replace('cx="3944126"', 'cx="4572000"').replace('algn="r"', 'algn="l"')
edit(5, [(t1_old, t1_new), (t2_old, t2_new), ("<a:t>拿起率较低</a:t>", "<a:t>新品难被第一次拿起</a:t>"), ("食药同材", "食药同源")],
     add=rect(6.5, 1.27, 2.6, 0.58, "FFFAF4")
     + tbox(6.54, 1.33, 5.0, 0.48, [("我们的初步判断", 2000, True, "860416"), ("（待验证）", 1400, True, RED)], algn="l", anchor="ctr")
     + tbox(7.6, 2.93, 5.6, 0.26, [("二手资料：门店近 2000 个 SKU，每月 10%–20% 末位淘汰（华创证券）", 1000, False, GREY)], algn="l")
     + tbox(7.6, 4.64, 5.6, 0.26, [("二手资料：红枣是认知度前三的“食药同源”食材（尼尔森）", 1000, False, GREY)], algn="l")
     + evidence(7.06, "专家访谈（丁其磊老师）；二手资料：华创证券、尼尔森；初步判断待调研验证"))
edit(6, add=evidence(7.06, "二手资料 + 组内定性分析"))
edit(7, add=tbox(8.3, 0.4, 4.75, 0.4, [("理论：价值网络＋服务主导逻辑", 1600, True, "7F7F7F")], algn="r", anchor="ctr")
     + rect(7.3, 4.86, 1.45, 0.28, RED, prst="roundRect", text="假设 H1 · 待验证")
     + evidence(7.25, "专家访谈（丁其磊老师）；组内分析", x=0.49))
print("done")
