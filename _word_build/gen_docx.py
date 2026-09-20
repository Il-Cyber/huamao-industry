# -*- coding: utf-8 -*-
"""生成《花茂村产业图谱——数据与内容整理》Word 文档"""
import docx
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

CN_BODY = '宋体'
CN_HEAD = '黑体'
EN = 'Arial'
BLACK = RGBColor(0, 0, 0)
GRAY_HDR = 'D9D9D9'

OUT = r'D:\Users\li\Desktop\花茂村产业图谱-数据与内容整理.docx'

doc = Document()

# ---------- 页面设置 ----------
sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(2.5)
sec.bottom_margin = Cm(2.5)
sec.left_margin = Cm(2.5)
sec.right_margin = Cm(2.5)

# ---------- 样式 ----------
def style_font(st, cn, en, size, bold=None, color=None):
    st.font.name = en
    st.font.size = Pt(size)
    if bold is not None:
        st.font.bold = bold
    if color is not None:
        st.font.color.rgb = color
    rpr = st.element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), cn)
    rfonts.set(qn('w:ascii'), en)
    rfonts.set(qn('w:hAnsi'), en)

def set_run(run, cn=CN_BODY, en=EN, size=12, bold=None, color=None, italic=None):
    run.font.name = en
    run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    if italic is not None:
        run.font.italic = italic
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), cn)
    rfonts.set(qn('w:ascii'), en)
    rfonts.set(qn('w:hAnsi'), en)

# Normal 正文
normal = doc.styles['Normal']
style_font(normal, CN_BODY, EN, 12, color=BLACK)
pf = normal.paragraph_format
pf.line_spacing = 1.5
pf.first_line_indent = Pt(24)  # 2 字符
pf.space_before = Pt(0)
pf.space_after = Pt(0)
pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

# 标题样式
h1 = doc.styles['Heading 1']
style_font(h1, CN_HEAD, EN, 16, bold=True, color=BLACK)
h1.paragraph_format.space_before = Pt(14)
h1.paragraph_format.space_after = Pt(6)
h1.paragraph_format.line_spacing = 1.2
h1.paragraph_format.first_line_indent = Pt(0)

h2 = doc.styles['Heading 2']
style_font(h2, CN_HEAD, EN, 14, bold=True, color=BLACK)
h2.paragraph_format.space_before = Pt(11)
h2.paragraph_format.space_after = Pt(5)
h2.paragraph_format.line_spacing = 1.2
h2.paragraph_format.first_line_indent = Pt(0)

h3 = doc.styles['Heading 3']
style_font(h3, CN_HEAD, EN, 12, bold=True, color=BLACK)
h3.paragraph_format.space_before = Pt(9)
h3.paragraph_format.space_after = Pt(4)
h3.paragraph_format.line_spacing = 1.2
h3.paragraph_format.first_line_indent = Pt(0)

cap = doc.styles['Caption']
style_font(cap, CN_BODY, EN, 10.5, color=BLACK)
cap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
cap.paragraph_format.first_line_indent = Pt(0)
cap.paragraph_format.space_before = Pt(4)
cap.paragraph_format.space_after = Pt(4)
cap.paragraph_format.line_spacing = 1.0

# ---------- 工具函数 ----------
def para(text='', style=None, indent=True, align=None, size=12, bold=None, cn=CN_BODY):
    p = doc.add_paragraph(style=style)
    if text:
        r = p.add_run(text)
        set_run(r, cn=cn, size=size, bold=bold)
    if not indent:
        p.paragraph_format.first_line_indent = Pt(0)
    if align is not None:
        p.alignment = align
    return p

def rich(parts, indent=True, align=None, size=12):
    """parts: list of (text, bold, italic, cn_font)"""
    p = doc.add_paragraph()
    for t, b, i, f in parts:
        r = p.add_run(t)
        set_run(r, cn=f, size=size, bold=b, italic=i)
    if not indent:
        p.paragraph_format.first_line_indent = Pt(0)
    if align is not None:
        p.alignment = align
    return p

def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def cell_text(cell, text, bold=False, align=None, size=10.5):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    if align is not None:
        p.alignment = align
    r = p.add_run(text)
    set_run(r, size=size, bold=bold)

def make_table(headers, rows, widths, aligns=None, caption=None):
    if caption:
        para(caption, style='Caption', indent=False, size=10.5)
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        cell_text(c, h, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        shade(c, GRAY_HDR)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for ci, val in enumerate(row):
            a = None
            if aligns:
                a = aligns[ci]
            cell_text(cells[ci], val, align=a)
    # 列宽
    for ci, w in enumerate(widths):
        for row in t.rows:
            row.cells[ci].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    # 收尾空段不缩进，行距1
    last = doc.paragraphs[-1]
    last.paragraph_format.first_line_indent = Pt(0)
    last.paragraph_format.line_spacing = 1.0
    last.paragraph_format.space_before = Pt(0)
    last.paragraph_format.space_after = Pt(2)
    return t

def add_hyperlink(paragraph, url, text, size=12):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), EN)
    rFonts.set(qn('w:hAnsi'), EN)
    rFonts.set(qn('w:eastAsia'), CN_BODY)
    rPr.append(rFonts)
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(size * 2)))
    rPr.append(sz)
    color = OxmlElement('w:color'); color.set(qn('w:val'), '9E2B25')
    rPr.append(color)
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single')
    rPr.append(u)
    new_run.append(rPr)
    t = OxmlElement('w:t'); t.text = text
    t.set(qn('xml:space'), 'preserve')
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def cell_source(cell, media, urls):
    """单元格内容：第一段媒体名+日期，后续每段一个原文链接。urls 可为 None/str/list"""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p1 = cell.paragraphs[0]
    p1.paragraph_format.first_line_indent = Pt(0)
    p1.paragraph_format.line_spacing = 1.0
    p1.paragraph_format.space_after = Pt(2)
    r1 = p1.add_run(media)
    set_run(r1, size=10.5)
    if urls is None:
        return
    if isinstance(urls, str):
        urls = [urls]
    for u in urls:
        p = cell.add_paragraph()
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        add_hyperlink(p, u, '（原文链接）', size=10.5)

def page_field(paragraph):
    r = paragraph.add_run()
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = 'PAGE'
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'end')
    r._r.append(f1); r._r.append(it); r._r.append(f2)
    set_run(r, size=10.5)

def toc_field(p):
    r = p.add_run()
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve')
    it.text = 'TOC \\o "1-2" \\h \\z \\u'
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.text = '（打开后请右键此处选择“更新域”以生成目录）'
    f3 = OxmlElement('w:fldChar'); f3.set(qn('w:fldCharType'), 'end')
    r._r.append(f1); r._r.append(it); r._r.append(f2); r._r.append(t); r._r.append(f3)
    set_run(r, size=12)

# ================= 第 1 节：标题 + 目录 =================
p = para('花茂村产业图谱', indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, size=18, bold=True, cn=CN_HEAD)
p.paragraph_format.space_before = Pt(24)
p.paragraph_format.space_after = Pt(6)
p.paragraph_format.line_spacing = 1.2

p = para('——数据与内容整理', indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, bold=True, cn=CN_HEAD)
p.paragraph_format.space_after = Pt(6)
p.paragraph_format.line_spacing = 1.2

p = para('贵州省遵义市播州区枫香镇花茂村 · 农业链 / 非遗文旅链 / 红色文化', indent=False,
         align=WD_ALIGN_PARAGRAPH.CENTER, size=12)
p.paragraph_format.space_after = Pt(18)

para('目录', indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, size=16, bold=True, cn=CN_HEAD)
toc_p = doc.add_paragraph()
toc_p.paragraph_format.first_line_indent = Pt(0)
toc_p.paragraph_format.line_spacing = 1.0
toc_field(toc_p)

# 分节：正文
sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
sec2.top_margin = Cm(2.5)
sec2.bottom_margin = Cm(2.5)
sec2.left_margin = Cm(2.5)
sec2.right_margin = Cm(2.5)
# 正文页码从 1 开始
sectPr = sec2._sectPr
pg = OxmlElement('w:pgNumType'); pg.set(qn('w:start'), '1')
sectPr.append(pg)
# 页脚页码（独立于第 1 节）
sec2.footer.is_linked_to_previous = False
fp = sec2.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
fp.paragraph_format.first_line_indent = Pt(0)
page_field(fp)

# ================= 一、项目概述 =================
doc.add_heading('一、项目概述', level=1)
doc.add_heading('（一）项目背景', level=2)
para('花茂村位于贵州省遵义市播州区枫香镇。曾经土地碎片、人口外流的“荒茅田”，如今以农业链与非遗文旅链双轮驱动：土地变资产、农民变股东，陶窑与纸坊变体验空间。本图谱以六层产业关系总图、两条产业链路径与四张案例卡，呈现花茂产业“从资源到结果”的完整关系。')
doc.add_heading('（二）核心公开指标', level=2)
# 表1 核心公开指标（来源列含超链接）
para('表1  核心公开指标底表（点击“原文链接”可查看出处报道）', style='Caption', indent=False, size=10.5)
DJB = 'http://dangjian.people.com.cn/BIG5/n1/2026/0707/c448938-40755191.html'  # 人民网党建调研行 2026.7
t1 = doc.add_table(rows=1, cols=4)
t1.style = 'Table Grid'
t1.alignment = WD_TABLE_ALIGNMENT.CENTER
t1.autofit = False
hdr1 = ['指标', '数值', '年份', '来源（媒体 / 原文链接）']
for i, h in enumerate(hdr1):
    c = t1.rows[0].cells[i]
    cell_text(c, h, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade(c, GRAY_HDR)
rows1 = [
    ('接待游客', '71.25 万人次', '2025',
     '人民网党建频道《党建调研行》2026.7', DJB),
    ('旅游总收入', '2400 万元', '2025',
     '人民网党建频道《党建调研行》2026.7', DJB),
    ('村集体经济积累', '1481 万元', '2025',
     '人民网党建频道《党建调研行》2026.7', DJB),
    ('人均可支配收入', '29650 元', '2025',
     '人民网党建频道《党建调研行》2026.7', DJB),
    ('全村面积 / 耕地', '9.8 km² / 约 5531 亩', '调研期',
     '《花茂村实践报告》（内部调研材料，无公开链接）', None),
    ('户数 / 人口', '1345 户 / 约 5166 人', '调研期',
     '《花茂村实践报告》（内部调研材料，无公开链接）', None),
]
for ri, (a, b, c_, m, u) in enumerate(rows1):
    cells = t1.add_row().cells
    cell_text(cells[0], a, align=WD_ALIGN_PARAGRAPH.LEFT)
    cell_text(cells[1], b, align=WD_ALIGN_PARAGRAPH.RIGHT)
    cell_text(cells[2], c_, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_source(cells[3], m, u)
for ci, w in enumerate([3.0, 4.4, 1.8, 6.8]):
    for row in t1.rows:
        row.cells[ci].width = Cm(w)
doc.add_paragraph().paragraph_format.space_after = Pt(2)
last = doc.paragraphs[-1]
last.paragraph_format.first_line_indent = Pt(0)
last.paragraph_format.line_spacing = 1.0
last.paragraph_format.space_before = Pt(0)
para('以上公开数据只保留年份、单位与来源最清楚的指标；口径冲突或数量级存疑的数据一律进入“待核实数据清单”，不进入核心图表、不计算比例与增速。')

# ================= 二、点位地图 =================
doc.add_heading('二、点位地图（高德地图产业坐标）', level=1)
doc.add_heading('（一）点位清单', level=2)
make_table(
    ['点位', '所属链条', '坐标（经度, 纬度）', '标签', '简介'],
    [
        ['花茂村陶艺馆（母氏陶艺馆）', '非遗文旅链', '106.5773, 27.6162', '制陶·体验', '省级非遗花茂土陶第四代传承人母先才经营，位于陶艺文化创意一条街；游客可参观成品、现场拉坯体验。'],
        ['花茂人家·古法造纸', '非遗文旅链', '106.577666, 27.617247', '造纸·文创', '白泥组黔北民居群中的古法造纸工坊，纸浆压花画、小夜灯、冰箱贴等文创在此制作销售。'],
        ['红色之家', '非遗文旅链', '106.5775, 27.6159', '餐饮·故事', '2013 年创建的花茂村第一家农家乐（桶井组），盬子鸡与红军故事结合，电视剧《花繁叶茂》取景地之一。'],
        ['我在花茂有块田·认养区', '农业链', '106.5786, 27.6147', '农事·研学', '九丰农业科技公司依托乡愁元素打造的农业观光体验园：游客可认养农田，参与犁地、播种、采摘等农事体验。'],
        ['苟坝会议陈列馆', '红色文化', '106.569087, 27.636930', '红色·参观', '红色参观的重要起点，与花茂村文旅场景相衔接；坐标经前期项目实测核验。'],
        ['苟坝会议会址', '红色文化', '106.567372, 27.649610', '红色·遗址', '1935 年苟坝会议召开地，周边有会议遗址、真理小道、马灯馆等点位。'],
    ],
    [3.6, 2.0, 3.4, 1.9, 5.1],
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
    caption='表2  产业点位清单（6 个）',
)
doc.add_heading('（二）坐标口径说明', level=2)
para('苟坝会议陈列馆、会址坐标为前期项目《苟坝花茂红色文化数字地图》实测核验值；花茂人家坐标为高德 POI 数据；村内四点位坐标为高德 POI 检索或按村组位置估算，误差约 20–100 米，可到现场微调。')

# ================= 三、六层产业关系总图 =================
doc.add_heading('三、六层产业关系总图', level=1)
para('从“资源基础”到“乡村结果”，25 个节点按六层逻辑排布，每条连线都标注明确关系词（流转土地、组织生产、吸纳就业、提供体验、带动消费等）。')
para('节点与连线依据《花茂村实践报告》及三份产业内容设计文档等内部材料整合，非网络公开数据；其中引用的公开指标出处见表1 与第八章来源清单。')
doc.add_heading('（一）六层结构', level=2)
make_table(
    ['层级', '名称', '图示颜色'],
    [
        ['第一层', '资源基础', '#7A8B6F'],
        ['第二层', '组织主体', '#9E2B25'],
        ['第三层', '产业活动', '#2E6B4F'],
        ['第四层', '产品与服务', '#B9814A'],
        ['第五层', '市场渠道', '#5B7C99'],
        ['第六层', '乡村结果与问题', '#4A3B2E'],
    ],
    [2.6, 4.6, 8.8],
    aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
    caption='表3  六层结构说明',
)
doc.add_heading('（二）节点清单（25 个）', level=2)
make_table(
    ['编号', '节点', '层级', '链条', '说明'],
    [
        ['n1', '耕地与农田', '第一层', '农业', '全村国土面积 9.8 平方公里、耕地约 5531 亩；经整治后由零散小块变为连片生产空间。'],
        ['n2', '土陶技艺', '第一层', '非遗', '省级非物质文化遗产，烧制技艺传承千年，是“陶瓷之乡”的产业根基。'],
        ['n3', '古法造纸技艺', '第一层', '非遗', '外婆传下的古法造纸，与纸浆压花结合，转化为可体验、可带走的文创。'],
        ['n4', '红色文化', '第一层', '红色', '苟坝会议、马灯精神、红军过境故事，构成研学和乡村旅游的资源底色。'],
        ['n5', '村党组织·村集体', '第二层', '双链', '“党总支＋合作社＋农户”的统筹核心：土地整合、群众动员、产业组织、利益协调。'],
        ['n6', '合作社', '第二层', '农业', '遵义绿动九丰蔬菜种植专业合作社：组织生产、连接农户、提供农业社会化服务。'],
        ['n7', '农业企业（九丰）', '第二层', '农业', '山东寿光九丰农业有限公司：导入技术、品种、管理方式与市场经验。'],
        ['n8', '非遗传承人·工坊', '第二层', '非遗', '土陶传承人母先才、“花茂人家”张胜迪及返乡人才等，传承技艺并参与产品创新。'],
        ['n9', '文旅经营者·农家乐', '第二层', '非遗', '红色之家等农家乐与民宿经营者，提供餐饮接待、故事讲述与住宿服务。'],
        ['n10', '土地流转·规模种植', '第三层', '农业', '高标准农田建设与宜机化改造后，蔬菜、脆红李、“稻＋”、烤烟等规模种植。'],
        ['n11', '土陶制作·烧制', '第三层', '非遗', '从塑形到入窑的完整制陶过程；柴窑改电窑是生产方式调整的典型细节。'],
        ['n12', '纸浆压花制作', '第三层', '非遗', '把传统工序转化为可观看、可参与的手作过程。'],
        ['n13', '红色参观·研学', '第三层', '红色', '苟坝会议纪念馆及遗址参观，延伸到乡村餐饮与交流场景。'],
        ['n14', '农事体验·认养', '第三层', '农业', '“我在花茂有块田”：犁地、播种、采摘，让游客从观看田园转向参与农事。'],
        ['n15', '农产品', '第四层', '农业', '露天蔬菜、脆红李、水稻“稻＋”、烤烟等多元农产品；2024 年蔬菜 1400 亩、“稻＋”1214 亩、烤烟 900 亩。'],
        ['n16', '陶艺产品·体验', '第四层', '非遗', '游客既能看成品，也能亲手拉坯“玩泥巴”。'],
        ['n17', '文创（压花画等）', '第四层', '非遗', '压花画、小夜灯、冰箱贴——“能带得走的乡愁”。'],
        ['n18', '特色餐饮·民宿', '第四层', '非遗', '盬子鸡、八仙桌、乡村故事：餐饮不仅是消费，也承载文化体验。'],
        ['n19', '研学课程·认养田', '第四层', '农业', '农事体验兼具劳动教育与自然教育属性，接入学研学路线。'],
        ['n20', '游客·研学团队', '第五层', '双链', '2025 年接待游客 71.25 万人次；研学团队是重要的体验与消费群体。'],
        ['n21', '农产品销售·订单', '第五层', '农业', '订单农业与本地市场直供；合作社向农户和周边村镇扩散技术、市场与信息。'],
        ['n22', '电商·线下特色店', '第五层', '非遗', '文创产品与特色农产品通过线下店铺和线上渠道触达游客。'],
        ['n23', '村民增收·集体经济', '第六层', '双链', '土地租金＋务工工资＋入股分红＋集体收益，“一地多收益”。'],
        ['n24', '文化传承·人才回流', '第六层', '双链', '非遗从“保护对象”变为“发展资源”，返乡大学生参与产品创新。'],
        ['n25', '问题与约束', '第六层', '问题', '住宿消费不足、非遗手工产能有限、青年人才不足、产业数据分散、公共服务有短板。'],
    ],
    [1.5, 3.0, 1.9, 1.5, 8.1],
    aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
    caption='表4  节点清单',
)
doc.add_heading('（三）关系连线清单（30 条）', level=2)
make_table(
    ['起点', '关系词', '终点', '起点', '关系词', '终点'],
    [
        ['村党组织·村集体', '领办组建', '合作社', '土陶制作·烧制', '烧制成品', '陶艺产品·体验'],
        ['村党组织·村集体', '引进合作', '农业企业（九丰）', '纸浆压花制作', '制作文创', '文创（压花画等）'],
        ['村党组织·村集体', '统筹推动', '土地流转·规模种植', '文旅经营者·农家乐', '经营接待', '特色餐饮·民宿'],
        ['合作社', '组织生产', '土地流转·规模种植', '红色文化', '承载记忆', '红色参观·研学'],
        ['农业企业（九丰）', '注入技术', '土地流转·规模种植', '红色参观·研学', '导入客流', '特色餐饮·民宿'],
        ['土地流转·规模种植', '生产供应', '农产品', '土地流转·规模种植', '土地租金', '村民增收·集体经济'],
        ['农产品', '对接订单', '农产品销售·订单', '农产品', '务工工资', '村民增收·集体经济'],
        ['农业企业（九丰）', '打造项目', '农事体验·认养', '特色餐饮·民宿', '增加就业', '村民增收·集体经济'],
        ['农事体验·认养', '提供体验', '研学课程·认养田', '非遗传承人·工坊', '吸引返乡', '文化传承·人才回流'],
        ['游客·研学团队', '参与研学', '研学课程·认养田', '纸浆压花制作', '进入渠道', '电商·线下特色店'],
        ['游客·研学团队', '到店消费', '特色餐饮·民宿', '电商·线下特色店', '触达游客', '游客·研学团队'],
        ['游客·研学团队', '体验购买', '陶艺产品·体验', '问题与约束', '制约·住宿消费不足', '特色餐饮·民宿'],
        ['游客·研学团队', '购买文创', '文创（压花画等）', '问题与约束', '制约·手工产能有限', '非遗传承人·工坊'],
        ['非遗传承人·工坊', '传承技艺', '土陶制作·烧制', '问题与约束', '制约·当日游为主', '游客·研学团队'],
        ['非遗传承人·工坊', '传承创新', '纸浆压花制作', '问题与约束', '制约·青年人才不足', '文化传承·人才回流'],
    ],
    [3.3, 3.0, 3.4, 3.3, 3.0, 3.4],
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
    caption='表5  关系连线清单（30 条，两栏排列）',
)
doc.add_heading('（四）交互说明', level=2)
para('点击任意节点，其上下游脉络会以所属链条颜色点亮、无关节点变暗，并显示关系词；再次点击或点“重置视图”恢复。问题节点（n25）以虚线描边；节点旁色点表示所属链条（农业链 #2E6B4F、非遗链 #B9814A、红色链 #9E2B25、双链 #8A7A5C、问题 #B3542C）。')

# ================= 四、农业路径图 =================
doc.add_heading('四、农业路径图：从“小田”到多功能农业', level=1)
para('土地碎片化、撂荒、小农分散经营是起点问题；经土地整治、集中流转与宜机化改造重构生产空间，由党组织、村集体、合作社与农业企业重新组织主体，最终走向生产＋就业＋体验＋旅游的多功能产业体系。')
doc.add_heading('（一）农业主链', level=2)
para('土地碎片化·撂荒 → 土地整治·集中流转·宜机化 → 党组织＋村集体＋合作社＋企业 → 规模化·订单化·技术化·社会化服务 → 租金＋工资＋分红＋集体收益 → 生产＋就业＋体验＋旅游＋文化')
doc.add_heading('（二）五个阶段详情', level=2)
doc.add_heading('1. 土地资源重构', level=3)
para('党组织主导，按“依法、自愿”原则推进土地经营权流转；村民小组协商定价，村集体股份经济联合社组织引进经营主体；同步推进高标准农田建设与宜机化改造，小块土地整合为连片生产空间。主体：党组织、村民小组、村集体联合社、农户、经营主体。指标：有效可耕面积 +5%，机械使用效率 ×7，每亩节约人工约 700 元，亩均产值约 5000 元。')
doc.add_heading('2. 组织主体重构', level=3)
para('分散农户 → 党组织统筹 → 村集体＋合作社＋农业企业，形成“党总支＋合作社＋农户”网络：合作社组织生产、提供社会化服务，山东寿光九丰农业导入技术、品种与管理经验，农户从单一生产者变为土地提供者、入社成员与产业工人。主体：村党组织、绿动九丰合作社、九丰农业、农户、村集体。指标：1000 户扶贫户加入合作社，50 户长期稳定就业。')
doc.add_heading('3. 农业生产方式升级', level=3)
para('规模化＋订单化＋技术化＋社会化服务（飞防、机耕、植保、育苗），发展露天蔬菜、脆红李、水稻“稻＋”综合种养、烤烟等多元农业，并向周边村镇农户扩散技术、市场与信息。主体：合作社、农业企业、农户、技术人员。指标：3000 亩露天蔬菜，1450 亩脆红李；2024 年蔬菜 1400 亩、“稻＋”1214 亩、烤烟 900 亩。')
doc.add_heading('4. 利益联结：一地多收益', level=3)
para('土地流转获租金、进企务工获工资、加入合作社获分红、集体经营获收益；增值收益按比例分配，让土地变成资产、农民成为股东。主体：农户、合作社、企业、村集体。指标：流转增值收益农户 70%·集体 30%，分红 170 万元，6 年发放工资近 1000 万元。')
doc.add_heading('5. 农业多功能化', level=3)
para('农田从单纯生产空间，转变为农产品生产、就业、景观、体验、旅游消费一体的特色场景；与陶瓷烧制、纸浆压花、红色文化、农家乐民宿结合，形成农文旅相互支撑的产业体系。主体：农户、村集体、企业、游客、文旅经营主体。典型场景：“我在花茂有块田”、农耕体验、农家乐。')
doc.add_heading('（三）路径结果', level=2)
para('农业产业持续发展、村民增收、集体经济增强、乡村活力提升。')
doc.add_heading('（四）两个案例说明', level=2)
rich([('案例一 · 土地资源重构（核心问题 → 核心做法 → 直接成效）', True, False, CN_HEAD)], indent=False)
para('山区土地分散、小块化制约机械化与规模经营。花茂村以“先集中、后流转、再付费”的制度方法整合土地，有效可耕面积增加 5%、机械效率提升 7 倍——土地流转不是简单“把地租出去”，而是制度、组织与基础设施的协同重构。')
rich([('案例二 · 农业多功能化（功能转换）', True, False, CN_HEAD)], indent=False)
para('农田 → 生产空间；农田＋游客 → 农事体验空间；农业＋文化 → 地方特色体验；农业＋旅游 → 餐饮、民宿与休闲消费。“我在花茂有块田”让游客从观看田园转向参与农事，实现农业价值增值。')
para('数据与阶段结构依据《花茂村农业发展内容设计》《花茂村实践报告》整理；指标均为材料原口径，冲突数据见“待核实数据清单”。')

# ================= 五、非遗文旅路径图 =================
doc.add_heading('五、非遗文旅路径图：从技艺到消费', level=1)
para('土陶与古法造纸两项非遗技艺，经传承人和返乡人才转化为可体验、可带走的产品；四个地点中心发散“资源—主体—活动—产品服务—渠道—结果”六类信息，由三条主线串成游览与消费路径。')
doc.add_heading('（一）非遗文旅主链', level=2)
para('土陶·古法造纸技艺（资源基础）→ 非遗传承人·返乡人才（组织主体）→ 手工制作·烧制（产业活动）→ 产品与体验·陶艺文创餐饮（产品与服务）→ 研学团队·游客（市场渠道）→ 销售·就业·文化传承（乡村结果）')
doc.add_heading('（二）四个地点中心（六类信息）', level=2)
make_table(
    ['地点', '资源', '主体', '活动', '产品服务', '渠道', '结果/问题'],
    [
        ['花茂村陶艺馆（土陶非遗中心点）', '土陶技艺（省级非遗，传承千年）', '传承人/陶艺经营者（母先才等）', '制陶与烧制：从塑形到入窑', '陶艺产品/动手拉坯体验', '游客/研学团队参观体验', '文化传承；生产方式调整（柴窑→电窑）'],
        ['古法造纸及纸浆压花体验点（造纸非遗中心点）', '古法造纸技艺', '非遗传承人/返乡人才（“花茂人家”张胜迪等）', '造纸/纸浆压花制作', '压花画/小夜灯/冰箱贴/手作体验', '游客/研学团队/线下销售', '文化传承；人才回流；手工产能有限（小而精）'],
        ['红色之家（餐饮文旅中心点）', '红色文化/乡村故事/地方饮食', '农家乐经营者/当地村民', '就餐/交流/听红军故事', '盬子鸡等特色餐饮/文化讲述', '红色参观游客/研学团队', '村民经营；延长消费链；住宿不足'],
        ['我在花茂有块田认养区（农事体验中心点）', '农田/乡村景观', '农户/项目运营者（九丰农业科技）', '犁地/播种/采摘', '农田认养/农事体验/研学', '游客/学生/研学团队', '农业传播；农旅融合；运营信息待核实'],
    ],
    [2.6, 2.3, 2.5, 2.2, 2.5, 2.2, 2.7],
    aligns=[WD_ALIGN_PARAGRAPH.LEFT]*7,
    caption='表6  四个地点中心六类信息',
)
doc.add_heading('（三）地点之间的三条主线（关系词连线）', level=2)
para('苟坝会议纪念馆及遗址 → 导入红色游客 → 红色之家（参观之后进入乡村餐饮和交流场景）')
para('红色之家 → 串联乡村游线 → 陶艺馆/造纸体验点（从餐饮接待延伸到非遗体验）')
para('陶艺馆/造纸体验点 → 组合研学内容 → 我在花茂有块田（从文化手作延伸到农事参与）')
doc.add_heading('（四）虚线提示', level=2)
para('四个地点共同指向“游客停留与综合消费”——目前部分游客以当日游为主，住宿消费不足，消费链仍有延伸空间。')
para('结构与节点文本依据《花茂村非遗文旅内容设计文档》；陶艺馆、造纸体验点等正式全称与经营数据待核实。')

# ================= 六、四张经营主体案例卡 =================
doc.add_heading('六、四张经营主体案例卡', level=1)
para('案例卡的选择原则是“材料完整，而不是知名度最高”。每张卡只保留材料能支持的信息，经营数据未经核实的以“待核实”标注，不把单个主体的收入、规模或模式扩大为全村结论。案例卡图片暂以空白占位，现场照片后续补充。')
def case_block(title, quote, field, chain, verify):
    doc.add_heading(title, level=2)
    rich([('引语：', True, False, CN_HEAD), (quote, False, True, CN_BODY)])
    rich([('实践现场：', True, False, CN_HEAD), (field, False, False, CN_BODY)])
    rich([('产业关系：', True, False, CN_HEAD), (chain, False, False, CN_BODY)])
    rich([('待核实：', True, False, CN_HEAD), (verify, False, False, CN_BODY)])

case_block('（一）案例一 · 非遗文旅（花茂村陶艺馆）',
           '从一团泥到一件器物，土陶技艺在制作、烧制和游客体验中继续流动。',
           '访谈中，彭龙芬讲到陶艺馆烧窑方式的变化：过去烧坛罐主要靠柴、一烧就是多天，后来改用电窑。这个关于“窑火”的细节，比单写“保护非遗”更能说明传统技艺怎样适应新的生产和生态要求。',
           '土陶技艺 → 传承人和经营者 → 制作烧制 → 陶艺产品与体验 → 游客研学 → 文化传承',
           '陶艺馆正式全称、传承人姓名和经营数据。')
case_block('（二）案例二 · 非遗文旅（古法造纸与纸浆压花）',
           '一张纸不只被展示，还被做成压花画、小夜灯和冰箱贴，进入游客可以体验和带走的日常物件。',
           '团队围绕手工造纸与传承人交流，并观察纸浆压花相关产品。传统技艺在这里不是静态陈列，而是经过返乡人才的创意转化，变成可以制作、使用和购买的文创；同时，纯手工也决定了它难以简单追求产量。',
           '古法造纸 → 传承人与返乡人才 → 纸浆压花制作 → 文创产品与体验 → 游客研学 → 文化传承和人才回流',
           '体验点、经营主体、传承人正式名称和销售渠道。')
case_block('（三）案例三 · 红色文旅（红色之家）',
           '一张八仙桌把地方饮食、乡村生活和红色故事连接在一起。',
           '团队走访苟坝会议纪念馆和会议遗址后，再看花茂村的文旅场景，能够更清楚地理解红色文化怎样从纪念空间延伸到村民经营：游客围坐在八仙桌前，在盬子鸡的香气里听村民讲红军过境时的军民故事。',
           '红色文化与乡村故事 → 农家乐经营者 → 特色餐饮与讲述 → 游客消费 → 村民经营',
           '经营者、成立时间、客流和收入数据。')
case_block('（四）案例四 · 农业文旅（我在花茂有块田认养区）',
           '游客从田埂边的观看者，变成犁地、播种和采摘的参与者。',
           '团队在实地考察中关注到这一认养区。泥土、农具和作物本身就是体验内容：游客亲手参与农事，能够理解农业生产过程，也让普通农田获得研学和旅游属性。该场景适合补用团队在田间观察或参与活动的照片。',
           '农田和乡村景观 → 农户与运营者 → 认养、犁地、播种、采摘 → 游客学生 → 农业体验和乡村传播',
           '运营主体、当前状态、收费方式、参与人数和收益分配。')
para('案例文本依据《花茂村非遗文旅内容设计文档》与《花茂村实践报告》。')

# ================= 七、数据与反馈 =================
doc.add_heading('七、数据与反馈', level=1)
doc.add_heading('（一）待核实数据清单', level=2)
para('以下数据均见于公开报道，因口径或单位冲突暂不进入核心图表；出处在右侧“公开报道出处”列，点击“原文链接”可核对。', indent=False)
para('表7  待核实数据清单（出处含原文链接）', style='Caption', indent=False, size=10.5)
t7 = doc.add_table(rows=1, cols=3)
t7.style = 'Table Grid'
t7.alignment = WD_TABLE_ALIGNMENT.CENTER
t7.autofit = False
hdr7 = ['待核实内容', '问题说明', '公开报道出处']
for i, h in enumerate(hdr7):
    c = t7.rows[0].cells[i]
    cell_text(c, h, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade(c, GRAY_HDR)
GZRB_PDF = 'http://szb.eyesnews.cn/pc/att/202403/08/a8c36c8a-1ae5-4a74-8df7-2055bd2be45d.pdf'   # 贵州日报 2024.3
ZY_GOV = 'https://www.zunyi.gov.cn/zwgk/zdlygk/mzzj/mzwhbh/202501/t20250123_86670687.html'      # 遵义市政府 2025.1
GZ_GOV = 'https://www.guizhou.gov.cn/home/gzyw/202504/t20250417_87533552.html'                   # 贵州省政府 2025.4
RMW_10948 = 'http://cpc.people.com.cn/n1/2017/0929/c412690-29566833.html'                        # 人民网 2017.9
GZDJY = 'https://gzzzb.gov.cn/ywzx/djqs/20230619/20230619_245664.shtml'                           # 贵州党建云 2023.6
CCTV_4950 = 'https://news.cctv.com/2019/07/27/ARTIwHXn7OfmL0zGROCaAkml190727.shtml'              # 央视网 2019.7
rows7 = [
    ('2023 年接待游客 118 万人次、旅游收入 3000 多元',
     '数量级明显不匹配，可能缺少“万元”等单位，不进入核心图表',
     '贵州日报 2024.3 专访彭龙芬（原文为“3000 多万元”）',
     [GZRB_PDF, ZY_GOV]),
    ('村集体经济积累从 26 万元“增加值”1402 万元',
     '起点、终点与增加值关系不清，不计算比例和增速',
     '贵州省人民政府网 2025.4（“集体收入累计从 26 万元增至 1402 万元”）',
     [GZ_GOV]),
    ('2014 年人均收入 10948 元 / 10984 元',
     '两种公开说法并存，不用于核心结论',
     '10948 元：人民网 2017.9；10984 元：贵州党建云 2023.6',
     [RMW_10948, GZDJY]),
    ('人口约 5166 人 / 2021 年 4950 人',
     '两个口径并存，不计算人均值',
     '5166 人：贵州日报天眼新闻 2024.12；4950 人：央视网 2019.7',
     [ZY_GOV, CCTV_4950]),
]
for a, b, m, urls in rows7:
    cells = t7.add_row().cells
    cell_text(cells[0], a, align=WD_ALIGN_PARAGRAPH.LEFT)
    cell_text(cells[1], b, align=WD_ALIGN_PARAGRAPH.LEFT)
    cell_source(cells[2], m, urls)
for ci, w in enumerate([4.2, 4.6, 7.2]):
    for row in t7.rows:
        row.cells[ci].width = Cm(w)
doc.add_paragraph().paragraph_format.space_after = Pt(2)
last = doc.paragraphs[-1]
last.paragraph_format.first_line_indent = Pt(0)
last.paragraph_format.line_spacing = 1.0
last.paragraph_format.space_before = Pt(0)
doc.add_heading('（二）调研反馈摘要', level=2)
para('反馈摘要采用“调研发现—报告依据—建议方向—待核实事项”四段式；所有建议均为可讨论方向，不作为已实施成果。', indent=False)
def fb(no, title, find, basis, advice, verify):
    doc.add_heading(str(no) + '. ' + title, level=3)
    rich([('调研发现：', True, False, CN_HEAD), (find, False, False, CN_BODY)], indent=False)
    rich([('报告依据：', True, False, CN_HEAD), (basis, False, False, CN_BODY)], indent=False)
    rich([('建议方向（可讨论）：', True, False, CN_HEAD), (advice, False, False, CN_BODY)], indent=False)
    rich([('待核实事项：', True, False, CN_HEAD), (verify, False, False, CN_BODY)], indent=False)

fb(1, '游客以当日游览为主，住宿消费不足',
   '游客更多以当天游览为主，住宿消费不足，旅游产业链仍有延伸空间。',
   '《花茂村实践报告》“产业发展持续升级压力”部分。',
   '将红色参观、非遗体验、农事体验与夜间住宿组合为半日、一日和两日产品。',
   '游客过夜率、住宿设施入住率等数据。')
fb(2, '非遗产品有特色，但纯手工产能有限',
   '非遗文创有特色，但受纯手工制作限制，生产规模有限，产品走“小而精、小而美”路线。',
   '传承人表述及《花茂村实践报告》非遗部分。',
   '区分体验产品、收藏产品与可标准化文创；不以简单扩产代替技艺保护。',
   '各产品线产能、成本与销量数据。')
fb(3, '青年人才参与仍不足',
   '非遗领域已出现返乡大学生参与创新，但整体青年人才仍不足。',
   '《花茂村实践报告》“青年人才参与乡村治理不足”部分。',
   '建立高校长期合作，让学生承担内容制作、产品设计和线上传播等可远程任务。',
   '返乡青年数量与就业结构。')
fb(4, '产业数据分散、更新口径不一',
   '游客人次、旅游收入、集体积累等数据年份与单位口径不一，存在冲突。',
   '公开报道与调研数据比对，发现数量级不匹配与口径冲突。',
   '建立年度指标清单，明确统计人、指标定义、数据年份与更新时间。',
   '见本页“待核实数据清单”。')
doc.add_heading('（三）公共服务约束', level=2)
para('医疗养老等公共服务仍有短板（村卫生室以基础设备为主，村医待遇与人才接续问题突出）。图谱可注明这一约束，但不能把产业增长直接等同于公共服务改善。')

# ================= 八、资料来源与项目边界 =================
doc.add_heading('八、资料来源与项目边界', level=1)
doc.add_heading('（一）资料来源清单', level=2)
para('每条史实、数字与关系均可回溯到实践报告、访谈记录或公开报道；存在冲突的数据不进入核心图表。', indent=False)
sources = [
    ('《花茂村实践报告》——上海交通大学集成电路学院红芯计划暑期社会实践团调研报告（调研主线）', None),
    ('访谈记录：《花茂村基层干部采访整合记录》《花茂村纸浆压花画非遗传承人访谈整合记录》（案例实践现场依据）', None),
    ('人民网贵州频道《花茂村：一个记得住乡愁的地方》(2023.7)', 'http://gz.people.com.cn/BIG5/n2/2023/0712/c194849-40490988.html'),
    ('新京报《播州区枫香镇花茂村：笑容在“乡愁”中绽放》(2025.8)', 'https://m.bjnews.com.cn/detail/175557519019420.html'),
    ('贵州省人民政府《遵义市花茂村农文旅融合描绘美丽画卷》(2024.6)', 'https://www.guizhou.gov.cn/ztzl/lycyh/sdxd/202406/t20240616_84876007.html'),
    ('国际在线《贵州遵义花茂村：打造花繁叶茂和美乡村新样板》(2025.6)', 'https://gz.cri.cn/2025-06-18/20546f75-2754-9e74-dbb8-6c42f3c23bee.html'),
    ('央视网《走进乡村看小康·古法造纸》(2021.8)', 'https://news.cctv.com/2021/08/28/ARTITgGYuBLyIr6eji3f8c7A210828.shtml'),
    ('中新网《花茂造纸》(2021.5)', 'https://www.chinanews.com.cn/cul/2021/05-17/9479115.shtml'),
    ('高德地图 POI 数据（点位坐标）；“花茂人家”坐标 106.577666, 27.617247；苟坝会议点位坐标来自前期项目《苟坝花茂红色文化数字地图》实测核验', None),
    ('人民网党建频道《党建调研行｜以业态之“新”赋能共富之“兴”》(2026.7)：2025 年游客 71.25 万人次、旅游收入突破 2400 万元、集体经济积累 1481 万元、人均可支配收入 29650 元', 'http://dangjian.people.com.cn/BIG5/n1/2026/0707/c448938-40755191.html'),
    ('新华网贵州频道《红色热土孕育新生活》(2026.9)：2025 年集体经济积累 1481 万元、人均可支配收入 29650 元', 'http://gz.news.cn/20260911/410f2875a4b04f4b8a782797b1041bca/c.html'),
    ('贵州省人民政府网《总书记来过之后：花茂村的笑容越来越灿烂》(2025.4)：集体收入累计从 26 万元增至 1402 万元', 'https://www.guizhou.gov.cn/home/gzyw/202504/t20250417_87533552.html'),
    ('遵义市人民政府网《花茂村：总书记话儿记心间，老百姓致富笑开颜》(2025.1)：2023 年接待游客 118 万余人次、旅游总收入 3437 万元', 'https://www.zunyi.gov.cn/zwgk/zdlygk/mzzj/mzwhbh/202501/t20250123_86670687.html'),
    ('贵州日报《花茂村十年发展小记》(2025.6)：面积 9.8 平方公里、26 个村民组 1345 户、户籍人口 5166 人', 'http://zyrb.zunyiol.cn/resfile/2025-06-16/06/zyrbszb-20250616-006.pdf'),
    ('贵州党建云《向总书记报告：花茂的笑脸越来越多了！》(2023.6)：人均可支配收入从 2014 年 10984 元增长至 2022 年 23163 元', 'https://gzzzb.gov.cn/ywzx/djqs/20230619/20230619_245664.shtml'),
    ('央视网《贵州花茂村：“农旅文一体化”打造美丽新农村》(2019.7)：1345 户 4950 人', 'https://news.cctv.com/2019/07/27/ARTIwHXn7OfmL0zGROCaAkml190727.shtml'),
    ('贵州日报（天眼新闻）《老区展新颜｜在花茂村 过向往的生活》(2024.12)：户籍人口 5166 人、1345 户；2023 年游客 118 万人次、旅游收入 3437 万元', 'http://m.toutiao.com/group/7446333820233597467/'),
]
for i, (txt, url) in enumerate(sources, 1):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run('[' + str(i) + '] ' + txt)
    set_run(r, size=12)
    if url:
        add_hyperlink(p, url, '（原文链接）', size=12)

doc.add_heading('（二）项目边界（三条红线）', level=2)
para('1. 实践报告是叙事主线，公开资料只用于补充和交叉核验。', indent=False)
para('2. 不同年份和不同统计口径的数据不能拼成同一时间点。', indent=False)
para('3. 个别经营主体的情况不能扩大为全村结论。', indent=False)
doc.add_heading('（三）页面说明', level=2)
para('本整理文档对应“花茂村产业图谱”项目网页（单文件静态网页），集成高德地图点位、六层产业关系总图、农业路径图、非遗文旅路径图、四张案例卡、数据与反馈与来源说明；页面图片暂以空白占位，现场照片后续补充。')

doc.save(OUT)
print('saved:', OUT)
