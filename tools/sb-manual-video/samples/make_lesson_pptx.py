"""매뉴얼 영상용 예시 PPTX — 클릭할 때마다 항목이 하나씩 나타나는 애니메이션 포함."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree

prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'

def text(slide, x, y, w, h, s, size, color='1F2937', bold=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    p = tb.text_frame.paragraphs[0]; p.text = s
    r = p.runs[0]; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = RGBColor.from_string(color)
    return tb

def card(slide, x, y, w, h, s, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid(); shp.fill.fore_color.rgb = RGBColor.from_string(fill); shp.line.fill.background()
    tf = shp.text_frame; tf.text = s
    r = tf.paragraphs[0].runs[0]; r.font.size = Pt(28); r.font.bold = True; r.font.color.rgb = RGBColor(255, 255, 255)
    return shp

def click_appear(slide, shapes):
    """도형마다 '클릭하면 나타나기' 애니메이션을 붙입니다."""
    ids = [s.shape_id for s in shapes]
    cid = [2]
    def nid():
        cid[0] += 1; return cid[0]
    pars = ''
    for spid in ids:
        a, b, c, d = nid(), nid(), nid(), nid()
        pars += (f'<p:par><p:cTn id="{a}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
                 f'<p:par><p:cTn id="{b}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                 f'<p:par><p:cTn id="{c}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">'
                 f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr>'
                 f'<p:cTn id="{d}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                 f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
                 f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par>'
                 f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    xml = (f'<p:timing xmlns:p="{NS}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
           f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>{pars}</p:childTnLst></p:cTn>'
           f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
           f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
           f'</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>')
    slide.element.append(etree.fromstring(xml))

s = prs.slides.add_slide(BLANK)
text(s, 1, 2.6, 11.3, 1.2, '광합성', 60, '065F46', True)
text(s, 1, 3.9, 11.3, 0.8, '2학년 과학 · 식물은 어떻게 양분을 만들까?', 28, '374151')

s = prs.slides.add_slide(BLANK)
text(s, 0.8, 1.3, 11.7, 0.9, '광합성에 필요한 것', 40, '065F46', True)
c1 = card(s, 0.8, 2.8, 3.7, 2.2, '☀️ 빛 에너지', 'F59E0B')
c2 = card(s, 4.8, 2.8, 3.7, 2.2, '🌬️ 이산화 탄소', '0EA5E9')
c3 = card(s, 8.8, 2.8, 3.7, 2.2, '💧 물', '2563EB')
click_appear(s, [c1, c2, c3])

s = prs.slides.add_slide(BLANK)
text(s, 0.8, 1.3, 11.7, 0.9, '광합성으로 만들어지는 것', 40, '065F46', True)
d1 = card(s, 0.8, 2.8, 5.6, 1.6, '포도당 → 녹말로 저장', '16A34A')
d2 = card(s, 6.9, 2.8, 5.6, 1.6, '산소 → 기공으로 배출', '0D9488')
q = text(s, 0.8, 5.0, 11.7, 0.9, '오늘의 질문: 밤에도 광합성이 일어날까?', 30, 'B91C1C', True)
click_appear(s, [d1, d2, q])

prs.save('lesson.pptx')
print('ok')
