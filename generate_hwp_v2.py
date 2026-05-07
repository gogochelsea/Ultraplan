# -*- coding: utf-8 -*-
"""
사업 계획서 수정본.hwp 자동 생성 스크립트
- 수정 사항:
  1) 재활용·건설폐기물 저감 주장 전면 제거
  2) 사업 규모 전반 축소 (5년차 매출 250억 → 50억, 인력 72명 → 24명 등)
  3) 표 5·6에 대한 산정 근거 단락 추가
  4) 경쟁사·경쟁기술 분석 및 표 2 삭제
  5) 표 4 단순화 + 인력 산정 근거 추가, 본문 전반 근거 보강
  6) 모든 표를 실제 HWP 표(TableCreate)로 재구성
"""
import os
import win32com.client as win32

BASE = r"C:\Users\filk2\Desktop\공부자료\학교수업\화공산"
IMG = os.path.join(BASE, "extracted_images")
OUT = os.path.join(BASE, "사업 계획서 수정본.hwp")

IMG_ORG_CHART     = os.path.join(IMG, "BIN0003.png")
IMG_AERIAL        = os.path.join(IMG, "BIN0004.png")
IMG_NONFLAM_DIA   = os.path.join(IMG, "BIN0001.jpg")
IMG_NONFLAM_PHOTO = os.path.join(IMG, "BIN0002.JPG")
IMG_CONE_PHOTO    = os.path.join(IMG, "BIN0005.JPG")
IMG_CONE_DIA      = os.path.join(IMG, "BIN0006.jpg")
IMG_GAS_PHOTO     = os.path.join(IMG, "BIN0007.JPG")
IMG_FACADE_RIG    = os.path.join(IMG, "BIN0008.png")
IMG_FACADE_FIRE   = os.path.join(IMG, "BIN0009.png")
IMG_GAS_DIA       = os.path.join(IMG, "BIN000A.jpg")
IMG_CABLE_TEST    = os.path.join(IMG, "BIN000C.png")
IMG_CABLE_BUNDLE  = os.path.join(IMG, "BIN000D.png")
IMG_PEN_FIRE      = os.path.join(IMG, "BIN000E.png")

# COM 초기화
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
try:
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
except Exception:
    pass
hwp.XHwpWindows.Item(0).Visible = False
hwp.SetMessageBoxMode(0x2FF1)
hwp.HAction.Run("FileNew")

# ──────────────────────────────────────────────────────────────────────────
# 헬퍼
# ──────────────────────────────────────────────────────────────────────────
def set_char_shape(size_pt=12, bold=False, italic=False, color=0, face="함초롬바탕"):
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    cs = hwp.HParameterSet.HCharShape
    cs.Height = int(size_pt * 100)
    cs.Bold = 1 if bold else 0
    cs.Italic = 1 if italic else 0
    cs.TextColor = color
    for ln in ["Hangul", "Latin", "Hanja", "Japanese", "Other", "Symbol", "User"]:
        setattr(cs, f"FaceName{ln}", face)
        setattr(cs, f"FontType{ln}", 1)
        setattr(cs, f"Offset{ln}", 0)
        setattr(cs, f"Ratio{ln}", 100)
        setattr(cs, f"Spacing{ln}", 0)
        setattr(cs, f"Size{ln}", 100)
    hwp.HAction.Execute("CharShape", cs.HSet)

def set_para_shape(align="left", line_space=160, space_before=0, space_after=0):
    hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
    ps = hwp.HParameterSet.HParaShape
    align_map = {"left": 0, "center": 1, "right": 2, "justify": 3, "distribute": 4}
    ps.AlignType = align_map.get(align, 0)
    ps.LineSpacing = line_space
    ps.LineSpacingType = 0
    ps.PrevSpacing = space_before
    ps.NextSpacing = space_after
    hwp.HAction.Execute("ParagraphShape", ps.HSet)

def insert_text(text):
    if not text:
        return
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

def br():
    hwp.HAction.Run("BreakPara")

def page_break():
    hwp.HAction.Run("BreakPage")

def write_line(text, size=12, bold=False, align="left", color=0):
    set_char_shape(size_pt=size, bold=bold, color=color)
    set_para_shape(align=align)
    insert_text(text)
    br()

def write_para(text, size=12, bold=False, align="justify"):
    set_char_shape(size_pt=size, bold=bold)
    set_para_shape(align=align)
    insert_text(text)
    br()

def insert_picture(path, width_mm=120, height_mm=None, caption=None):
    if not os.path.exists(path):
        print(f"[WARN] image not found: {path}")
        return
    try:
        from PIL import Image
        with Image.open(path) as im:
            iw, ih = im.size
        aspect = ih / iw
    except Exception:
        aspect = 0.7
    if height_mm is None:
        height_mm = width_mm * aspect
        if height_mm > 130:
            height_mm = 130
            width_mm = height_mm / aspect

    set_para_shape(align="center")
    set_char_shape(size_pt=12)

    HWPMM = 283.464567
    w_unit = int(width_mm * HWPMM)
    h_unit = int(height_mm * HWPMM)

    try:
        hwp.InsertPicture(path, True, 2, False)
    except Exception:
        try:
            hwp.InsertPicture(path, True, 0, False)
        except Exception as e2:
            print("InsertPicture err:", e2)

    try:
        last = hwp.HeadCtrl
        target = None
        while last:
            if last.CtrlID == "gso":
                target = last
            last = last.Next
        if target:
            prop = target.Properties
            prop.SetItem("Width",  w_unit)
            prop.SetItem("Height", h_unit)
            target.Properties = prop
    except Exception as e:
        print("Resize err:", e)

    br()
    if caption:
        set_char_shape(size_pt=11, bold=False)
        set_para_shape(align="center")
        insert_text(caption)
        br()
    set_para_shape(align="left")

def make_table(headers, rows, col_widths_mm, body_align="center", row_h_mm=12):
    """실제 HWP 표 생성 - 헤더 1행 + 본문 N행
    - 매 호출마다 hwp.CreateAction/CreateSet으로 독립 파라미터셋 생성
      (HParameterSet 싱글턴 재사용 시 두 번째 TableCreate부터 COM 오류 발생)
    - 행 높이는 내용에 맞춰 자동 확장 (HeightType=0)
    - 표 빠져나오기는 SetPos(부모 단락) + MoveLineEnd + BreakPara
    """
    cols = len(headers)
    nrows = len(rows) + 1
    total_w = sum(col_widths_mm)

    set_para_shape(align="center")
    set_char_shape(size_pt=10)

    saved_pos = hwp.GetPos()

    # 신규 액션·파라미터셋 (싱글턴 재사용 회피)
    act = hwp.CreateAction("TableCreate")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("Rows", nrows)
    pset.SetItem("Cols", cols)
    pset.SetItem("WidthType", 2)
    pset.SetItem("HeightType", 0)
    pset.SetItem("WidthValue", hwp.MiliToHwpUnit(total_w))
    pset.SetItem("HeightValue", hwp.MiliToHwpUnit(row_h_mm * nrows))
    arr_w = pset.CreateItemArray("ColWidth", cols)
    for i, w in enumerate(col_widths_mm):
        arr_w.SetItem(i, hwp.MiliToHwpUnit(w))
    arr_h = pset.CreateItemArray("RowHeight", nrows)
    for i in range(nrows):
        arr_h.SetItem(i, hwp.MiliToHwpUnit(row_h_mm))
    act.Execute(pset)

    # 셀 채우기
    all_data = [headers] + rows
    for r, row in enumerate(all_data):
        for c, val in enumerate(row):
            set_char_shape(size_pt=10, bold=(r == 0))
            set_para_shape(align=("center" if r == 0 else body_align), line_space=150)
            insert_text(str(val))
            if not (r == nrows - 1 and c == cols - 1):
                hwp.HAction.Run("TableRightCell")

    # 표 빠져나오기 — 부모 단락 위치로 SetPos 후 MoveLineEnd → BreakPara
    hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
    hwp.HAction.Run("MoveLineEnd")
    set_char_shape(size_pt=12, bold=False)
    set_para_shape(align="left")
    br()

# ──────────────────────────────────────────────────────────────────────────
# 표지
# ──────────────────────────────────────────────────────────────────────────
for _ in range(4):
    write_line("", size=12)

write_line("사 업 계 획 서", size=32, bold=True, align="center")
write_line("", size=12)
write_line("", size=12)
write_line("─────────────────────────────────", size=14, align="center")
write_line("", size=12)
write_line("할로겐프리 하이브리드 난연 시스템 기반", size=20, bold=True, align="center")
write_line("친환경 외벽 단열보드 「Safe-Core™」", size=20, bold=True, align="center")
write_line("", size=12)
write_line("─────────────────────────────────", size=14, align="center")

for _ in range(8):
    write_line("", size=12)

write_line("성균관대학교 일반대학원 화학공학과", size=16, align="center")
write_line("", size=12)
write_line("대 표 자  :  박  정  빈", size=16, align="center")
write_line("", size=12)
write_line("제 출 일  :  2026. 05. 07.", size=16, align="center")

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 목차 (경쟁사·경쟁기술 분석 항목 삭제)
# ──────────────────────────────────────────────────────────────────────────
write_line("목   차", size=12, bold=True, align="center")
write_line("", size=12)

toc = [
    ("1. 사업 개요 및 기술 소개", True),
    ("    -사업 추진 배경 및 필요성", False),
    ("    -사업 비전과 미션", False),
    ("    -제품 및 기술 개요", False),
    ("    -팀 구성 요약", False),
    ("2. BM 타당성", True),
    ("    -신제품 및 서비스 추진전략과 사업화 계획", False),
    ("    -BM 구체화 및 시장 진입 전략", False),
    ("    -구체적인 비즈니스 모델", False),
    ("3. 시장 분석", True),
    ("    -국내외 산업 동향", False),
    ("    -시장 규모 및 성장성", False),
    ("    -SWOT 분석", False),
    ("4. 추진 일정 및 단계별 계획", True),
    ("    -단계별 로드맵", False),
    ("    -주요 마일스톤", False),
    ("5. 조직 구성 및 인력 운영 계획", True),
    ("    -조직도", False),
    ("    -단계별 인력 확보 계획", False),
    ("    -역량 강화 방안", False),
    ("6. 재무 계획 및 자금 조달", True),
    ("    -매출 및 손익 추정", False),
    ("    -투자 계획", False),
    ("    -자금 조달 방안", False),
    ("7. 리스크 분석 및 대응 방안", True),
    ("8. 기대 효과 및 결론", True),
]
for line, bold in toc:
    write_line(line, size=12, bold=bold, align="left")

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 1장. 사업 개요 및 기술 소개
# ──────────────────────────────────────────────────────────────────────────
write_line("1. 사업 개요 및 기술 소개", size=12, bold=True)
write_line("", size=12)

write_line("-사업 추진 배경 및 필요성", size=12, bold=True)
write_para(
    "최근 10년간 의정부 도시형생활주택 화재(2015), 제천 스포츠센터 화재(2017), 울산 주상복합 화재(2020) 등 외벽 가연성 단열재로 인한 대형 인명피해가 반복적으로 발생하였다. "
    "소방청 화재통계연감(2022)에 따르면 2017~2021년 5년간 외벽 마감재가 화재 확대 원인으로 분류된 건수는 연평균 220건 이상이며, 이로 인한 사망자 수는 누적 70명을 상회한다. "
    "이들 사고의 공통 원인은 EPS(발포 폴리스티렌) 및 PU(폴리우레탄)폼이 외벽 시스템에 다층 구조로 적용되어 굴뚝효과(stack effect)에 의해 화염이 수직 방향으로 급속히 확산된 데 있다. "
    "이에 국토교통부는 「건축물 마감재료의 난연성능 및 화재 확산 방지구조 기준」(국토교통부 고시 제2022-84호) 개정을 통해 6층 이상 또는 높이 22 m 이상 건축물의 외벽 마감재료에 대해 실대형 모형 화재시험(KS F 8414)과 마감재 단위 불연·준불연 등급 동시 충족을 의무화하였다. "
    "그러나 현행 시장의 무기계 단열재(미네랄울)는 단열 성능이 낮고(λ ≈ 0.034 W/m·K), 페놀폼(PF)은 흡습 시 강도가 저하되며, 할로겐계 난연제(브롬계)는 다이옥신 방출 우려로 EU REACH·국내 화학물질평가법 규제 대상이다. "
    "본 사업은 인계·질소계·실리콘계 하이브리드 난연 시스템과 무기계 충전제(Al(OH)₃, Mg(OH)₂)를 결합한 친환경 외벽 단열보드 'Safe-Core™'를 개발하여 이러한 시장의 미충족 수요(unmet need)를 해소하고자 한다."
)
insert_picture(IMG_FACADE_FIRE, caption="<그림 1> 외벽 마감재 화재 확산 사례 (KFPA 실대형 모형 화재시험)")

write_line("-사업 비전과 미션", size=12, bold=True)
write_para(
    "본 사업의 비전은 '국내 외벽 단열재 시장의 안전·환경 패러다임 전환을 선도하는 친환경 본질안전(intrinsically safe) 소재 기업'이다. "
    "미션으로는 다음 3가지를 설정한다. 첫째, 할로겐프리·저독성 난연 처방 표준화를 통해 국내 건축물의 화재 인명피해를 감축한다(KFPA 사고통계 기반 시뮬레이션 시 약 30% 감축 가능). "
    "둘째, 단열성능 향상(λ ≤ 0.022 W/m·K)을 통해 건축물 운영단계 에너지 사용량을 절감하고 건축물 부문 탄소중립 목표 달성에 기여한다. "
    "셋째, KS F ISO 5660-1·1182·KS F 8414 인증의 동시 획득을 통해 국내 외벽 단열재 시장 내 안정적 입지를 확보한다. "
    "중장기 목표는 3년차 매출 24억 원, 5년차 매출 50억 원 및 국내 준불연 외벽 단열재 신규 수요(SOM) 시장 점유율 4% 달성이다."
)

write_line("-제품 및 기술 개요", size=12, bold=True)
write_para(
    "Safe-Core™는 페놀폼(PF) 코어를 인-질소 시너지 난연제(APP·MEL)와 실리콘계 그래프트 코폴리머로 후처리하고, 표면을 Al(OH)₃ 충전 무기계 마감층으로 이중 코팅한 다층 단열보드이다. "
    "핵심 화공 메커니즘은 ① Al(OH)₃의 흡열 분해(약 220 kJ/kg, 180~200 ℃에서 H₂O 생성)에 의한 표면 냉각, "
    "② APP(암모늄 폴리포스페이트)의 탄화층 형성에 의한 산소·열 차단, "
    "③ 멜라민의 NH₃ 방출에 의한 가연성 가스 희석, "
    "④ 실리콘 라디칼 스캐빈저에 의한 연쇄반응 차단의 4중 난연 시너지이다. "
    "산소소비원리(Oxygen Consumption Principle, Huggett 상수 13.1 kJ/g) 기반 콘칼로리미터(KS F ISO 5660-1) 사내 평가에서 최대 열방출률(pHRR) 약 220 kW/m², 총방출열량(THR_10min) 약 6.8 MJ/m² 수준으로 측정되어 준불연재료 합격 기준(pHRR ≤ 200~250 kW/m², THR ≤ 8 MJ/m²)을 만족한다. "
    "본 측정 결과는 2025년 KFPA 현장실습 기간 중 실측 데이터에 근거하며, 동일 처방의 양산 단계 재현성 확보를 위해 통계적 공정관리(SPC) 기법을 병행 도입할 계획이다."
)
insert_picture(IMG_CONE_DIA, caption="<그림 2> 콘칼로리미터 시험 원리 (KS F ISO 5660-1)")

write_line("-팀 구성 요약", size=12, bold=True)
write_para(
    "대표(박정빈, 화학공학 석사)는 한국방재시험연구원(KFPA) 현장실습을 통해 KS F ISO 1182·5660-1·2271 등 화재 표준시험의 화학공학적 해석 경험을 보유하고, 무기계 충전제 분해 동역학(Arrhenius 모델) 및 인계 난연제 시너지 효과 정량화 역량을 갖추고 있다. "
    "연구개발 책임자는 고분자·복합재 합성 박사급 1인을 채용 예정이며, 생산 책임자는 단열재 양산 라인 10년 이상 경력자를 영입할 계획이다. "
    "초기 자문위원으로 KFPA 방재연구센터 출신 인사 1인을 위촉하여 시험·인증 전반의 컨설팅을 받는다."
)

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 2장. BM 타당성 (재활용 표현 제거, 규모 축소)
# ──────────────────────────────────────────────────────────────────────────
write_line("2. BM 타당성", size=12, bold=True)
write_line("", size=12)

write_line("-신제품 및 서비스 추진전략과 사업화 계획", size=12, bold=True)
write_para(
    "본 사업은 외벽 단열재 시장의 규제 강화 흐름에 정면으로 대응하는 친환경 준불연 단열보드를 핵심 사업 아이템으로 한다. "
    "도입 단계에서는 중견 시공사(시공능력평가 30~100위권) 3~5사 및 단열재 전문 대리점과 파일럿 공급 계약을 체결하여 실증 트랙레코드를 확보한다. "
    "성장 단계에서는 KS F ISO 5660-1 준불연 인증 및 KS F 8414 실대형 시험 통과 데이터를 마케팅 자산화하여 공공·민간 발주 시방서에 본 제품군을 등재한다. "
    "확산 단계에서는 동일 난연 플랫폼을 ESS(에너지저장장치) 패키지 단열재, 전기차 배터리 모듈 격벽재 등 인접 시장으로 검토 단계 확장한다. "
    "본 단계 전략은 학교 과제 사업화 모델로서 실현 가능한 단년도 자금 부담을 고려해 국내 시장 우선 안착 후 인접 응용 검토 순으로 설계되었다."
)

write_line("-BM 구체화 및 시장 진입 전략", size=12, bold=True)
write_para(
    "본 사업의 비즈니스 모델은 ① 사전기획 단계 → ② 시장진입 단계 → ③ 양산·확산 단계의 3단계 진행을 표준 골격으로 한다. "
    "사전기획 단계(1차연도)에서는 처방 최적화, 파일럿 라인(연 1,000 m³) 구축, 시제품 인증을 완료한다. "
    "시장진입 단계(2~3차연도)에서는 중견 시공사 시방 등재 및 연 8,000 m³ 양산능력 확보를 목표로 한다. "
    "양산·확산 단계(4~5차연도)에서는 임대형 양산 거점에 연 10,000 m³ 자동화 라인을 구축하고 인접 응용시장 진입을 검토한다."
)

write_line("-구체적인 비즈니스 모델", size=12, bold=True)
write_para(
    "1. 고객 세분화 및 타겟팅: 1차 타겟은 6층 이상 공동주택·주상복합 시공 중견 건설사(시공능력평가 30~100위권) 약 30개사 및 단열재 전문 대리점이며, 2차 타겟은 학교·요양시설 리모델링 시장 및 LH·SH 등 공공발주 기관이다."
)
write_para(
    "2. 가격 책정 전략: 기존 미네랄울 대비 단열성능(λ ≤ 0.022 W/m·K)이 약 35% 우수하고 페놀폼(PF) 대비 흡습률(JIS A 9521 기준 1.2%)이 약 50% 낮은 점을 근거로 m³당 18만~22만 원의 프리미엄 가격을 책정하며, 누적 발주량 500 m³ 이상 고객에 대해 단계적 할인(최대 8%) 구조를 적용한다."
)
write_para(
    "3. 판매 채널 구축: 직접 영업(중견 시공사)과 대리점 채널(중소 시공·리모델링 시장)을 병행한다. 초기에는 KFPA·KCL 등 공인기관 시험성적서를 영업 자산으로 활용하고, 2차연도부터는 자체 웹사이트 기반 온라인 견적·기술자료 다운로드 페이지를 운영하여 영업 효율을 제고한다."
)
write_para(
    "4. 인력 확보 및 조직 구성: 핵심 R&D 인력(석·박사급 2인)과 양산 엔지니어(3~5인)를 단계별로 확보하며, 시험·인증·품질관리 인력은 KFPA·KCL·FITI 출신 경력직 1인을 우선 채용한다."
)
write_para(
    "5. 효율적인 유통망 구축: 수도권 직영 거점 1개소와 권역별 위탁 물류센터(중부·영남) 2개소를 활용하고, 시공 현장 24~48시간 직배송 체계를 통해 재고 회전율을 경쟁사 대비 약 1.3배 수준으로 유지한다."
)
write_para(
    "6. 마케팅 및 홍보 전략: 건축·소방·화학공학 학회 및 KOREA BUILD 부스 운영을 통해 기술 신뢰도를 확보한다. 또한 콘칼로리미터 실측 데이터를 영상화한 디지털 콘텐츠를 발주처에 배포하여 '데이터 기반 안전성 검증' 메시지를 일관되게 전달한다."
)
insert_picture(IMG_FACADE_RIG, caption="<그림 3> 외벽 마감재 실대형 모형 화재시험기 (KS F 8414)")

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 3장. 시장 분석 (경쟁사 분석 + 표 2 삭제)
# ──────────────────────────────────────────────────────────────────────────
write_line("3. 시장 분석", size=12, bold=True)
write_line("", size=12)

write_line("-국내외 산업 동향", size=12, bold=True)
write_para(
    "국내 건축용 단열재 시장은 2024년 기준 약 2.4조 원 규모로 추정되며(한국건설기술연구원, 「건축물 단열재 시장 동향 보고서」, 2024), "
    "외벽 준불연·불연 단열재 비중은 2018년 약 12%에서 2024년 약 38%로 빠르게 상승하였다. "
    "이는 국토교통부 고시 제2022-84호 시행과 함께 EPS·PU 단일 시스템 대비 무기계·페놀계 단열재 채택 비율이 가파르게 증가한 결과이다. "
    "글로벌 시장에서는 EU 그린딜 및 미국 IRA(인플레이션 감축법) 시행 이후 친환경·저탄소 단열재에 대한 수요가 연 7~9% 성장 중이며, "
    "특히 할로겐프리 난연제 시장은 2023~2030년 연평균 성장률(CAGR) 8.2%로 예측된다(MarketsandMarkets, "
    "「Halogen-free Flame Retardants Market Report」, 2024). "
    "본 사업은 위 두 동인(국내 규제 강화·글로벌 친환경 수요)이 동시에 작용하는 시장 구간에 위치한다."
)

write_line("-시장 규모 및 성장성", size=12, bold=True)
write_para(
    "본 사업의 시장 규모는 TAM(전체 단열재 시장) 약 2.4조 원, "
    "SAM(외벽 준불연·불연 단열재) 약 9,100억 원, "
    "SOM(친환경 하이브리드 단열보드 신규 수요) 약 1,200억 원으로 산정된다. "
    "5년차 SOM 점유율 4% 달성 시 본 사업의 직접 매출은 약 48억 원 수준으로 추정되며, "
    "이는 학교 과제 단계의 사업화 모델로서 단년도 운영 가능한 인력·설비 규모(연 10,000 m³ 양산)와 정합한다."
)
write_line("<표 1> 시장 규모 및 점유 목표", size=11, bold=True, align="center")
make_table(
    headers=["구분", "정의", "규모(원)", "출처"],
    rows=[
        ["TAM", "국내 단열재 전체 시장", "약 2.4조",        "한국건설기술연구원(2024)"],
        ["SAM", "외벽 준불연·불연 단열재", "약 9,100억",   "한국건설기술연구원(2024) 재가공"],
        ["SOM", "친환경 하이브리드 단열보드", "약 1,200억", "MarketsandMarkets(2024) 재가공"],
        ["목표", "5년차 SOM 점유율 4%",       "약 48억",     "본 사업 추정"],
    ],
    col_widths_mm=[20, 55, 30, 45],
    body_align="left",
)

write_line("-SWOT 분석", size=12, bold=True)
make_table(
    headers=["구분", "내용"],
    rows=[
        ["Strengths",     "4중 난연 시너지 처방, KFPA 시험 노하우, 할로겐프리 친환경 처방"],
        ["Weaknesses",   "신생 브랜드 인지도 부족, 양산 초기 단가 경쟁력 미흡"],
        ["Opportunities","외벽 불연·준불연 의무화 확대, 친환경 인증 수요 증가"],
        ["Threats",      "대형사의 후발 진입, 원자재(APP·실리콘) 가격 변동, 규제 추가 강화"],
    ],
    col_widths_mm=[28, 122],
    body_align="left",
    row_h_mm=10,
)

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 4장. 추진 일정 및 단계별 계획 (규모 축소)
# ──────────────────────────────────────────────────────────────────────────
write_line("4. 추진 일정 및 단계별 계획", size=12, bold=True)
write_line("", size=12)

write_line("-단계별 로드맵", size=12, bold=True)
write_para(
    "1차연도(2026)는 처방 최적화 및 핵심 특허 출원, 파일럿 라인(연 1,000 m³) 구축, KS F ISO 5660-1·1182 인증 획득을 목표로 한다. "
    "2차연도(2027)는 KS F 8414 실대형 모형 인증 및 시제품 양산(연 3,000 m³), 중견 시공사 3~5사 시방 등재를 추진한다. "
    "3차연도(2028)는 양산 라인 증설(연 8,000 m³)과 LH·SH 등 공공발주 시범 진입을 목표로 한다. "
    "4~5차연도(2029~2030)는 자동화 라인(연 10,000 m³) 구축 및 인접 시장(ESS·EV 단열재) 진입 검토를 추진한다. "
    "본 로드맵의 양산 capacity는 단열재 압출·코팅 라인의 시간당 표준 처리량(약 1.5 m³/시간)과 1일 2교대 가동 가정에 근거한다."
)
write_line("<표 2> 단계별 로드맵 요약", size=11, bold=True, align="center")
make_table(
    headers=["연차", "핵심 활동", "양산 능력"],
    rows=[
        ["1차(2026)", "처방 최적화·파일럿 라인 구축·KS 인증",       "연 1,000 m³"],
        ["2차(2027)", "KS F 8414 인증·시제품 양산·시방 등재",       "연 3,000 m³"],
        ["3차(2028)", "양산 라인 증설·공공발주 시범 진입",           "연 8,000 m³"],
        ["4차(2029)", "자동화 라인 구축·운영 안정화",                  "연 10,000 m³"],
        ["5차(2030)", "인접 시장(ESS·EV) 진입 검토",                  "연 10,000 m³"],
    ],
    col_widths_mm=[25, 95, 30],
    body_align="left",
)

write_line("-주요 마일스톤", size=12, bold=True)
make_table(
    headers=["시점", "마일스톤"],
    rows=[
        ["Q1 2026", "핵심 특허 2건 출원, KFPA 콘칼로리미터 사내 시험 통과(pHRR ≤ 250 kW/m²)"],
        ["Q3 2026", "파일럿 라인 가동, KS F ISO 5660-1 준불연 등급 획득"],
        ["Q2 2027", "KS F 8414 실대형 시험 통과, 중견 시공사 3사 PoC 완료"],
        ["Q4 2027", "Pre-Series A 투자 10억 원 유치"],
        ["Q3 2028", "누적 매출 24억 원 달성, 공공발주 시범 진입"],
        ["Q4 2030", "누적 매출 50억 원 및 영업이익 흑자 전환"],
    ],
    col_widths_mm=[28, 122],
    body_align="left",
)
insert_picture(IMG_NONFLAM_PHOTO, caption="<그림 4> 불연성 시험기(KS F ISO 1182) — 인증 단계 핵심 장비")

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 5장. 조직 구성 및 인력 운영 계획 (표 4 단순화 + 근거)
# ──────────────────────────────────────────────────────────────────────────
write_line("5. 조직 구성 및 인력 운영 계획", size=12, bold=True)
write_line("", size=12)

write_line("-조직도", size=12, bold=True)
write_para(
    "본 사업의 조직은 대표이사 직속의 ① 연구개발본부, ② 생산·품질본부, ③ 영업·경영지원본부의 3본부 체제로 구성한다. "
    "초기에는 조직 규모를 최소화하기 위해 시험·인증 기능은 연구개발본부 산하에 통합 운영하고, 구매·물류 기능은 생산·품질본부 내에서 겸직 형태로 운영한다."
)
insert_picture(IMG_ORG_CHART, caption="<그림 5> 조직 구성 참조 (KFPA 조직도 재해석)")

write_line("-단계별 인력 확보 계획", size=12, bold=True)
write_para(
    "인력 확보 규모는 연차별 양산 capacity(<표 2>) 및 매출 규모(<표 4>)에 비례하여 산정한다. "
    "구체적 산정 근거는 다음과 같다. "
    "① 연구개발 인력은 처방 1건당 시험·반복최적화 1인×6개월 공수를 기준으로 1차 2명에서 5차 4명으로 단계 확대한다. "
    "② 생산·품질 인력은 단열재 압출·코팅 라인 1개소당 교대 운영 인력 4명(2교대 × 라인 1인 + 품질 1인 + 보조 1인) 표준에 근거하여 양산 라인 수 증가에 따라 비례 확대한다. "
    "③ 영업·경영지원 인력은 거래처(중견 시공사 + 대리점) 30개사 기준 영업 1인당 10개사 담당 표준에 근거하여 산정한다. "
    "본 계획은 단년도 인건비 부담(평균 인건비 7천만 원 가정 시)이 매출 대비 과도해지지 않는 범위 내에서 점진 채용을 원칙으로 한다."
)
write_line("<표 3> 연차별 인력 확보 계획 (단위: 명)", size=11, bold=True, align="center")
make_table(
    headers=["연차", "연구개발", "생산·품질", "영업·경영지원", "합계"],
    rows=[
        ["1차(2026)", "2", "1", "1", "4"],
        ["3차(2028)", "3", "6", "3", "12"],
        ["5차(2030)", "4", "14", "6", "24"],
    ],
    col_widths_mm=[28, 28, 28, 36, 25],
)

write_line("-역량 강화 방안", size=12, bold=True)
write_para(
    "산학협력으로 성균관대 화학공학과와 공동연구 협의체를 운영하고, KFPA·KCL과 시험·인증 표준화 협력 협약(MOU)을 체결한다. "
    "사내 교육은 분기별 KS F ISO 5660-1·8414 시험 절차 워크숍과 화재공학(Fire Science) 입문 과정을 운영하여 전 직원의 기술 이해도를 균등화한다. "
    "외부 전문가 자문단(난연제·고분자 각 1인)을 위촉하여 분기 1회 자문 회의를 진행한다."
)

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 6장. 재무 계획 및 자금 조달 (표 5,6 + 산정 근거)
# ──────────────────────────────────────────────────────────────────────────
write_line("6. 재무 계획 및 자금 조달", size=12, bold=True)
write_line("", size=12)

write_line("-매출 및 손익 추정", size=12, bold=True)
write_para(
    "매출 추정의 근거는 다음과 같다. 본 제품의 평균 단가는 비즈니스 모델 ②항에서 책정한 m³당 20만 원을 기준으로 한다. "
    "1차연도(2026)는 파일럿 라인(연 1,000 m³) 가동 후 인증·시범 공급 단계로 가동률 약 75%인 750 m³ 공급을 가정하여 매출 약 1.5억 원을 산정한다. "
    "2차연도(2027)는 시제품 양산(연 3,000 m³) 단계로, KS F 8414 인증 후 중견 시공사 시방 등재가 점진 진행되는 점을 고려해 가동률 약 75%인 4,250 m³ 공급(2차 라인 증설분 포함)으로 매출 약 8.5억 원을 산정한다. "
    "3차연도(2028)는 양산 라인 증설(연 8,000 m³) 단계로, 공공발주 시범 진입을 반영해 가동률 약 75%인 약 10,500 m³ 공급(증설 누적 capacity 14,000 m³ 기준)으로 매출 약 21억 원을 산정한다. "
    "매출원가율은 무기계 충전제(Al(OH)₃ 약 1,800원/kg), 페놀 코어(약 4,500원/kg), APP·MEL 난연제(약 6,000~9,000원/kg) 시장 단가와 인건비·전력비(라인당 월 약 600만 원)를 합산하여 산정하였으며, 1차 85%에서 3차 73% 수준으로 양산 학습효과(Wright's law, 누적생산 2배 시 단가 약 12% 감소 가정)에 의해 점진 개선되는 것으로 추정하였다."
)
write_line("<표 4> 3개년 추정 손익계산서 (단위: 백만 원)", size=11, bold=True, align="center")
make_table(
    headers=["구  분", "1차(2026)", "2차(2027)", "3차(2028)"],
    rows=[
        ["매출액",       "150",  "850",  "2,100"],
        ["매출원가",     "128",  "680",  "1,540"],
        ["매출총이익",    "22",  "170",    "560"],
        ["판관비",       "280",  "400",    "500"],
        ["영업이익",    "-258", "-230",     "60"],
        ["영업이익률", "-172.0%","-27.1%",  "2.9%"],
    ],
    col_widths_mm=[35, 32, 32, 32],
)

write_line("-투자 계획", size=12, bold=True)
write_para(
    "투자 항목별 산정 근거는 다음과 같다. "
    "(1) R&D 비용은 KFPA·KCL의 콘칼로리미터(KS F ISO 5660-1) 시험 단가 약 200만 원/회, 실대형 모형 화재시험(KS F 8414) 단가 약 5천만 원/회의 공인 시험 단가에 처방 변형 시 반복 시험을 가정하여 연 8~15회 분량의 시험비와 연구 인건비(석·박사 평균 8천만 원/년), 원료 시약비를 합산하였다. "
    "(2) 설비 투자는 1차연도 파일럿 라인(소형 이축압출기 + 표면 코팅 + 절단 라인, 시장 견적 약 3억 원) 도입을 시작으로, 2차연도에 양산 보조 설비(자동 절단·포장기, 약 7억 원), 3차연도에 본 양산 라인(연 5,000 m³급 자동화, 견적 약 12억 원)을 단계 증설하는 것으로 산정하였다. "
    "(3) 운영자금은 <표 3>의 연차별 인력 수에 평균 인건비(석·박사 8천만 원, 학사 5천만 원)를 적용한 인건비, 임대료(150 m² × 평균 평당 5만 원/월 가정), 학회·전시회·디지털 마케팅 비용을 합산한 값이다."
)
write_line("<표 5> 3개년 투자 계획 (단위: 백만 원)", size=11, bold=True, align="center")
make_table(
    headers=["항목", "1차(2026)", "2차(2027)", "3차(2028)", "산정 근거 요약"],
    rows=[
        ["R&D",    "250",  "350",  "400",  "공인 시험 단가 + R&D 인건비 + 시약비"],
        ["설비",   "300",  "700", "1,200", "파일럿→양산 보조→본 양산 라인 증설"],
        ["운영자금","300",  "450",  "600",  "인건비(<표 3>) + 임대·마케팅"],
        ["합  계", "850", "1,500","2,200", "3개년 누적 약 45.5억 원"],
    ],
    col_widths_mm=[22, 22, 22, 22, 62],
    body_align="left",
)

write_line("-자금 조달 방안", size=12, bold=True)
write_para(
    "초기 자기자본(공동창업자) 2억 원과 정부지원사업(창업성장기술개발 R&D 2억 원, 소부장 자율형 4억 원)을 1차연도에 확보한다. "
    "2차연도에는 Pre-Series A 투자 10억 원을 VC 중심으로 유치하고, 3차연도에는 Series A 15억 원과 산업은행·중소벤처기업진흥공단 시설자금 융자 5억 원을 결합하여 양산 라인 증설 자금을 조달한다. "
    "조달 우선순위는 정부지원→VC→정책금융 순이며, 지분 희석을 최소화하기 위해 R&D 단계는 비희석성 자금 비중을 70% 이상으로 유지한다. "
    "위 조달계획의 총 누적 조달액은 약 38억 원으로, <표 5> 누적 투자 약 45.5억 원과 영업현금흐름·운전자본 변동의 차이를 정부지원·정책금융 후속 활용으로 보완한다."
)
insert_picture(IMG_GAS_PHOTO, caption="<그림 6> 가스유해성 시험기(KS F 2271) — 인증 비용 핵심 항목")

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 7장. 리스크 분석 및 대응 방안 (재활용 표현 제거, 규모 축소 반영)
# ──────────────────────────────────────────────────────────────────────────
write_line("7. 리스크 분석 및 대응 방안", size=12, bold=True)
write_line("", size=12)

write_line("-기술 리스크", size=12, bold=True)
write_para(
    "처방 안정성 저하 리스크는 양산 환경에서 입자 크기·습도 변화로 pHRR 재현성이 낮아질 가능성이며, 대응방안으로 SPC(통계적 공정관리) 도입 및 In-line FT-IR 모니터링으로 변동요인을 실시간 제어한다. "
    "또한 인증 미통과 리스크는 KS F 8414 실대형 시험 1차 미통과 가능성이며, 대응방안으로 KFPA·KCL 사전 예비시험을 2회 이상 거치고, 외벽 시스템 단위로 다중 처방 트랙(Plan A·B)을 병렬 진행한다. "
    "본 대응 비용은 <표 5> R&D 항목 내 사전 예비시험 4~6회 분량으로 사전 반영되어 있다."
)

write_line("-시장 리스크", size=12, bold=True)
write_para(
    "대형 경쟁사 진입 리스크는 글로벌 단열재 기업의 한국 시장 가격공세 가능성이며, 대응방안으로 친환경(할로겐프리)·고성능 차별화 메시지와 공공발주 시방 등재로 진입장벽을 구축한다. "
    "규제 변동 리스크는 외벽 마감재 등급 기준 추가 강화 가능성이며, 대응방안으로 등급 상향(준불연→불연) 처방 후속 R&D를 1년 선행 운영한다."
)

write_line("-재무 리스크", size=12, bold=True)
write_para(
    "양산 라인 투자 부담 리스크는 3차연도 약 12억 원 설비투자에 따른 현금흐름 악화 가능성이며, 대응방안으로 시설자금 정책융자(산업은행·중진공) 및 리스 방식 분산 조달로 단년도 부담을 30% 이내로 제한한다. "
    "원자재 가격 변동 리스크는 APP·실리콘·Al(OH)₃ 가격 급등 가능성이며, 대응방안으로 2개 이상 공급선 이원화와 6개월 단위 선구매 헷지 계약을 운영한다."
)

write_line("-규제 리스크", size=12, bold=True)
write_para(
    "화학물질평가법·REACH 등록 리스크는 신규 난연제 조성의 등록 지연 가능성이며, 대응방안으로 기존 등록물질(APP·MEL·실리콘) 위주로 처방을 구성하고, 신규 물질은 학술용 소량 등록으로 우선 대응한다. "
    "건축법령 해석 차이 리스크는 지자체별 외벽 마감재 적용 해석 편차이며, 대응방안으로 국토안전관리원·KFPA 공동 기술 가이드 발간으로 해석 통일성을 제고한다."
)
insert_picture(IMG_PEN_FIRE, caption="<그림 7> 내화충전구조 시험 — 규제 강화 흐름의 단면")

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 8장. 기대 효과 및 결론 (재활용 표현 제거, 규모 축소 반영)
# ──────────────────────────────────────────────────────────────────────────
write_line("8. 기대 효과 및 결론", size=12, bold=True)
write_line("", size=12)

write_line("-경제적 효과", size=12, bold=True)
write_para(
    "5년차 누적 매출 약 50억 원, 누적 고용 24명, 협력사 5개사를 통한 간접 고용 약 30명을 창출한다. "
    "수입 의존 페놀폼·할로겐 난연제 일부를 국산 친환경 처방으로 대체하여 연 약 20억 원 수준의 수입대체 효과가 기대된다."
)

write_line("-사회적·환경적 효과", size=12, bold=True)
write_para(
    "본 제품 적용 시 외벽 화재로 인한 인명피해 발생률을 통계적으로 약 30% 저감할 수 있을 것으로 추정되며(KFPA 사고통계 기반 시뮬레이션), "
    "할로겐프리·저독성 처방 구조를 통해 화재 시 다이옥신·HBr 등 유해가스 방출량이 기존 할로겐계 대비 사실상 0에 수렴한다. "
    "또한 단열성능 향상에 따른 건축물 운영단계 에너지 사용량을 m²당 약 12% 절감하여 탄소중립 목표에 기여한다."
)

write_line("-기술적 파급 효과", size=12, bold=True)
write_para(
    "본 사업의 4중 난연 처방 플랫폼은 ESS 셀 격벽재, 전기차 배터리 모듈 단열재, 데이터센터 케이블 트레이용 난연 슬리브 등 인접 응용으로 확장 가능하다. "
    "또한 콘칼로리미터 데이터를 FDS(Fire Dynamics Simulator) 입력값으로 변환하는 디지털 트윈 모듈은 화재 시뮬레이션 기반 컨설팅 서비스로의 확장도 가능하다."
)

write_line("-결론", size=12, bold=True)
write_para(
    "본 사업은 반복적 외벽 화재로 야기된 사회적 비용을 화학공학적 원리에 기반한 4중 난연 시너지 처방으로 근본 해결한다는 명확한 가치제안을 가진다. "
    "KFPA 현장에서 검증된 표준시험 데이터, 할로겐프리·저독성 친환경 처방 구조, 단계적 자금 조달 전략의 3박자가 결합되어 기술·시장·재무 측면 모두에서 사업화 타당성이 높다. "
    "본 계획서에 따른 단계별 실행을 통해 5년 내 누적 매출 약 50억 원과 국내 외벽 단열재 시장 내 친환경 표준 기업으로의 자리매김을 달성할 수 있을 것으로 판단된다."
)

# 페이지번호
try:
    hwp.HAction.GetDefault("PageNumPos", hwp.HParameterSet.HPageNumPos.HSet)
    pn = hwp.HParameterSet.HPageNumPos
    pn.Pos = 8
    pn.NumberFormat = 0
    pn.SideChar = 1
    hwp.HAction.Execute("PageNumPos", pn.HSet)
except Exception as e:
    print("PageNum err:", e)

# 저장
hwp.SaveAs(OUT, "HWP", "")
print("Saved:", OUT)
hwp.Quit()
