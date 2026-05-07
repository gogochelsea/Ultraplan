# -*- coding: utf-8 -*-
"""
사업 계획서.hwp 자동 생성 스크립트 (Hancom Office COM Automation)
- 본문 12pt, 단일 행간 (160% HWP 표준 단일 행간)
- 표지는 글자 크기 예외
"""
import os, sys, time
import win32com.client as win32

BASE = r"C:\Users\filk2\Desktop\공부자료\학교수업\화공산"
IMG = os.path.join(BASE, "extracted_images")
OUT = os.path.join(BASE, "사업 계획서.hwp")

# Image map (use original extracted files for hwp insert; png for bmp originals so HWP accepts)
IMG_ORG_CHART     = os.path.join(IMG, "BIN0003.png")  # 방재시험연구원 조직도 → BM 참조
IMG_AERIAL        = os.path.join(IMG, "BIN0004.png")  # 연구원 전경
IMG_NONFLAM_DIA   = os.path.join(IMG, "BIN0001.jpg")  # 불연성 시험 도식
IMG_NONFLAM_PHOTO = os.path.join(IMG, "BIN0002.JPG")  # 불연성 시험기 사진
IMG_CONE_PHOTO    = os.path.join(IMG, "BIN0005.JPG")  # 콘칼로리미터 사진
IMG_CONE_DIA      = os.path.join(IMG, "BIN0006.jpg")  # 콘칼로리미터 도식
IMG_GAS_PHOTO     = os.path.join(IMG, "BIN0007.JPG")  # 가스유해성 시험기
IMG_FACADE_RIG    = os.path.join(IMG, "BIN0008.png")  # 외벽 실물모형 시험기
IMG_FACADE_FIRE   = os.path.join(IMG, "BIN0009.png")  # 외벽 화재 장면
IMG_GAS_DIA       = os.path.join(IMG, "BIN000A.jpg")  # 가스유해성 도식
IMG_CABLE_TEST    = os.path.join(IMG, "BIN000C.png")  # 케이블 화염전파 시험
IMG_CABLE_BUNDLE  = os.path.join(IMG, "BIN000D.png")  # 케이블 묶음
IMG_PEN_FIRE      = os.path.join(IMG, "BIN000E.png")  # 내화충전구조

# ──────────────────────────────────────────────────────────────────────────
# COM 초기화
# ──────────────────────────────────────────────────────────────────────────
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
try:
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
except Exception:
    pass
hwp.XHwpWindows.Item(0).Visible = False
hwp.SetMessageBoxMode(0x2FF1)  # 메시지박스 자동 처리

# 새 문서
hwp.HAction.Run("FileNew")

# ──────────────────────────────────────────────────────────────────────────
# 헬퍼 함수
# ──────────────────────────────────────────────────────────────────────────
def set_char_shape(size_pt=12, bold=False, italic=False, color=0, face="함초롬바탕"):
    """문자 모양 설정 — Ratio/Size 100, FontType=1(TTF) 필수, 빈 값이면 렌더 실패"""
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    cs = hwp.HParameterSet.HCharShape
    cs.Height = int(size_pt * 100)
    cs.Bold = 1 if bold else 0
    cs.Italic = 1 if italic else 0
    cs.TextColor = color
    for ln in ["Hangul", "Latin", "Hanja", "Japanese", "Other", "Symbol", "User"]:
        setattr(cs, f"FaceName{ln}", face)
        setattr(cs, f"FontType{ln}", 1)   # TTF
        setattr(cs, f"Offset{ln}", 0)
        setattr(cs, f"Ratio{ln}", 100)
        setattr(cs, f"Spacing{ln}", 0)
        setattr(cs, f"Size{ln}", 100)
    hwp.HAction.Execute("CharShape", cs.HSet)

def set_para_shape(align="left", line_space=160, space_before=0, space_after=0):
    """문단 모양 설정"""
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
    """현재 위치에 텍스트 삽입"""
    if not text:
        return
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

def br():
    """문단 나눔 (Enter)"""
    hwp.HAction.Run("BreakPara")

def page_break():
    """페이지 나눔"""
    hwp.HAction.Run("BreakPage")

def write_line(text, size=12, bold=False, align="left", color=0):
    """한 줄 쓰고 엔터"""
    set_char_shape(size_pt=size, bold=bold, color=color)
    set_para_shape(align=align)
    insert_text(text)
    br()

def write_para(text, size=12, bold=False, align="justify"):
    """문단(자동 줄바꿈) 쓰고 엔터"""
    set_char_shape(size_pt=size, bold=bold)
    set_para_shape(align=align)
    insert_text(text)
    br()

def insert_picture(path, width_mm=120, height_mm=None, caption=None):
    """이미지 삽입 — width_mm 기준, 비율 유지"""
    if not os.path.exists(path):
        print(f"[WARN] image not found: {path}")
        return
    # 원본 비율 계산
    try:
        from PIL import Image
        with Image.open(path) as im:
            iw, ih = im.size
        aspect = ih / iw
    except Exception:
        aspect = 0.7
    if height_mm is None:
        height_mm = width_mm * aspect
        # 한 페이지에 너무 크지 않도록 최대 높이 130mm 제한
        if height_mm > 130:
            height_mm = 130
            width_mm = height_mm / aspect

    set_para_shape(align="center")
    set_char_shape(size_pt=12)

    # HWPUNIT: 1 mm ≈ 283.465
    HWPMM = 283.464567
    w_unit = int(width_mm * HWPMM)
    h_unit = int(height_mm * HWPMM)

    # InsertPicture: filename, Embedded, sizeoption (0=원본 1=종이맞춤 2=셀맞춤 3=...)
    # 반환값: 컨트롤 객체 (가능하다면)
    ctrl = None
    try:
        ctrl = hwp.InsertPicture(path, True, 2, False)  # 2: SetSpecificSize
    except Exception as e:
        try:
            ctrl = hwp.InsertPicture(path, True, 0, False)
        except Exception as e2:
            print("InsertPicture err:", e2)

    # 마지막 컨트롤이 방금 삽입한 이미지 — 크기 직접 설정
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
        set_char_shape(size_pt=12, bold=False)
        set_para_shape(align="center")
        insert_text(caption)
        br()

# ──────────────────────────────────────────────────────────────────────────
# 표지 (글자 크기 예외)
# ──────────────────────────────────────────────────────────────────────────
# 빈 줄
for _ in range(4):
    write_line("", size=12)

write_line("사 업 계 획 서", size=32, bold=True, align="center")
write_line("", size=12)
write_line("", size=12)

write_line("─────────────────────────────────", size=14, align="center")
write_line("", size=12)
write_line("할로겐프리 하이브리드 난연 시스템 기반", size=20, bold=True, align="center")
write_line("친환경·재활용 외벽 단열보드 「Safe-Core™」", size=20, bold=True, align="center")
write_line("", size=12)
write_line("─────────────────────────────────", size=14, align="center")

for _ in range(8):
    write_line("", size=12)

write_line("성균관대학교 일반대학원 화학공학과", size=16, align="center")
write_line("", size=12)
write_line("대 표 자  :  박  정  빈", size=16, align="center")
write_line("", size=12)
write_line("제 출 일  :  2026. 05. 06.", size=16, align="center")

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 목차
# ──────────────────────────────────────────────────────────────────────────
write_line("목   차", size=12, bold=True, align="center")
write_line("", size=12)

toc = [
    ("1. 사업 개요 및 기술 소개", "1"),
    ("    -사업 추진 배경 및 필요성", ""),
    ("    -사업 비전과 미션", ""),
    ("    -제품 및 기술 개요", ""),
    ("    -팀 구성 요약", ""),
    ("2. BM 타당성", "5"),
    ("    -신제품 및 서비스 추진전략과 사업화 계획", ""),
    ("    -BM 구체화 및 시장 진입 전략", ""),
    ("    -구체적인 비즈니스 모델", ""),
    ("3. 시장 및 경쟁 분석", "10"),
    ("    -국내외 산업 동향", ""),
    ("    -시장 규모 및 성장성", ""),
    ("    -경쟁사 및 경쟁 기술 분석", ""),
    ("    -SWOT 분석", ""),
    ("4. 추진 일정 및 단계별 계획", "14"),
    ("    -단계별 로드맵", ""),
    ("    -주요 마일스톤", ""),
    ("5. 조직 구성 및 인력 운영 계획", "16"),
    ("    -조직도", ""),
    ("    -단계별 인력 확보 계획", ""),
    ("    -역량 강화 방안", ""),
    ("6. 재무 계획 및 자금 조달", "18"),
    ("    -매출 및 손익 추정", ""),
    ("    -투자 계획", ""),
    ("    -자금 조달 방안", ""),
    ("7. 리스크 분석 및 대응 방안", "21"),
    ("8. 기대 효과 및 결론", "23"),
]
for line, p in toc:
    if line.startswith("    "):
        write_line(line, size=12, align="left")
    else:
        write_line(line, size=12, bold=True, align="left")

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 1장. 사업 개요 및 기술 소개
# ──────────────────────────────────────────────────────────────────────────
write_line("1. 사업 개요 및 기술 소개", size=12, bold=True)
write_line("", size=12)

write_line("-사업 추진 배경 및 필요성", size=12, bold=True)
write_para(
    "최근 10년간 의정부 도시형생활주택(2015), 제천 스포츠센터(2017), 울산 주상복합(2020) 등 외벽 가연성 단열재로 인한 대형 인명사고가 반복적으로 발생하였다. "
    "이들 사고의 공통 원인은 EPS(발포 폴리스티렌) 및 PU(폴리우레탄)폼이 외벽 시스템에 다층 구조로 적용되어 굴뚝효과(stack effect)에 의해 화염이 수직 방향으로 급속히 확산된 데 있다. "
    "이에 국토교통부는 고시 제2023-24호를 통해 6층 이상 또는 높이 22 m 이상 건축물의 외벽 마감재료에 대해 실대형 모형 화재시험(KS F 8414)과 마감재 단위 불연·준불연 등급 동시 충족을 의무화하였다. "
    "그러나 현행 시장의 무기계 단열재(미네랄울)는 단열 성능이 낮고, 페놀폼(PF)은 흡습 시 강도가 저하되며, 할로겐계 난연제(브롬계)는 다이옥신 방출 우려로 EU REACH·국내 화평법 규제 대상이다. "
    "본 사업은 인계·질소계·실리콘계 하이브리드 난연 시스템과 무기계 충전제(Al(OH)₃, Mg(OH)₂)를 결합한 친환경·재활용 가능 외벽 단열보드 'Safe-Core™' 를 개발하여 이러한 시장의 미충족 수요(unmet need)를 해소하고자 한다."
)
insert_picture(IMG_FACADE_FIRE, caption="<그림 1> 외벽 마감재 화재 확산 사례 (KFPA 실대형 모형 화재시험)")

write_line("-사업 비전과 미션", size=12, bold=True)
write_para(
    "본 사업의 비전은 '국내 외벽 단열재 시장의 안전·환경 패러다임 전환을 선도하는 친환경 본질안전(intrinsically safe) 소재 기업' 이다. "
    "미션으로는 다음 3가지를 설정한다. 첫째, 할로겐프리·저독성 난연 처방 표준화를 통해 국내 건축물의 화재 인명피해를 감축한다. "
    "둘째, 폐기 시 재활용 가능한 무기·유기 복합 단열재를 상용화하여 건설폐기물 발생량을 저감한다. "
    "셋째, KS·ISO·EN 인증 동시 획득으로 동남아·중동·EU 시장 진출 기반을 확보한다. "
    "중장기 목표는 3년차 매출 80억 원, 5년차 250억 원 및 국내 준불연 외벽 단열재 시장 점유율 8% 달성이다."
)

write_line("-제품 및 기술 개요", size=12, bold=True)
write_para(
    "Safe-Core™ 는 페놀폼(PF) 코어를 인-질소 시너지 난연제(APP·MEL)와 실리콘계 그래프트 코폴리머로 후처리하고, 표면을 Al(OH)₃ 충전 무기계 마감층으로 이중 코팅한 다층 단열보드이다. "
    "핵심 화공 메커니즘은 ① Al(OH)₃의 흡열 분해(약 220 kJ/kg, 180~200 ℃에서 H₂O 생성)에 의한 표면 냉각, "
    "② APP(암모늄 폴리포스페이트)의 탄화층 형성에 의한 산소·열 차단, "
    "③ 멜라민의 NH₃ 방출에 의한 가연성 가스 희석, "
    "④ 실리콘 라디칼 스캐빈저에 의한 연쇄반응 차단의 4중 난연 시너지이다. "
    "산소소비원리(Oxygen Consumption Principle, Huggett 상수 13.1 kJ/g) 기반 콘칼로리미터(KS F ISO 5660-1) 사내 평가에서 최대 열방출률(pHRR) 약 220 kW/m², 총방출열량(THR_10min) 약 6.8 MJ/m² 수준으로 측정되어 준불연재료 합격 기준(pHRR ≤ 200~250 kW/m², THR ≤ 8 MJ/m²)을 만족한다."
)
insert_picture(IMG_CONE_DIA, caption="<그림 2> 콘칼로리미터 시험 원리 (KS F ISO 5660-1)")

write_line("-팀 구성 요약", size=12, bold=True)
write_para(
    "대표(박정빈, 화학공학 석사)는 한국방재시험연구원(KFPA) 현장실습을 통해 KS F ISO 1182·5660-1·2271 등 화재 표준시험의 화학공학적 해석 경험을 보유하고, 무기계 충전제 분해 동역학(Arrhenius 모델) 및 인계 난연제 시너지 효과 정량화 역량을 갖추고 있다. "
    "연구개발 책임자는 고분자·복합재 합성 박사급 전문 인력을 채용 예정이며, 생산 책임자는 단열재 양산 라인 10년 경력자를 영입 계획이다. "
    "초기 자문위원으로 KFPA 방재연구센터 출신 인사 2인을 위촉하여 시험·인증 전반의 컨설팅을 받는다."
)

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 2장. BM 타당성
# ──────────────────────────────────────────────────────────────────────────
write_line("2. BM 타당성", size=12, bold=True)
write_line("", size=12)

write_line("-신제품 및 서비스 추진전략과 사업화 계획", size=12, bold=True)
write_para(
    "본 사업은 외벽 단열재 시장의 규제 강화 흐름에 정면으로 대응하는 친환경 준불연 단열보드를 핵심 사업 아이템으로 한다. "
    "도입 단계에서는 특정 건설사(상위 10대 시공사) 및 PF·미네랄울 유통 채널과 파일럿 공급 계약을 체결하여 실증 트랙레코드를 확보한다. "
    "성장 단계에서는 KS F ISO 5660-1 준불연 인증 및 KS F 8414 실대형 시험 통과 데이터를 마케팅 자산화하여 공공·민간 발주 시방서에 본 제품군을 등재한다. "
    "확산 단계에서는 EU CE·EN 13501-1, 미국 ASTM E84·NFPA 285 추가 인증을 통해 글로벌 시장에 진입하며, 시장 성숙기에는 동일 난연 플랫폼을 ESS(에너지저장장치) 패키지 단열재, 전기차 배터리 모듈 격벽재 등 인접 시장으로 확장한다."
)

write_line("-BM 구체화 및 시장 진입 전략", size=12, bold=True)
write_para(
    "본 사업의 비즈니스 모델은 ① 사전기획 단계 → ② 시장진입 단계 → ③ 양산·글로벌 단계의 3단계 진행을 표준 골격으로 한다. "
    "사전기획 단계(1차연도)에서는 처방 최적화, 파일럿 라인 구축, 시제품 인증을 완료한다. "
    "시장진입 단계(2~3차연도)에서는 국내 상위 시공사 시방 등재 및 연 5만 m³ 양산능력 확보를 목표로 한다. "
    "양산·글로벌 단계(4~5차연도)에서는 충북 청주에 연 20만 m³ 규모 자동화 라인을 증설하고 동남아·중동에 합작 법인을 설립한다."
)

write_line("-구체적인 비즈니스 모델", size=12, bold=True)
write_para(
    "1. 고객 세분화 및 타겟팅: 1차 타겟은 6층 이상 공동주택·주상복합 시공 상위 30개 건설사 및 단열재 전문 대리점이며, 2차 타겟은 LH·SH 등 공공발주 기관과 학교·요양시설 리모델링 시장이다."
)
write_para(
    "2. 가격 책정 전략: 기존 미네랄울 대비 단열성능(λ ≤ 0.022 W/m·K)이 30% 우수하고 PF 대비 흡습률이 50% 낮은 점을 근거로 m³당 18만~22만 원의 프리미엄 가격을 책정하며, 대량 공급 시 단계적 할인 구조를 적용한다."
)
write_para(
    "3. 판매 채널 구축: 직접 영업(대형 시공사)과 대리점 채널(중소 시공·리모델링 시장)을 병행한다. 초기에는 KFPA·KCL 등 공인기관 시험성적서를 영업 자산으로 활용하고, 2차연도부터는 자체 e커머스 플랫폼을 통해 견적·발주·배송 전 과정을 디지털화한다."
)
write_para(
    "4. 인력 확보 및 조직 구성: 핵심 R&D 인력(석·박사급 4인)과 양산 엔지니어(7인)를 단계별로 확보하며, 시험·인증·품질관리 인력은 KFPA·KCL·FITI 출신 경력직을 우선 채용한다."
)
write_para(
    "5. 효율적인 유통망 구축: 권역별 거점 물류센터(수도권·중부·영남·호남) 4개소를 구축하고, 시공 현장 24시간 직배송 체계를 통해 재고 회전율을 경쟁사 대비 1.5배 수준으로 유지한다."
)
write_para(
    "6. 마케팅 및 홍보 전략: 건축·소방·화학공학 학회 및 KOREA BUILD·세계화재안전엑스포 부스 운영을 통해 기술 신뢰도를 확보한다. 또한 콘칼로리미터 실측 데이터를 영상화한 디지털 콘텐츠를 발주처에 배포하여 '데이터 기반 안전성 검증' 메시지를 일관되게 전달한다."
)
insert_picture(IMG_FACADE_RIG, caption="<그림 3> 외벽 마감재 실대형 모형 화재시험기 (KS F 8414)")

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 3장. 시장 및 경쟁 분석
# ──────────────────────────────────────────────────────────────────────────
write_line("3. 시장 및 경쟁 분석", size=12, bold=True)
write_line("", size=12)

write_line("-국내외 산업 동향", size=12, bold=True)
write_para(
    "국내 건축용 단열재 시장은 2024년 기준 약 2.4조 원 규모로 추정되며, 외벽 준불연·불연 단열재 비중은 2018년 12%에서 2024년 38%로 빠르게 상승하였다(한국건설기술연구원, 2024). "
    "이는 국토교통부 고시 강화와 함께 EPS·PU 단일 시스템 대비 무기계·페놀계 단열재의 채택 비율이 가파르게 증가한 결과이다. "
    "글로벌 시장에서는 EU 그린딜 및 미국 IRA(인플레이션 감축법) 시행 이후 친환경·저탄소 단열재에 대한 수요가 연 7~9% 성장 중이며, 특히 할로겐프리 난연제 시장은 2023~2030년 CAGR 8.2%로 예측된다(MarketsandMarkets, 2024)."
)

write_line("-시장 규모 및 성장성", size=12, bold=True)
write_para(
    "본 사업의 시장 규모는 TAM(전체 단열재 시장) 약 2.4조 원, SAM(외벽 준불연·불연 단열재) 약 9,100억 원, SOM(친환경 하이브리드 단열보드 신규 수요) 약 1,200억 원으로 산정된다. "
    "5년차 SOM 점유율 8% 달성 시 본 사업의 직접 매출은 약 96억 원 수준이며, ESS·EV 인접 시장 확장 시 누적 매출은 250억 원 이상으로 확대 가능하다."
)
write_line("<표 1> 시장 규모 및 점유 목표", size=12, bold=True, align="center")
write_para(
    "  TAM(국내 단열재 전체) : 약 2.4조 원\n  SAM(외벽 준불연·불연 단열재) : 약 9,100억 원\n  SOM(친환경 하이브리드 단열보드) : 약 1,200억 원\n  5년차 목표 점유율 : SOM의 8%\n  출처 : 한국건설기술연구원(2024), MarketsandMarkets(2024) 재가공"
)

write_line("-경쟁사 및 경쟁 기술 분석", size=12, bold=True)
write_para(
    "국내 경쟁사로는 K사(미네랄울계), B사(페놀폼계), L사(EPS 난연그레이드)가 있으며, 모두 단일 코어 구조로 단열성능 또는 난연성능 중 하나의 한계를 보인다. "
    "본 제품은 다층 구조와 4중 난연 시너지로 두 성능을 동시에 충족한다는 차별성을 가진다. "
    "특히 흡습률(JIS A 9521 기준) 및 시공 절단성에서 페놀폼 단독 대비 우위에 있다."
)
write_line("<표 2> 경쟁사 비교", size=12, bold=True, align="center")
write_para(
    "  당사(Safe-Core™) : λ 0.022 W/m·K, pHRR 220 kW/m², 흡습률 1.2%, 가격 200/단위\n"
    "  K사(미네랄울)    : λ 0.034 W/m·K, pHRR 80 kW/m²,  흡습률 0.5%, 가격 130/단위\n"
    "  B사(페놀폼)      : λ 0.020 W/m·K, pHRR 280 kW/m², 흡습률 2.5%, 가격 180/단위\n"
    "  L사(EPS 난연)    : λ 0.036 W/m·K, pHRR 950 kW/m², 흡습률 0.8%, 가격 100/단위"
)
insert_picture(IMG_CONE_PHOTO, caption="<그림 4> 콘칼로리미터(KS F ISO 5660-1) — 본 제품 성능 검증 장비")

write_line("-SWOT 분석", size=12, bold=True)
write_para(
    "  Strengths : 4중 난연 시너지 처방, KFPA 시험 노하우, 친환경(할로겐프리)\n"
    "  Weaknesses : 신생 브랜드 인지도 부족, 양산 초기 단가 경쟁력 미흡\n"
    "  Opportunities : 외벽 불연·준불연 의무화 확대, EU·중동 친환경 인증 수요 증가\n"
    "  Threats : 대형사의 후발 진입, 원자재(APP·실리콘) 가격 변동, 화평법 규제 강화"
)

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 4장. 추진 일정 및 단계별 계획
# ──────────────────────────────────────────────────────────────────────────
write_line("4. 추진 일정 및 단계별 계획", size=12, bold=True)
write_line("", size=12)

write_line("-단계별 로드맵", size=12, bold=True)
write_para(
    "1차연도(2026)는 처방 최적화 및 핵심 특허 출원, 파일럿 라인(연 3,000 m³) 구축, KS F ISO 5660-1·1182 인증 획득을 목표로 한다. "
    "2차연도(2027)는 KS F 8414 실대형 모형 인증 및 시제품 양산(연 20,000 m³), 상위 5개 시공사 시방 등재를 추진한다. "
    "3차연도(2028)는 청주 본공장 준공(연 50,000 m³)과 LH·SH 공공발주 본격 진입을 목표로 한다. "
    "4~5차연도(2029~2030)는 동남아 합작법인 설립 및 EU CE·미국 ASTM 인증 획득, 인접 시장(ESS·EV 단열재) 진출을 추진한다."
)
write_line("<표 3> 단계별 로드맵 (간트 요약)", size=12, bold=True, align="center")
write_para(
    "  1차연도 ▌▌▌▌  : 처방·파일럿·인증\n"
    "  2차연도   ▌▌▌▌: 실대형 인증·시방 등재\n"
    "  3차연도     ▌▌▌▌: 본공장 준공·공공 진입\n"
    "  4차연도       ▌▌▌▌: 글로벌 인증·해외 합작\n"
    "  5차연도         ▌▌▌▌: 인접 시장 확장(ESS·EV)"
)

write_line("-주요 마일스톤", size=12, bold=True)
write_para(
    "  Q1 2026 : 핵심 특허 3건 출원, KFPA 콘칼로리미터 사내 시험 통과(pHRR ≤ 250 kW/m²)\n"
    "  Q3 2026 : 파일럿 라인 가동, KS F ISO 5660-1 준불연 등급 획득\n"
    "  Q2 2027 : KS F 8414 실대형 시험 통과, 상위 5개 시공사 PoC 완료\n"
    "  Q4 2027 : Series A 투자(50억 원) 유치\n"
    "  Q3 2028 : 본공장 준공식, 누적 매출 80억 원 돌파\n"
    "  Q2 2029 : EU CE 인증 및 첫 수출 계약\n"
    "  Q4 2030 : 누적 매출 250억 원 및 영업이익률 12% 달성"
)
insert_picture(IMG_NONFLAM_PHOTO, caption="<그림 5> 불연성 시험기(KS F ISO 1182) — 인증 단계 핵심 장비")

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 5장. 조직 구성 및 인력 운영 계획
# ──────────────────────────────────────────────────────────────────────────
write_line("5. 조직 구성 및 인력 운영 계획", size=12, bold=True)
write_line("", size=12)

write_line("-조직도", size=12, bold=True)
write_para(
    "본 사업의 조직은 대표이사 직속의 ① 연구개발본부, ② 생산·품질본부, ③ 영업·마케팅본부, ④ 경영지원본부의 4본부 체제로 구성한다. "
    "연구개발본부는 난연처방팀·복합소재팀·시험인증팀의 3개 팀으로, 생산·품질본부는 양산팀·QC팀·구매팀의 3개 팀으로 운영한다."
)
insert_picture(IMG_ORG_CHART, caption="<그림 6> 조직 구성 참조 (KFPA 조직도 재해석)")

write_line("-단계별 인력 확보 계획", size=12, bold=True)
write_line("<표 4> 연차별 인력 확보 계획", size=12, bold=True, align="center")
write_para(
    "  1차연도 : 연구 4 / 생산 3 / 영업 2 / 지원 1 = 총 10명\n"
    "  2차연도 : 연구 6 / 생산 8 / 영업 4 / 지원 2 = 총 20명\n"
    "  3차연도 : 연구 8 / 생산 18 / 영업 6 / 지원 3 = 총 35명\n"
    "  4차연도 : 연구 10 / 생산 28 / 영업 9 / 지원 5 = 총 52명\n"
    "  5차연도 : 연구 12 / 생산 40 / 영업 13 / 지원 7 = 총 72명"
)

write_line("-역량 강화 방안", size=12, bold=True)
write_para(
    "산학협력으로 성균관대·한양대 화학공학·고분자공학과와 공동연구실을 설립하고, KFPA·KCL과 시험·인증 표준화 협력 협약(MOU)을 체결한다. "
    "사내 교육은 분기별 ISO 5660-1·8414 시험 절차 워크숍과 화재공학(Fire Science) 입문 과정을 운영하여 전 직원의 기술 이해도를 균등화한다. "
    "외부 전문가 자문단(난연제·고분자·건축법규 각 1인)을 위촉하여 분기 1회 자문 회의를 진행한다."
)

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 6장. 재무 계획 및 자금 조달
# ──────────────────────────────────────────────────────────────────────────
write_line("6. 재무 계획 및 자금 조달", size=12, bold=True)
write_line("", size=12)

write_line("-매출 및 손익 추정", size=12, bold=True)
write_line("<표 5> 3개년 추정 손익계산서 (단위 : 백만 원)", size=12, bold=True, align="center")
write_para(
    "  구  분         |  1차연도 |  2차연도 |  3차연도\n"
    "  매출액         |     800 |    3,500 |    8,000\n"
    "  매출원가       |     680 |    2,520 |    5,440\n"
    "  매출총이익     |     120 |      980 |    2,560\n"
    "  판관비         |     350 |      650 |    1,200\n"
    "  영업이익       |    -230 |      330 |    1,360\n"
    "  영업이익률     |   -28.8%|     9.4% |    17.0%"
)

write_line("-투자 계획", size=12, bold=True)
write_line("<표 6> 3개년 투자 계획 (단위 : 백만 원)", size=12, bold=True, align="center")
write_para(
    "  R&D(처방·시제품·인증)        : 1차 600 / 2차 800 / 3차 900\n"
    "  설비(파일럿·본공장 라인)     : 1차 1,500 / 2차 3,000 / 3차 6,500\n"
    "  운영자금(인건비·마케팅)       : 1차 700 / 2차 1,200 / 3차 1,800\n"
    "  합  계                        : 1차 2,800 / 2차 5,000 / 3차 9,200"
)

write_line("-자금 조달 방안", size=12, bold=True)
write_para(
    "초기 자기자본(공동창업자) 5억 원과 정부지원사업(창업성장기술개발 R&D 5억 원, 소부장 핵심전략기술 자율형 7억 원)을 1차연도에 확보한다. "
    "2차연도에는 Series A 투자 50억 원을 VC 중심으로 유치하고, 3차연도에는 Series B 100억 원과 산업은행·신용보증기금 시설자금 융자 80억 원을 결합하여 본공장 준공 자금을 조달한다. "
    "조달 우선순위는 정부지원→VC→정책금융 순이며, 지분 희석을 최소화하기 위해 R&D 단계는 비희석성 자금 비중을 70% 이상으로 유지한다."
)
insert_picture(IMG_GAS_PHOTO, caption="<그림 7> 가스유해성 시험기(KS F 2271) — 인증 비용 핵심 항목")

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 7장. 리스크 분석 및 대응 방안
# ──────────────────────────────────────────────────────────────────────────
write_line("7. 리스크 분석 및 대응 방안", size=12, bold=True)
write_line("", size=12)

write_line("-기술 리스크", size=12, bold=True)
write_para(
    "  처방 안정성 저하 리스크 : 양산 환경에서 입자 크기·습도 변화로 pHRR 재현성이 낮아질 가능성. 대응방안은 SPC(통계적 공정관리) 도입 및 In-line FT-IR 모니터링으로 변동요인을 실시간 제어한다.\n"
    "  인증 미통과 리스크 : KS F 8414 실대형 시험 1차 미통과 가능성. 대응방안은 KFPA·KCL 사전 예비시험을 2회 이상 거치고, 외벽 시스템 단위로 다중 처방 트랙(Plan A·B)을 병렬 진행한다."
)

write_line("-시장 리스크", size=12, bold=True)
write_para(
    "  대형 경쟁사 진입 리스크 : 글로벌 단열재 기업의 한국 시장 가격공세. 대응방안은 친환경·재활용 차별화 메시지와 공공발주 시방 등재로 진입장벽을 구축한다.\n"
    "  규제 변동 리스크 : 외벽 마감재 등급 기준 추가 강화. 대응방안은 등급 상향(준불연→불연) 처방 후속 R&D를 1년 선행 운영한다."
)

write_line("-재무 리스크", size=12, bold=True)
write_para(
    "  본공장 투자 부담 리스크 : 3차연도 65억 원 설비투자에 따른 현금흐름 악화. 대응방안은 시설자금 정책융자(산업은행·중진공) 및 리스 방식 분산 조달로 단년도 부담을 30% 이내로 제한한다.\n"
    "  원자재 가격 변동 리스크 : APP·실리콘·Al(OH)₃ 가격 급등. 대응방안은 2개 이상 공급선 이원화와 6개월 단위 선구매 헷지 계약을 운영한다."
)

write_line("-규제 리스크", size=12, bold=True)
write_para(
    "  화평법·REACH 등록 리스크 : 신규 난연제 조성의 등록 지연. 대응방안은 기존 등록물질(APP·MEL·실리콘) 위주로 처방을 구성하고, 신규 물질은 학술용 소량 등록으로 우선 대응한다.\n"
    "  건축법령 해석 차이 리스크 : 지자체별 외벽 마감재 적용 해석 편차. 대응방안은 국토안전관리원·KFPA 공동 기술 가이드 발간으로 해석 통일성을 제고한다."
)
insert_picture(IMG_PEN_FIRE, caption="<그림 8> 내화충전구조 시험 — 규제 강화 흐름의 단면")

page_break()

# ──────────────────────────────────────────────────────────────────────────
# 8장. 기대 효과 및 결론
# ──────────────────────────────────────────────────────────────────────────
write_line("8. 기대 효과 및 결론", size=12, bold=True)
write_line("", size=12)

write_line("-경제적 효과", size=12, bold=True)
write_para(
    "5년차 누적 매출 250억 원, 누적 고용 72명, 협력사 12개사를 통한 간접 고용 약 140명을 창출한다. "
    "수입 의존 페놀폼·할로겐 난연제 일부를 국산 친환경 처방으로 대체하여 연 약 80억 원의 수입대체 효과가 기대된다."
)

write_line("-사회적·환경적 효과", size=12, bold=True)
write_para(
    "본 제품 적용 시 외벽 화재로 인한 인명피해 발생률을 통계적으로 30% 이상 저감할 수 있을 것으로 추정되며(KFPA 사고통계 기반 시뮬레이션), "
    "할로겐프리·재활용 가능 구조로 폐기 시 다이옥신·HBr 등 유해가스 방출이 사실상 0에 수렴한다. "
    "또한 단열성능 향상에 따른 건축물 운영단계 에너지 사용량을 m²당 약 12% 절감하여 탄소중립 목표에 기여한다."
)

write_line("-기술적 파급 효과", size=12, bold=True)
write_para(
    "본 사업의 4중 난연 처방 플랫폼은 ESS 셀 격벽재, 전기차 배터리 모듈 단열재, 데이터센터 케이블 트레이용 난연 슬리브 등 인접 응용으로 확장 가능하다. "
    "또한 콘칼로리미터 데이터를 FDS(Fire Dynamics Simulator) 입력값으로 변환하는 디지털 트윈 모듈은 화재 시뮬레이션 SaaS 사업 모델로의 확장도 가능하다."
)

write_line("-결론", size=12, bold=True)
write_para(
    "본 사업은 반복적 외벽 화재로 야기된 사회적 비용을 화학공학적 원리에 기반한 4중 난연 시너지 처방으로 근본 해결한다는 명확한 가치제안을 가진다. "
    "KFPA 현장에서 검증된 표준시험 데이터, 친환경·재활용 가능한 처방 구조, 단계적 자금 조달 전략의 3박자가 결합되어 기술·시장·재무 측면 모두에서 사업화 타당성이 높다. "
    "본 계획서에 따른 단계별 실행을 통해 5년 내 250억 원 매출과 국내 외벽 단열재 시장의 친환경 표준 기업으로의 자리매김을 달성할 수 있을 것으로 판단된다."
)

# ──────────────────────────────────────────────────────────────────────────
# 페이지번호 - N - 형태로 하단 중앙
# ──────────────────────────────────────────────────────────────────────────
try:
    hwp.HAction.GetDefault("PageNumPos", hwp.HParameterSet.HPageNumPos.HSet)
    pn = hwp.HParameterSet.HPageNumPos
    pn.Pos = 8        # 가운데 아래
    pn.NumberFormat = 0  # 1, 2, 3
    pn.SideChar = 1   # 양쪽 - -
    hwp.HAction.Execute("PageNumPos", pn.HSet)
except Exception as e:
    print("PageNum err:", e)

# ──────────────────────────────────────────────────────────────────────────
# 전체 글자 크기 12pt 강제 (표지 제외)
# 표지는 첫 페이지로 가정 — 첫 페이지나눔 이전 영역은 건드리지 않는다.
# 안전을 위해 본 단계는 생략하고, 작성 시 12pt로 명시했으니 통과.
# ──────────────────────────────────────────────────────────────────────────

# 저장
hwp.SaveAs(OUT, "HWP", "")
print("Saved:", OUT)

hwp.Quit()
