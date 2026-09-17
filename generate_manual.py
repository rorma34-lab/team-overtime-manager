import os
import sys
import shutil
from pathlib import Path

# 콘솔 출력 인코딩 설정 (Windows cp949 방지)
if sys.stdout:
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

OUTPUT_DIR = Path(__file__).resolve().parent / "downloads"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
IMG_DIR = OUTPUT_DIR / "illustrations"
IMG_DIR.mkdir(parents=True, exist_ok=True)

PPTX_PATH = OUTPUT_DIR / "Overtime_System_Manual.pptx"
USER_PPTX_PATH = OUTPUT_DIR / "Overtime_User_Manual.pptx"
ADMIN_PPTX_PATH = OUTPUT_DIR / "Overtime_Admin_Manual.pptx"

# 폰트 로드 헬퍼 (맑은 고딕 또는 기본)
def get_font(size, bold=False):
    try:
        font_name = "malgunbd.ttf" if bold else "malgun.ttf"
        return ImageFont.truetype(font_name, size)
    except Exception:
        try:
            return ImageFont.truetype("malgun.ttf", size)
        except Exception:
            try:
                return ImageFont.truetype("arial.ttf", size)
            except Exception:
                return ImageFont.load_default()

CANVAS_W = 1160
CANVAS_H = 720

# ----------------- 지시선 및 사용자용 친절한 말풍선 핀 헬퍼 -----------------

def draw_leader_pin(d, target_xy, pin_xy, badge_num, label_text, color="#ea580c", text_color="#ffffff", sub_hint=None):
    """
    UI 목업 위에 선명한 지시선과 눈에 띄는 번호 뱃지, 쉬운 설명 라벨을 그리는 헬퍼
    (캔버스 경계 밖으로 절대 잘리지 않도록 안전 클리핑 적용)
    """
    tx, ty = target_xy
    px, py = pin_xy

    # 1. 지시선 (선명한 3px 고대비 선)
    d.line([(tx, ty), (px, py)], fill=color, width=3)
    # 타겟 포인트 원 (도넛 모양)
    d.ellipse([tx - 6, ty - 6, tx + 6, ty + 6], fill=color, outline="#ffffff", width=2)

    # 2. 텍스트 너비 및 높이 계산
    font = get_font(13, bold=True)
    badge_str = f" {badge_num}  {label_text} "
    try:
        bbox = font.getbbox(badge_str)
        text_w = (bbox[2] - bbox[0]) + 20
    except Exception:
        text_w = len(badge_str) * 14 + 20
    text_h = 28

    # 핀 박스 좌상단 계산 (경계선 자동 클리핑)
    bx1 = px if px >= tx else px - text_w
    bx1 = max(15, min(bx1, CANVAS_W - text_w - 15))
    by1 = py - 14
    by1 = max(15, min(by1, CANVAS_H - text_h - 40))
    bx2 = bx1 + text_w
    by2 = by1 + text_h

    # 부드러운 그림자
    d.rounded_rectangle([bx1 + 2, by1 + 3, bx2 + 2, by2 + 3], radius=8, fill="#cbd5e1")
    # 메인 뱃지 박스
    d.rounded_rectangle([bx1, by1, bx2, by2], radius=8, fill=color, outline="#ffffff", width=2)
    d.text((bx1 + text_w // 2, by1 + text_h // 2), badge_str, font=font, fill=text_color, anchor="mm")

    # 서브 힌트(말풍선 꼬리표)가 있을 경우 아래에 노란색 팁 박스 추가
    if sub_hint:
        s_font = get_font(11, bold=False)
        try:
            s_bbox = s_font.getbbox(sub_hint)
            sw = (s_bbox[2] - s_bbox[0]) + 16
        except Exception:
            sw = len(sub_hint) * 11 + 16
        sh = 22
        sx1 = max(15, min(bx1, CANVAS_W - sw - 15))
        sy1 = by2 + 4
        d.rounded_rectangle([sx1, sy1, sx1 + sw, sy1 + sh], radius=6, fill="#fef08a", outline="#eab308", width=1)
        d.text((sx1 + 8, sy1 + 4), sub_hint, font=s_font, fill="#854d0e")


def draw_bubble_card(d, xy, title, desc, bg="#fef9c3", border="#facc15", text_c="#713f12"):
    """사용자용 귀여운 꿀팁 말풍선 카드"""
    x1, y1, x2, y2 = xy
    d.rounded_rectangle([x1 + 2, y1 + 3, x2 + 2, y2 + 3], radius=10, fill="#e2e8f0")
    d.rounded_rectangle([x1, y1, x2, y2], radius=10, fill=bg, outline=border, width=2)
    d.text((x1 + 20, y1 + 12), title, font=get_font(14, bold=True), fill=text_c)
    d.text((x1 + 20, y1 + 36), desc, font=get_font(12, bold=False), fill=text_c)


# ----------------- UI 목업 삽화 이미지 생성기들 (고해상도 & 친절 지시선) -----------------

def create_mockup_login():
    """삽화 1: 사번 입력 화면 목업 (지시선 및 번호 배지)"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f1f5f9")
    d = ImageDraw.Draw(img)

    # 상단 브랜드 헤더 영역
    d.rounded_rectangle([180, 25, 980, 75], radius=10, fill="#1e3a8a")
    d.text((210, 39), "★ 스마트 특근 관리 시스템", font=get_font(16, bold=True), fill="#ffffff")
    d.rounded_rectangle([870, 33, 955, 67], radius=12, fill="#3b82f6")
    d.text((912, 50), "v1.30", font=get_font(13, bold=True), fill="#ffffff", anchor="mm")

    # 카드 배경
    d.rounded_rectangle([180, 95, 980, 585], radius=16, fill="#ffffff", outline="#cbd5e1", width=2)
    # 아이콘 원
    d.ellipse([540, 115, 620, 195], fill="#eff6ff", outline="#3b82f6", width=3)
    d.text((580, 155), "TIME", font=get_font(18, bold=True), fill="#2563eb", anchor="mm")

    # 타이틀
    d.text((580, 225), "사원번호로 입장", font=get_font(24, bold=True), fill="#0f172a", anchor="mm")
    d.text((580, 258), "사원번호 6자리 숫자를 입력하시면 본인 화면으로 안전하게 입장합니다.", font=get_font(14), fill="#64748b", anchor="mm")

    # 입력 필드
    d.text((280, 295), "사원번호 (6자리 숫자 입력)", font=get_font(14, bold=True), fill="#1e293b")
    d.rounded_rectangle([280, 320, 880, 380], radius=10, fill="#ffffff", outline="#3b82f6", width=3)
    d.text((310, 340), "123456", font=get_font(19, bold=True), fill="#0f172a")

    # 버튼
    d.rounded_rectangle([280, 405, 880, 470], radius=10, fill="#2563eb")
    d.text((580, 437), "입장하기 →", font=get_font(17, bold=True), fill="#ffffff", anchor="mm")

    # 하단 신규 등록 팝업 힌트 박스
    d.rounded_rectangle([280, 490, 880, 555], radius=8, fill="#eff6ff", outline="#93c5fd", width=2)
    d.text((580, 510), "[미등록 사번은 즉시 회원등록 창이 열립니다]", font=get_font(13, bold=True), fill="#1e40af", anchor="mm")
    d.text((580, 532), "성명과 소속 부서를 선택하면 즉시 등록되어 대시보드로 이동합니다.", font=get_font(12), fill="#3b82f6", anchor="mm")

    # --- [지시선 Leader Lines & Callouts] ---
    draw_leader_pin(d, (912, 50), (980, 20), "①", "v1.30 최신 버전 배지", color="#0284c7")
    draw_leader_pin(d, (580, 350), (30, 350), "②", "6자리 숫자 사번 입력창", color="#ea580c", sub_hint="[안내] 6자리 숫자 필수")
    draw_leader_pin(d, (880, 437), (960, 437), "③", "[입장하기] 원클릭!", color="#16a34a", sub_hint="즉시 대시보드 입장")
    draw_leader_pin(d, (580, 520), (30, 520), "④", "신규 사원 즉시 등록", color="#9333ea", sub_hint="[안내] 소속 부서 선택")

    # 하단 꿀팁 카드
    draw_bubble_card(d, (180, 605, 980, 695), "[사원번호 간편 입장 안내]",
                     "1. 사원번호는 6자리 숫자로 입력해야 정상 처리됩니다. (총괄관리자 제외)\n"
                     "2. 미등록 사번은 성명과 소속 부서를 선택하면 즉시 등록되어 로그인됩니다.")

    path = IMG_DIR / "mockup_login.png"
    img.save(path)
    return str(path)


def create_mockup_calendar_apply():
    """삽화 2: 특근 신청 달력 및 폼 목업 (지시선 포함)"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f1f5f9")
    d = ImageDraw.Draw(img)

    # 전체 카드
    d.rounded_rectangle([150, 20, 1010, 590], radius=14, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((180, 40), "[신규 특근 신청하기] (달력 콕 누르면 끝!)", font=get_font(20, bold=True), fill="#0f172a")
    d.text((180, 68), "달력에서 일할 날짜를 콕 누르고, 사유 적고 [특근 신청 저장]만 누르면 돼요!", font=get_font(13), fill="#64748b")

    # 좌측: 입력 폼
    d.text((180, 100), "1. 특근 종류 3가지 중 하나 골라요", font=get_font(13, bold=True), fill="#1e293b")
    # 칩들
    d.rounded_rectangle([180, 125, 290, 168], radius=8, fill="#eff6ff", outline="#2563eb", width=2)
    d.text((235, 146), "● 일반휴일", font=get_font(12, bold=True), fill="#2563eb", anchor="mm")
    d.rounded_rectangle([300, 125, 410, 168], radius=8, fill="#ffffff", outline="#cbd5e1")
    d.text((355, 146), "대체근무", font=get_font(12), fill="#64748b", anchor="mm")
    d.rounded_rectangle([420, 125, 530, 168], radius=8, fill="#ffffff", outline="#cbd5e1")
    d.text((475, 146), "법정휴일", font=get_font(12), fill="#64748b", anchor="mm")

    # 날짜 입력 박스
    d.text((180, 185), "시작일 (달력 누르면 자동)", font=get_font(12), fill="#64748b")
    d.rounded_rectangle([180, 205, 345, 245], radius=6, fill="#f8fafc", outline="#cbd5e1")
    d.text((195, 218), "2026-09-12", font=get_font(13, bold=True), fill="#0f172a")

    d.text((365, 185), "종료일 (연속 근무 가능)", font=get_font(12), fill="#64748b")
    d.rounded_rectangle([365, 205, 530, 245], radius=6, fill="#f8fafc", outline="#cbd5e1")
    d.text((380, 218), "2026-09-13", font=get_font(13, bold=True), fill="#0f172a")

    # 사유 및 프로젝트
    d.text((180, 260), "2. 프로젝트 번호 & 장소 (가이드 예시)", font=get_font(12, bold=True), fill="#1e293b")
    d.rounded_rectangle([180, 280, 530, 318], radius=6, fill="#ffffff", outline="#cbd5e1")
    d.text((195, 292), "BT2601-L1  |  본사5층", font=get_font(12), fill="#334155")

    d.text((180, 335), "3. 일하는 이유 (간단히 작성)", font=get_font(12, bold=True), fill="#1e293b")
    d.rounded_rectangle([180, 355, 530, 420], radius=6, fill="#ffffff", outline="#cbd5e1")
    d.text((195, 375), "프로그램 개발", font=get_font(13), fill="#334155")

    # 큰 신청 버튼
    d.rounded_rectangle([180, 445, 530, 505], radius=8, fill="#10b981")
    d.text((355, 475), "💾 특근 신청 저장하기 (신청 끝!)", font=get_font(16, bold=True), fill="#ffffff", anchor="mm")

    # 우측: 달력 목업
    d.rounded_rectangle([560, 100, 985, 505], radius=10, fill="#f8fafc", outline="#cbd5e1", width=2)
    d.text((585, 118), "2026년 9월 달력", font=get_font(16, bold=True), fill="#0f172a")
    d.text((770, 120), "▶ 날짜를 마우스/손가락으로 콕!", font=get_font(11, bold=True), fill="#2563eb")

    # 달력 헤더
    days = ["일", "월", "화", "수", "목", "금", "토"]
    for i, day in enumerate(days):
        col_c = "#ef4444" if i == 0 else ("#2563eb" if i == 6 else "#475569")
        d.text((595 + i * 55, 150), day, font=get_font(12, bold=True), fill=col_c)

    # 달력 날짜 셀들
    for row in range(4):
        for col in range(7):
            day_num = row * 7 + col - 1
            if 1 <= day_num <= 30:
                cx = 580 + col * 55
                cy = 180 + row * 65
                is_selected = (day_num in [12, 13])
                if is_selected:
                    d.rounded_rectangle([cx, cy, cx + 48, cy + 56], radius=6, fill="#dbeafe", outline="#2563eb", width=2)
                    d.text((cx + 6, cy + 6), str(day_num), font=get_font(12, bold=True), fill="#1d4ed8")
                    d.rounded_rectangle([cx + 4, cy + 30, cx + 44, cy + 50], radius=4, fill="#2563eb")
                    d.text((cx + 24, cy + 40), "선택", font=get_font(10, bold=True), fill="#ffffff", anchor="mm")
                else:
                    d.rounded_rectangle([cx, cy, cx + 48, cy + 56], radius=6, fill="#ffffff", outline="#e2e8f0")
                    num_c = "#ef4444" if col == 0 else ("#2563eb" if col == 6 else "#334155")
                    d.text((cx + 6, cy + 6), str(day_num), font=get_font(11), fill=num_c)

    # 하단 꿀팁 말풍선
    draw_bubble_card(d, (150, 605, 1010, 695), "[특근 신청 핵심 가이드 (v1.32)!]",
                     "1. 주말(토/일)에 일할 땐 [일반휴일] 버튼 누르고, 달력에서 일할 날짜 콕 찍고 [저장] 누르면 끝나요!\n"
                     "2. 동일한 날짜는 중복 신청이 자동으로 방지되므로 안심하고 신청할 수 있어요!\n"
                     "3. 상단 [🔄 새로고침] 버튼을 누르면 모바일이나 PC 어디서나 실시간 최신 정보로 1초 만에 맞춰져요!")

    # --- [지시선 Leader Lines & Callouts] ---
    draw_leader_pin(d, (235, 146), (30, 146), "①", "특근 종류 콕 누르기", color="#2563eb", sub_hint="주말엔 일반휴일!")
    draw_leader_pin(d, (930, 280), (1010, 280), "②", "달력에서 날짜 콕 찍기", color="#0284c7", sub_hint="파란색으로 변해요!")
    draw_leader_pin(d, (355, 385), (30, 385), "③", "무슨 일 하는지 적기", color="#ea580c")
    draw_leader_pin(d, (530, 475), (1010, 475), "④", "[특근 저장] 누르면 끝!", color="#10b981", sub_hint="관리자님께 바로 전송")

    path = IMG_DIR / "mockup_apply.png"
    img.save(path)
    return str(path)


def create_mockup_my_records():
    """삽화 3: 내 내역 조회 및 승인 상태 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f1f5f9")
    d = ImageDraw.Draw(img)

    d.rounded_rectangle([150, 20, 1010, 590], radius=14, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((180, 40), "[내가 신청한 특근 확인하기] (초록색 도장 쾅!)", font=get_font(20, bold=True), fill="#0f172a")
    d.text((180, 68), "관리자님이 승인해줬는지, 대체휴일로 쉰 날은 언제인지 한눈에 쏙 보여요!", font=get_font(13), fill="#64748b")

    # 내역 카드 1: 확인완료 & 대체휴일 사용건
    d.rounded_rectangle([180, 105, 980, 235], radius=10, fill="#ffffff", outline="#cbd5e1", width=2)
    d.rectangle([180, 105, 190, 235], fill="#10b981") # 초록색 띠
    d.text((210, 120), "2026-09-12 (토) ~ 2026-09-13 (일)", font=get_font(17, bold=True), fill="#0f172a")
    d.rounded_rectangle([530, 115, 605, 143], radius=6, fill="#dcfce7", outline="#86efac")
    d.text((567, 129), "일반휴일", font=get_font(11, bold=True), fill="#15803d", anchor="mm")

    # 초록색 승인완료 배지
    d.rounded_rectangle([780, 115, 955, 145], radius=6, fill="#10b981")
    d.text((867, 130), "✓ 확인완료 (승인도장)", font=get_font(12, bold=True), fill="#ffffff", anchor="mm")

    d.text((210, 155), "• 사유: 1공장 제어설비 긴급 점검  |  장소: 1공장 제어실  |  일수: 2일", font=get_font(13), fill="#475569")

    # 대체휴일 배지 (하이라이트)
    d.rounded_rectangle([210, 185, 560, 220], radius=6, fill="#fef3c7", outline="#f59e0b", width=2)
    d.text((225, 196), "[휴가] 대체휴일 쉰 날: 2026-09-25 (1.0일 쉼)", font=get_font(12, bold=True), fill="#b45309")

    # 수정/삭제 버튼
    d.rounded_rectangle([805, 185, 860, 218], radius=6, fill="#f1f5f9", outline="#cbd5e1")
    d.text((832, 201), "수정", font=get_font(11, bold=True), fill="#475569", anchor="mm")
    d.rounded_rectangle([870, 185, 925, 218], radius=6, fill="#fee2e2", outline="#fca5a5")
    d.text((897, 201), "삭제", font=get_font(11, bold=True), fill="#dc2626", anchor="mm")

    # 내역 카드 2: 승인대기건
    d.rounded_rectangle([180, 255, 980, 370], radius=10, fill="#ffffff", outline="#cbd5e1", width=2)
    d.rectangle([180, 255, 190, 370], fill="#f59e0b") # 노란색 띠
    d.text((210, 270), "2026-09-20 (일)", font=get_font(17, bold=True), fill="#0f172a")
    d.rounded_rectangle([390, 265, 465, 293], radius=6, fill="#eff6ff", outline="#93c5fd")
    d.text((427, 279), "대체근무", font=get_font(11, bold=True), fill="#1d4ed8", anchor="mm")

    # 노란색 승인대기 배지
    d.rounded_rectangle([780, 265, 955, 295], radius=6, fill="#fef3c7", outline="#f59e0b")
    d.text((867, 280), "● 승인대기 (검토중)", font=get_font(12, bold=True), fill="#b45309", anchor="mm")

    d.text((210, 310), "• 사유: 전장배선 라인 사전 포설  |  장소: 배선실  |  일수: 1일", font=get_font(13), fill="#475569")

    # 꿀팁 말풍선
    draw_bubble_card(d, (150, 400, 1010, 570), "[내 신청 내역 확인 가이드!]",
                     "1. 초록색 배지(✓ 확인완료)가 있으면 관리자 선생님이 승인 도장을 쾅 찍어준 거예요!\n"
                     "2. 노란색 [휴가] 배지는 '특근 대신 평일에 하루 쉰 날'을 친절하게 알려주는 표시예요!\n"
                     "3. 날짜가 바뀌었으면 [수정], 취소할 땐 [삭제]를 누르면 끝나요!")

    # --- [지시선 Leader Lines & Callouts] ---
    draw_leader_pin(d, (867, 130), (980, 60), "①", "초록색 승인 완료 도장!", color="#10b981", sub_hint="관리자님 확인 완료")
    draw_leader_pin(d, (385, 202), (30, 202), "②", "대체휴일 쉰 날 표시", color="#f59e0b", sub_hint="내가 쉰 날짜 확인!")
    draw_leader_pin(d, (897, 201), (980, 201), "③", "언제든 [수정]/[삭제]", color="#dc2626")
    draw_leader_pin(d, (867, 280), (980, 280), "④", "노란색 승인대기 표시", color="#ea580c", sub_hint="관리자님 검토 중")

    path = IMG_DIR / "mockup_my.png"
    img.save(path)
    return str(path)


def create_mockup_admin_calendar():
    """삽화 4: 관리자 모드 - 달력 및 일괄 승인 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f1f5f9")
    d = ImageDraw.Draw(img)

    d.rounded_rectangle([150, 20, 1010, 590], radius=14, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((180, 35), "[관리자 모드] 달력 보고 도장 쾅쾅! (원클릭 승인)", font=get_font(20, bold=True), fill="#0f172a")

    # 상단 부서 선택 칩
    d.text((180, 68), "우리 팀 콕 찍어보기:", font=get_font(13, bold=True), fill="#64748b")
    teams = ["[전체보기]", "제어실", "전장배선팀", "전장설계팀", "PLC제어팀"]
    for i, t in enumerate(teams):
        bg = "#2563eb" if i == 1 else "#f1f5f9"
        tc = "#ffffff" if i == 1 else "#475569"
        d.rounded_rectangle([320 + i * 125, 62, 435 + i * 125, 92], radius=6, fill=bg)
        d.text((377 + i * 125, 77), t, font=get_font(11, bold=True), fill=tc, anchor="mm")

    # 달력 영역
    d.rounded_rectangle([180, 110, 580, 420], radius=10, fill="#f8fafc", outline="#cbd5e1")
    d.text((195, 125), "2026년 9월 우리 팀 특근 달력", font=get_font(14, bold=True), fill="#0f172a")

    # 특정 날짜(9월 12일) 셀 강조
    d.rounded_rectangle([360, 190, 500, 280], radius=8, fill="#eff6ff", outline="#2563eb", width=2)
    d.text((370, 200), "12 (토)", font=get_font(13, bold=True), fill="#1d4ed8")
    d.rounded_rectangle([368, 225, 492, 248], radius=4, fill="#10b981")
    d.text((430, 236), "정진규(특근)", font=get_font(11, bold=True), fill="#ffffff", anchor="mm")
    d.rounded_rectangle([368, 252, 492, 275], radius=4, fill="#f59e0b")
    d.text((430, 263), "[휴가]김철수", font=get_font(11, bold=True), fill="#ffffff", anchor="mm")

    # 우측: 당일 상세 명단 및 일괄 확인
    d.rounded_rectangle([600, 110, 980, 420], radius=10, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((620, 125), "9월 12일 (토) 근무자 명단", font=get_font(14, bold=True), fill="#0f172a")

    # 일괄 확인 버튼 (빅 버튼)
    d.rounded_rectangle([620, 155, 960, 200], radius=8, fill="#10b981")
    d.text((790, 177), "✓ [당일 전원 일괄 확인] (한 번에 승인!)", font=get_font(14, bold=True), fill="#ffffff", anchor="mm")

    # 명단 리스트
    d.rounded_rectangle([620, 215, 960, 265], radius=6, fill="#f8fafc", outline="#e2e8f0")
    d.text((635, 228), "정진규 (실장) - 제어설비 정기점검", font=get_font(12, bold=True), fill="#0f172a")
    d.text((635, 246), "특근분류: 일반휴일 (1일)  |  승인완료", font=get_font(11), fill="#10b981")

    d.rounded_rectangle([620, 275, 960, 325], radius=6, fill="#f8fafc", outline="#e2e8f0")
    d.text((635, 288), "김철수 (선임) - [대체휴일] 쉼", font=get_font(12, bold=True), fill="#b45309")
    d.text((635, 306), "사유: 지난주 특근 대체 휴식  |  1.0일 차감", font=get_font(11), fill="#ea580c")

    # 꿀팁 박스
    draw_bubble_card(d, (150, 440, 1010, 570), "[관리자 선생님을 위한 초간단 승인법!]",
                     "1. 팀원들이 일한 날짜를 달력에서 누르면 오른쪽에 일한 친구 명단이 쫙 나와요!\n"
                     "2. 초록색 [당일 전원 일괄 확인] 버튼을 누르면 그날 일한 모든 친구에게 한 번에 승인 도장이 찍혀요!\n"
                     "3. 대체휴일로 쉬는 친구도 달력에 노란색으로 보여서 누가 쉬는지 한눈에 알 수 있어요!")

    # 지시선
    draw_leader_pin(d, (377 + 125, 77), (30, 77), "①", "우리 팀만 쏙 골라보기", color="#2563eb")
    draw_leader_pin(d, (430, 263), (30, 263), "②", "쉬는 친구도 달력에 표시", color="#f59e0b")
    draw_leader_pin(d, (790, 177), (980, 80), "③", "[당일 전원 일괄 확인] 쾅!", color="#10b981", sub_hint="한 번에 전원 승인")
    draw_leader_pin(d, (790, 240), (980, 240), "④", "당일 일한 상세 내용 확인", color="#0284c7")

    path = IMG_DIR / "mockup_admin_cal.png"
    img.save(path)
    return str(path)


def create_mockup_user_mgmt():
    """삽화 5: 관리자 모드 - 팀원 및 소속팀 관리 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f1f5f9")
    d = ImageDraw.Draw(img)

    d.rounded_rectangle([150, 20, 1010, 590], radius=14, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((180, 35), "[팀원과 소속팀 척척 관리하기]", font=get_font(20, bold=True), fill="#0f172a")

    # 상단 관리 버튼들
    d.rounded_rectangle([180, 75, 380, 115], radius=6, fill="#f59e0b")
    d.text((280, 95), "[팀] [소속팀(부서) 관리]", font=get_font(13, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([395, 75, 570, 115], radius=6, fill="#2563eb")
    d.text((482, 95), "📥 엑셀 팀원 대량 등록", font=get_font(12, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([585, 75, 760, 115], radius=6, fill="#10b981")
    d.text((672, 95), "+ 신규 팀원 1명 추가", font=get_font(12, bold=True), fill="#ffffff", anchor="mm")

    # 소속팀 관리 모달 팝업 미리보기 (우측 상단 겹침 효과)
    d.rounded_rectangle([610, 130, 980, 330], radius=10, fill="#ffffff", outline="#f59e0b", width=3)
    d.text((630, 145), "소속팀 관리 (새로운 팀 만들기)", font=get_font(14, bold=True), fill="#b45309")
    d.text((630, 172), "새 팀 이름 입력:", font=get_font(12), fill="#475569")
    d.rounded_rectangle([630, 195, 860, 230], radius=6, fill="#ffffff", outline="#cbd5e1")
    d.text((640, 207), "PLC제어팀", font=get_font(12, bold=True), fill="#0f172a")
    d.rounded_rectangle([870, 195, 960, 230], radius=6, fill="#f59e0b")
    d.text((915, 212), "+ 추가", font=get_font(12, bold=True), fill="#ffffff", anchor="mm")
    d.text((630, 245), "현재 등록된 팀 목록: (삭제된 팀은 안 나와요!)", font=get_font(11, bold=True), fill="#334155")
    d.text((630, 268), "• 제어실   • 전장배선팀   • 전장설계팀   • PLC제어팀", font=get_font(12, bold=True), fill="#2563eb")
    d.text((630, 295), "★ 삭제된 팀은 다시 부활하지 않고 깔끔하게 정리돼요!", font=get_font(11, bold=True), fill="#10b981")

    # 좌측: 팀원 테이블 목업
    d.rounded_rectangle([180, 130, 590, 420], radius=10, fill="#f8fafc", outline="#cbd5e1")
    d.rectangle([180, 130, 590, 165], fill="#e2e8f0")
    d.text((195, 142), "사번", font=get_font(11, bold=True), fill="#334155")
    d.text((275, 142), "이름", font=get_font(11, bold=True), fill="#334155")
    d.text((345, 142), "소속팀", font=get_font(11, bold=True), fill="#334155")
    d.text((440, 142), "관리자권한", font=get_font(11, bold=True), fill="#334155")
    d.text((530, 142), "관리", font=get_font(11, bold=True), fill="#334155")

    members = [
        ("ADMIN01", "슈퍼관리자", "제어실", "총괄 ON", "수정"),
        ("113019", "정진규", "제어실", "OFF", "수정/삭제"),
        ("2024001", "김철수", "기술연구팀", "OFF", "수정/삭제"),
        ("2026002", "이영희", "전장배선팀", "OFF", "수정/삭제")
    ]
    for idx, (m_id, m_name, m_team, m_adm, m_act) in enumerate(members):
        my = 175 + idx * 45
        d.text((195, my), m_id, font=get_font(11), fill="#0f172a")
        d.text((275, my), m_name, font=get_font(11, bold=True), fill="#0f172a")
        d.text((345, my), m_team, font=get_font(11), fill="#2563eb")
        d.text((445, my), m_adm, font=get_font(11), fill="#16a34a" if "ON" in m_adm else "#64748b")
        d.rounded_rectangle([520, my - 4, 580, my + 20], radius=4, fill="#fee2e2")
        d.text((550, my + 8), m_act, font=get_font(9), fill="#dc2626", anchor="mm")

    # 꿀팁 카드
    draw_bubble_card(d, (150, 440, 1010, 570), "[소속팀 및 팀원 관리 가이드!]",
                     "1. [[팀] 소속팀 관리]에서 새로운 팀 이름을 넣으면 회원가입 창에 바로 나타나요!\n"
                     "2. 삭제한 팀은 다시 생겨나지 않도록 완벽하게 정리했어요!\n"
                     "3. 팀원이 퇴사했으면 [삭제] 버튼을 눌러 깔끔하게 정리해요!")

    # 지시선
    draw_leader_pin(d, (280, 95), (30, 50), "①", "[소속팀 관리] 버튼", color="#f59e0b")
    draw_leader_pin(d, (915, 212), (980, 160), "②", "새로운 팀 1초 만에 추가", color="#0284c7")
    draw_leader_pin(d, (750, 270), (980, 350), "③", "삭제된 팀은 다시 안 나와요!", color="#10b981")
    draw_leader_pin(d, (550, 220), (30, 220), "④", "팀원 수정 및 삭제", color="#dc2626")

    path = IMG_DIR / "mockup_user_mgmt.png"
    img.save(path)
    return str(path)


def create_mockup_settlement():
    """삽화 6: 진짜 일한 날(실특근) 자동 계산 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f1f5f9")
    d = ImageDraw.Draw(img)

    d.rounded_rectangle([150, 20, 1010, 590], radius=14, fill="#ffffff", outline="#cbd5e1", width=2)

    # 공식 배너 (초록색 예쁜 카드)
    d.rounded_rectangle([180, 30, 980, 100], radius=10, fill="#ecfdf5", outline="#10b981", width=3)
    d.text((200, 42), "★ 컴퓨터가 1초 만에 계산하는 진짜 일한 날(실특근) 공식!", font=get_font(15, bold=True), fill="#065f46")
    d.text((200, 68), "진짜 일한 날 = 일반휴일에 일한 날 - 대체휴일로 쉰 날 (대체근무와 법정휴일은 알아서 쏙 빼줘요!)", font=get_font(13, bold=True), fill="#047857")

    # 팀별 카드 2개
    d.rounded_rectangle([180, 115, 570, 235], radius=10, fill="#ffffff", outline="#cbd5e1", width=2)
    d.rectangle([180, 115, 190, 235], fill="#2563eb")
    d.text((205, 130), "[팀] 제어실", font=get_font(16, bold=True), fill="#0f172a")
    d.text((490, 132), "팀원: 3명", font=get_font(12), fill="#64748b")
    d.rounded_rectangle([200, 155, 550, 185], radius=6, fill="#f8fafc")
    d.text((210, 164), "총신청 7일  |  법정제외 0일  |  일반휴일 7일  |  대휴 0.0일", font=get_font(11), fill="#475569")
    d.text((205, 200), "★ 우리 팀 최종 실특근:", font=get_font(13, bold=True), fill="#065f46")
    d.text((380, 195), "7.0일 (일반 7 - 대휴 0)", font=get_font(15, bold=True), fill="#059669")

    d.rounded_rectangle([590, 115, 980, 235], radius=10, fill="#ffffff", outline="#cbd5e1", width=2)
    d.rectangle([590, 115, 600, 235], fill="#f59e0b")
    d.text((615, 130), "[팀] 기술연구팀", font=get_font(16, bold=True), fill="#0f172a")
    d.text((900, 132), "팀원: 1명", font=get_font(12), fill="#64748b")
    d.rounded_rectangle([610, 155, 960, 185], radius=6, fill="#f8fafc")
    d.text((620, 164), "총신청 6일  |  법정제외 1일  |  일반휴일 5일  |  대휴 1.0일", font=get_font(11), fill="#475569")
    d.text((615, 200), "★ 우리 팀 최종 실특근:", font=get_font(13, bold=True), fill="#065f46")
    d.text((790, 195), "4.0일 (일반 5 - 대휴 1.0)", font=get_font(15, bold=True), fill="#059669")

    # 개인별 정산 명세 테이블
    d.rounded_rectangle([180, 250, 980, 420], radius=10, fill="#ffffff", outline="#cbd5e1", width=2)
    d.rectangle([180, 250, 980, 285], fill="#f1f5f9")
    d.text((195, 262), "사번", font=get_font(11, bold=True), fill="#334155")
    d.text((275, 262), "이름", font=get_font(11, bold=True), fill="#334155")
    d.text((350, 262), "소속팀", font=get_font(11, bold=True), fill="#334155")
    d.text((450, 262), "일반휴일(일)", font=get_font(11, bold=True), fill="#2563eb")
    d.text((565, 262), "대휴사용(일)", font=get_font(11, bold=True), fill="#ea580c")
    d.text((680, 262), "★ 최종실특근", font=get_font(12, bold=True), fill="#059669")
    d.text((810, 262), "계산 방법", font=get_font(11, bold=True), fill="#64748b")

    # 행들
    d.text((195, 305), "113019", font=get_font(11), fill="#0f172a")
    d.text((275, 305), "정진규", font=get_font(11, bold=True), fill="#0f172a")
    d.text((350, 305), "제어실", font=get_font(11), fill="#475569")
    d.text((480, 305), "7", font=get_font(12, bold=True), fill="#2563eb")
    d.text((595, 305), "0.0", font=get_font(12), fill="#64748b")
    d.text((700, 305), "7.0일", font=get_font(13, bold=True), fill="#059669")
    d.text((810, 305), "7 - 0 = 7.0일", font=get_font(11), fill="#059669")

    d.text((195, 350), "2024001", font=get_font(11), fill="#0f172a")
    d.text((275, 350), "김철수", font=get_font(11, bold=True), fill="#0f172a")
    d.text((350, 350), "기술연구팀", font=get_font(11), fill="#475569")
    d.text((480, 350), "5", font=get_font(12, bold=True), fill="#2563eb")
    d.text((595, 350), "1.0", font=get_font(12, bold=True), fill="#ea580c")
    d.text((700, 350), "4.0일", font=get_font(13, bold=True), fill="#059669")
    d.text((810, 350), "5 - 1.0 = 4.0일 (대휴 뺌!)", font=get_font(11), fill="#059669")

    # 엑셀 다운로드 버튼
    d.rounded_rectangle([800, 380, 965, 415], radius=6, fill="#2563eb")
    d.text((882, 397), "📥 엑셀로 내보내기", font=get_font(11, bold=True), fill="#ffffff", anchor="mm")

    # 꿀팁 박스
    draw_bubble_card(d, (150, 440, 1010, 570), "[실특근 정산 및 자동 계산 원리!]",
                     "1. 주말(일반휴일)에 일한 날에서 대체휴일로 쉰 날을 쏙 빼면 진짜 일한 날이 돼요!\n"
                     "2. 김철수 삼촌은 5번 일하고 1번 쉬었으니까: 5 - 1 = 4일 인정!\n"
                     "3. 사람이 계산기 두드릴 필요 없이 컴퓨터가 실수 없이 1초 만에 척척 계산해줘요!")

    # 지시선
    draw_leader_pin(d, (580, 68), (580, 15), "①", "쉬운 뺄셈 공식", color="#10b981")
    draw_leader_pin(d, (380, 195), (30, 195), "②", "팀별 실특근 합계", color="#2563eb")
    draw_leader_pin(d, (700, 350), (980, 330), "③", "개인별 뺄셈 자동 계산", color="#ea580c")
    draw_leader_pin(d, (882, 397), (980, 410), "④", "엑셀 파일로 즉시 저장", color="#0284c7")

    path = IMG_DIR / "mockup_settlement.png"
    img.save(path)
    return str(path)


def create_mockup_excel_27():
    """삽화 7: 요구사항 27 - 엑셀 3대 개편 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f1f5f9")
    d = ImageDraw.Draw(img)

    d.rounded_rectangle([150, 20, 1010, 590], radius=14, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((180, 35), "[엑셀 파일이 3배 더 똑똑해졌어요!] (글자 없이 숫자만 쏙!)", font=get_font(20, bold=True), fill="#0f172a")

    # 시트 탭
    d.rounded_rectangle([180, 75, 410, 105], radius=6, fill="#1e3a8a")
    d.text((295, 90), "📄 [시트1] 휴일일자별_특근현황", font=get_font(11, bold=True), fill="#ffffff", anchor="mm")
    d.rounded_rectangle([420, 75, 650, 105], radius=6, fill="#0284c7")
    d.text((535, 90), "📄 [시트2] 개인별_휴일합산_정산표", font=get_font(11, bold=True), fill="#ffffff", anchor="mm")

    # 시트 1 테이블 목업
    d.rounded_rectangle([180, 115, 980, 275], radius=8, fill="#ffffff", outline="#cbd5e1")
    d.rectangle([180, 115, 980, 145], fill="#1e3a8a")
    d.text((195, 127), "휴일일자", font=get_font(10, bold=True), fill="#ffffff")
    d.text((275, 127), "요일", font=get_font(10, bold=True), fill="#ffffff")
    d.text((315, 127), "성명", font=get_font(10, bold=True), fill="#ffffff")
    d.text((380, 127), "사번", font=get_font(10, bold=True), fill="#ffffff")
    d.text((455, 127), "소속팀", font=get_font(10, bold=True), fill="#ffffff")
    d.text((545, 127), "특근분류", font=get_font(10, bold=True), fill="#ffffff")
    d.text((645, 127), "휴일일수(숫자만!)", font=get_font(10, bold=True), fill="#fef08a")
    d.text((775, 127), "대체휴가일수", font=get_font(10, bold=True), fill="#fed7aa")
    d.text((885, 127), "대휴사용일", font=get_font(10, bold=True), fill="#ffffff")

    # 행 1
    d.text((195, 158), "2026-09-05", font=get_font(10), fill="#0f172a")
    d.text((280, 158), "토", font=get_font(10, bold=True), fill="#2563eb")
    d.text((315, 158), "정진규", font=get_font(10, bold=True), fill="#0f172a")
    d.text((380, 158), "113019", font=get_font(10), fill="#475569")
    d.text((455, 158), "제어실", font=get_font(10), fill="#475569")
    d.text((545, 158), "일반휴일", font=get_font(10), fill="#2563eb")
    d.text((685, 158), "1", font=get_font(13, bold=True), fill="#b45309")
    d.text((815, 158), "0.0", font=get_font(10), fill="#64748b")
    d.text((895, 158), "-", font=get_font(10), fill="#64748b")

    # 행 2
    d.text((195, 192), "2026-09-12", font=get_font(10), fill="#0f172a")
    d.text((280, 192), "토", font=get_font(10, bold=True), fill="#2563eb")
    d.text((315, 192), "김철수", font=get_font(10, bold=True), fill="#0f172a")
    d.text((380, 192), "2024001", font=get_font(10), fill="#475569")
    d.text((455, 192), "기술연구팀", font=get_font(10), fill="#475569")
    d.text((545, 192), "일반휴일", font=get_font(10), fill="#2563eb")
    d.text((685, 192), "1", font=get_font(13, bold=True), fill="#b45309")
    d.text((815, 192), "1.0", font=get_font(11, bold=True), fill="#ea580c")
    d.text((885, 192), "2026-09-25", font=get_font(10, bold=True), fill="#059669")

    # 합계 행
    d.rectangle([180, 220, 980, 250], fill="#fef3c7")
    d.text((195, 230), "합계 (엑셀 수식 자동계산)", font=get_font(10, bold=True), fill="#92400e")
    d.text((660, 230), "=SUM(G4:G5) -> 2", font=get_font(11, bold=True), fill="#b45309")
    d.text((800, 230), "=SUM(H4:H5) -> 1.0", font=get_font(11, bold=True), fill="#ea580c")

    # 시트 2 테이블 목업 (개인별 합산)
    d.rounded_rectangle([180, 290, 980, 420], radius=8, fill="#ffffff", outline="#0284c7")
    d.rectangle([180, 290, 980, 320], fill="#0284c7")
    d.text((195, 302), "사번", font=get_font(10, bold=True), fill="#ffffff")
    d.text((275, 302), "성명", font=get_font(10, bold=True), fill="#ffffff")
    d.text((345, 302), "소속팀", font=get_font(10, bold=True), fill="#ffffff")
    d.text((440, 302), "대체근무(일)", font=get_font(10, bold=True), fill="#ffffff")
    d.text((550, 302), "법정휴일(일)", font=get_font(10, bold=True), fill="#ffffff")
    d.text((660, 302), "일반휴일(일)", font=get_font(10, bold=True), fill="#ffffff")
    d.text((770, 302), "대휴사용(일)", font=get_font(10, bold=True), fill="#ffffff")
    d.text((880, 302), "★ 최종실특근", font=get_font(11, bold=True), fill="#fef08a")

    # 시트 2 행
    d.text((195, 335), "2024001", font=get_font(10), fill="#0f172a")
    d.text((275, 335), "김철수", font=get_font(10, bold=True), fill="#0f172a")
    d.text((345, 335), "기술연구팀", font=get_font(10), fill="#475569")
    d.text((475, 335), "0", font=get_font(10), fill="#64748b")
    d.text((585, 335), "1", font=get_font(10), fill="#64748b")
    d.text((695, 335), "5", font=get_font(12, bold=True), fill="#2563eb")
    d.text((805, 335), "1.0", font=get_font(12, bold=True), fill="#ea580c")
    d.rounded_rectangle([875, 330, 955, 355], radius=4, fill="#dcfce7")
    d.text((915, 342), "4.0일", font=get_font(12, bold=True), fill="#15803d", anchor="mm")

    # 꿀팁 카드
    draw_bubble_card(d, (150, 440, 1010, 570), "[개편된 엑셀 양식 활용 가이드!]",
                     "1. [27-1] 하루하루 일한 날짜 순서대로 차례차례 펼쳐져요!\n"
                     "2. [27-2] '1일' 글자 대신 숫자 '1'만 들어있어서 마우스로 긁기만 해도 합계가 척척 나와요!\n"
                     "3. [27-3] 두 번째 시트에는 친구별로 총 몇 번 일했는지 4가지 휴일수가 완벽히 정리되어 있어요!")

    # 지시선
    draw_leader_pin(d, (230, 158), (30, 158), "①", "[27-1] 날짜순 쫙 전개", color="#0284c7")
    draw_leader_pin(d, (685, 158), (980, 80), "②", "[27-2] 순수 숫자 '1'만 쏙", color="#b45309", sub_hint="수식 합계 자동화")
    draw_leader_pin(d, (600, 335), (30, 335), "③", "[27-3] 개인별 4대 휴일수 합산", color="#16a34a")
    draw_leader_pin(d, (915, 342), (980, 375), "④", "일반휴일 - 대휴 = 실특근", color="#10b981")

    path = IMG_DIR / "mockup_excel_27.png"
    img.save(path)
    return str(path)


# ----------------- 메인 PPTX 빌더 함수들 -----------------

COLOR_PRIMARY = RGBColor(30, 58, 138)     # Deep Royal Navy (#1E3A8A)
COLOR_ACCENT = RGBColor(234, 88, 12)     # Vivid Orange (#EA580C)
COLOR_SUCCESS = RGBColor(16, 185, 129)   # Emerald (#10B981)
COLOR_TEXT_MAIN = RGBColor(15, 23, 42)   # Dark Slate (#0F172A)
COLOR_TEXT_MUTED = RGBColor(100, 116, 139) # Slate Grey (#64748B)
COLOR_BG_CARD = RGBColor(248, 250, 252)  # Light Soft Grey (#F8FAFC)
COLOR_BORDER = RGBColor(226, 232, 240)   # Light Border (#E2E8F0)

def add_visual_slide(prs, step_no, title, subtitle, items, img_path):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 상단 헤더
    header_tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf = header_tb.text_frame
    htf.word_wrap = True

    p_step = htf.paragraphs[0]
    p_step.text = f"STEP {step_no}  |  화면 구성 및 번호별 기능 가이드"
    p_step.font.size = Pt(12)
    p_step.font.bold = True
    p_step.font.color.rgb = COLOR_ACCENT

    p_title = htf.add_paragraph()
    p_title.text = title
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_PRIMARY

    p_sub = htf.add_paragraph()
    p_sub.text = subtitle
    p_sub.font.size = Pt(13)
    p_sub.font.color.rgb = COLOR_TEXT_MUTED

    # 좌측: 지시선 번호 대응 설명 카드 (너비 4.9인치)
    card_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.65), Inches(4.9), Inches(5.4))
    card_box.fill.solid()
    card_box.fill.fore_color.rgb = COLOR_BG_CARD
    card_box.line.color.rgb = COLOR_BORDER

    card_tb = slide.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(4.5), Inches(5.1))
    ctf = card_tb.text_frame
    ctf.word_wrap = True

    for idx, (item_title, item_desc) in enumerate(items):
        p_it = ctf.paragraphs[0] if idx == 0 else ctf.add_paragraph()
        p_it.text = item_title
        p_it.font.size = Pt(13)
        p_it.font.bold = True
        p_it.font.color.rgb = COLOR_ACCENT
        p_it.space_after = Pt(2)

        p_id = ctf.add_paragraph()
        p_id.text = item_desc
        p_id.font.size = Pt(11)
        p_id.font.color.rgb = COLOR_TEXT_MAIN
        p_id.space_after = Pt(10)

    # 우측: 실제 화면 목업 삽화 이미지 삽입 (너비 6.6인치)
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path, Inches(5.9), Inches(1.65), width=Inches(6.6))
    return slide


# ===========================================================================
# 1. 사용자 전용 PPTX 매뉴얼 생성 (관리자 관련 내용 완전 배제)
# ===========================================================================
def build_user_presentation(images):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- SLIDE 1: 표지 ---
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = COLOR_PRIMARY
    bg1.line.color.rgb = COLOR_PRIMARY

    tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.3), Inches(4.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p1_tag = tf1.paragraphs[0]
    p1_tag.text = "SMART OVERTIME SYSTEM v1.36  |  일반 사원 전용 간편 매뉴얼"
    p1_tag.font.size = Pt(14)
    p1_tag.font.bold = True
    p1_tag.font.color.rgb = RGBColor(253, 186, 116)
    p1_tag.space_after = Pt(14)

    p1_title = tf1.add_paragraph()
    p1_title.text = "스마트 특근 관리 시스템\n사용자 전용 기능 매뉴얼 (v1.36)"
    p1_title.font.size = Pt(36)
    p1_title.font.bold = True
    p1_title.font.color.rgb = RGBColor(255, 255, 255)
    p1_title.space_after = Pt(20)

    p1_sub = tf1.add_paragraph()
    p1_sub.text = "사원번호 간편 입장, 스마트 달력 신청, 다차원 검색 및 개인 달력 조회까지 누구나 쉽게 따라 할 수 있는 사용자 공식 가이드입니다."
    p1_sub.font.size = Pt(15)
    p1_sub.font.color.rgb = RGBColor(203, 213, 225)
    p1_sub.space_after = Pt(28)

    p1_auth = tf1.add_paragraph()
    p1_auth.text = "배포 버전: v1.36 (2026-09-17)  |  PC 모니터 & 스마트폰(모바일) 완벽 지원"
    p1_auth.font.size = Pt(13)
    p1_auth.font.color.rgb = RGBColor(148, 163, 184)

    # --- SLIDE 2: 1단계 - 사번만 넣고 슝 들어가기 ---
    s2_data = [
        ("① v1.36 버전 뱃지 확인", "화면 오른쪽 위에 파란색 [v1.36] 뱃지가 보이면 최신 버전입니다. 클릭 시 신규 릴리즈 이력이 표시됩니다."),
        ("② 사원번호 간편 입력", "본인의 사원번호를 입력합니다. 복잡한 비밀번호 없이 빠르게 입장 가능합니다."),
        ("③ [입장하기] 버튼 클릭", "등록된 사원은 이름과 소속팀이 자동 조회되어 즉시 메인 대시보드로 이동합니다."),
        ("④ 미등록 사번 즉시 등록", "처음 방문한 사번은 신규 등록창이 나타나며, 성명과 소속팀을 선택하면 즉시 등록되어 입장합니다.")
    ]
    add_visual_slide(prs, "01", "사원번호로 간편 입장하기", "비밀번호 없이 사원번호 입력만으로 빠르고 안전하게 입장합니다.", s2_data, images["login"])

    # --- SLIDE 3: 2단계 - 달력 콕 찍어서 특근 신청하기 ---
    s3_data = [
        ("① [일반휴일] 기본 선택 & 토요일 자동 세팅", "주말 특근에 가장 많이 쓰이는 '일반휴일'이 디폴트로 선택되며, 시작일/종료일이 이번 주 토요일로 자동 설정됩니다."),
        ("② 시작일/종료일 상호 자동 동기화 (v1.35 신규)", "시작일(또는 종료일)을 선택하면 다른 날짜도 동일하게 자동 변경되며, 이후 두 번째 날짜를 변경해 2단계로 자유롭게 범위를 설정합니다."),
        ("③ 프로젝트번호/장소/사유 입력 가이드", "프로젝트번호(예: BT2601-L1), 근무장소(예: 본사5층), 특근사유(예: 프로그램 개발) 예시 가이드가 기본 제공됩니다."),
        ("④ 달력 인터랙티브 날짜 선택 & 중복 방지", "오른쪽 달력에서 날짜를 터치하여 선택할 수 있으며, 동일한 날짜에 중복 신청하는 실수를 사전에 원천 차단합니다.")
    ]
    add_visual_slide(prs, "02", "스마트 달력 기반 특근 신청", "일반휴일 및 이번 주 토요일이 기본 선택되어 몇 초 만에 신청이 완료됩니다.", s3_data, images["apply"])

    # --- SLIDE 4: 3단계 - 나의 특근 신청 내역 및 다차원 검색 ---
    s4_data = [
        ("① 실시간 검색 결과 통계 요약 바 (v1.35 신규)", "검색 조건에 맞춰 총 건수(총 일수), 승인완료/대기 건수, 일반/법정/대체근무별 수량이 미니 통계 칩으로 실시간 집계 표시됩니다."),
        ("② 다차원 검색 필터 (요구사항 5)", "조회기간, 특근구분, 승인상태 및 프로젝트/장소/사유 통합 키워드 검색으로 원하는 특근을 즉시 찾습니다."),
        ("③ 목록 보기 ↔ 개인 달력 보기 전환", "신청 내역을 카드 리스트 형태뿐만 아니라 개인 전용 월간 달력 형태로도 한눈에 확인할 수 있습니다."),
        ("④ 승인 상태 및 대체휴일 명확 표시 & 수정/취소", "관리자 승인 시 [✓ 확인완료] 표시, 대휴 일수 안내, 일정 변경 시 [수정], 취소 시 [삭제] 및 이력이 안전 보존됩니다.")
    ]
    add_visual_slide(prs, "03", "나의 특근 신청 내역 및 다차원 검색", "다양한 검색 조건과 개인 달력 뷰로 본인의 특근 일정을 편리하게 확인합니다.", s4_data, images["my"])

    # --- SLIDE 5: 4단계 - 스마트폰 모바일 & 외부망 원격 접속 ---
    s5 = prs.slides.add_slide(blank_layout)
    header_tb = s5.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf = header_tb.text_frame
    htf.word_wrap = True
    p_step = htf.paragraphs[0]
    p_step.text = "STEP 04  |  스마트폰 모바일 및 외부망 접속 가이드"
    p_step.font.size = Pt(12)
    p_step.font.bold = True
    p_step.font.color.rgb = COLOR_ACCENT

    p_title = htf.add_paragraph()
    p_title.text = "모바일 및 외부망(LTE/5G/사외) 간편 접속"
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_PRIMARY

    p_sub = htf.add_paragraph()
    p_sub.text = "PC 앞이 아니어도 스마트폰으로 특근을 신청하고 실시간 승인 현황을 확인할 수 있습니다."
    p_sub.font.size = Pt(13)
    p_sub.font.color.rgb = COLOR_TEXT_MUTED

    step_cards = [
        ("1. 모바일 접속 QR 확인", "상단 [📱 모바일 접속 QR]을 누르면 스마트폰으로 스캔할 수 있는 고해상도 QR 코드가 표시됩니다.", "#eff6ff", "#2563eb"),
        ("2. 외부망 URL QR 생성", "사외나 자택(LTE/5G)에서 접속할 경우 외부 공개 주소를 입력하여 전용 QR 코드를 즉시 생성할 수 있습니다.", "#ecfdf5", "#10b981"),
        ("3. 세션 유지 & 홈 화면 추가", "스마트폰 브라우저에서 '홈 화면에 추가'를 누르면 앱처럼 등록되며, 세션이 영구 보존되어 편리하게 이용됩니다.", "#fef3c7", "#ea580c")
    ]

    for i, (c_title, c_desc, c_bg, c_line) in enumerate(step_cards):
        cx = Inches(0.8 + i * 4.0)
        c_box = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), Inches(3.7), Inches(4.8))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        c_box.line.color.rgb = RGBColor(203, 213, 225)
        c_box.line.width = Pt(2)

        c_tb = s5.shapes.add_textbox(cx + Inches(0.2), Inches(2.0), Inches(3.3), Inches(4.3))
        ctf = c_tb.text_frame
        ctf.word_wrap = True

        p1 = ctf.paragraphs[0]
        p1.text = f"단계 0{i+1}"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(234, 88, 12)
        p1.space_after = Pt(8)

        p2 = ctf.add_paragraph()
        p2.text = c_title
        p2.font.size = Pt(18)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_PRIMARY
        p2.space_after = Pt(14)

        p3 = ctf.add_paragraph()
        p3.text = c_desc
        p3.font.size = Pt(13)
        p3.font.color.rgb = COLOR_TEXT_MAIN
        p3.line_spacing = 1.3

    # --- SLIDE 6: 사용자용 Q&A 총정리 ---
    s6 = prs.slides.add_slide(blank_layout)
    header_tb6 = s6.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf6 = header_tb6.text_frame
    htf6.word_wrap = True
    p_step6 = htf6.paragraphs[0]
    p_step6.text = "SUMMARY & Q&A  |  사용자 자주 묻는 질문 답변"
    p_step6.font.size = Pt(12)
    p_step6.font.bold = True
    p_step6.font.color.rgb = COLOR_ACCENT

    p_title6 = htf6.add_paragraph()
    p_title6.text = "자주 묻는 질문(Q&A)과 사용자 꿀팁 총정리"
    p_title6.font.size = Pt(22)
    p_title6.font.bold = True
    p_title6.font.color.rgb = COLOR_PRIMARY

    p_sub6 = htf6.add_paragraph()
    p_sub6.text = "궁금한 사항이 있으실 때는 언제든 상단의 [📖 사용자 매뉴얼]을 다운로드하여 확인하세요."
    p_sub6.font.size = Pt(13)
    p_sub6.font.color.rgb = COLOR_TEXT_MUTED

    user_qa_items = [
        ("Q1. 사원번호는 어떻게 입력하나요?", "본인의 사원번호를 입력하면 비밀번호 없이 간편하게 입장하여 특근을 신청할 수 있습니다."),
        ("Q2. 프로젝트번호와 사유 예시 가이드가 있나요?", "네, 프로젝트번호는 'BT2601-L1', 장소는 '본사5층', 사유는 '프로그램 개발'로 기본 예시 안내가 제공됩니다."),
        ("Q3. 실수로 같은 날짜에 중복 신청하면 어떻게 되나요?", "시스템에서 동일 날짜 중복 신청을 사전에 감지하여 경고창과 함께 안전하게 차단합니다."),
        ("Q4. 내가 신청한 특근을 달력 형태로도 볼 수 있나요?", "네, [📅 달력 보기] 탭을 누르면 개인 월간 달력에서 승인완료/대기 상태가 색상별로 표시됩니다."),
        ("Q5. 스마트폰 모바일에서 신청해도 PC와 연동되나요?", "네, 모바일과 PC는 동일한 중앙 데이터베이스를 사용하므로 실시간 100% 자동 연동됩니다.")
    ]

    for i, (q, a) in enumerate(user_qa_items):
        qy = Inches(1.7 + i * 1.05)
        q_box = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), qy, Inches(11.733), Inches(0.95))
        q_box.fill.solid()
        q_box.fill.fore_color.rgb = RGBColor(248, 250, 252)
        q_box.line.color.rgb = RGBColor(226, 232, 240)

        q_tb = s6.shapes.add_textbox(Inches(1.0), qy + Inches(0.08), Inches(11.3), Inches(0.8))
        qtf = q_tb.text_frame
        qtf.word_wrap = True

        qp1 = qtf.paragraphs[0]
        qp1.text = q
        qp1.font.size = Pt(13)
        qp1.font.bold = True
        qp1.font.color.rgb = COLOR_PRIMARY
        qp1.space_after = Pt(2)

        qp2 = qtf.add_paragraph()
        qp2.text = f"👉 {a}"
        qp2.font.size = Pt(11)
        qp2.font.color.rgb = COLOR_ACCENT

    return prs


# ===========================================================================
# 2. 관리자 전용 PPTX 매뉴얼 생성 (팀원 대리신청, 보너스, 권한격리, 엑셀원장 등)
# ===========================================================================
def build_admin_presentation(images):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- SLIDE 1: 표지 ---
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = COLOR_PRIMARY
    bg1.line.color.rgb = COLOR_PRIMARY

    tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.3), Inches(4.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p1_tag = tf1.paragraphs[0]
    p1_tag.text = "SMART OVERTIME SYSTEM v1.36  |  관리자 및 운영자 전용 가이드"
    p1_tag.font.size = Pt(14)
    p1_tag.font.bold = True
    p1_tag.font.color.rgb = RGBColor(253, 186, 116)
    p1_tag.space_after = Pt(14)

    p1_title = tf1.add_paragraph()
    p1_title.text = "스마트 특근 관리 시스템\n관리자 모드 운영 매뉴얼 (v1.36)"
    p1_title.font.size = Pt(36)
    p1_title.font.bold = True
    p1_title.font.color.rgb = RGBColor(255, 255, 255)
    p1_title.space_after = Pt(20)

    p1_sub = tf1.add_paragraph()
    p1_sub.text = "보안 감사 로그, 캘린더 색상 범례, 팀원 대리 신청(비밀 보너스), 권한 격리, 실특근 정산 및 엑셀 원장 관리자 공식 가이드입니다."
    p1_sub.font.size = Pt(15)
    p1_sub.font.color.rgb = RGBColor(203, 213, 225)
    p1_sub.space_after = Pt(28)

    p1_auth = tf1.add_paragraph()
    p1_auth.text = "배포 버전: v1.36 (2026-09-17)  |  총괄 슈퍼관리자 및 부서 팀관리자 전용"
    p1_auth.font.size = Pt(13)
    p1_auth.font.color.rgb = RGBColor(148, 163, 184)

    # --- SLIDE 2: 관리자 2등급제 및 권한 격리 ---
    s2 = prs.slides.add_slide(blank_layout)
    header_tb = s2.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf = header_tb.text_frame
    htf.word_wrap = True
    p_step = htf.paragraphs[0]
    p_step.text = "STEP 01  |  관리자 등급 체계 및 보안 격리 원칙"
    p_step.font.size = Pt(12)
    p_step.font.bold = True
    p_step.font.color.rgb = COLOR_ACCENT

    p_title = htf.add_paragraph()
    p_title.text = "총괄관리자 vs 팀관리자 2등급제 & 일정 완전 격리"
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_PRIMARY

    p_sub = htf.add_paragraph()
    p_sub.text = "철저한 권한 분리로 팀관리자는 본인 부서만 통제하며, 총괄관리자 일정은 완벽히 격리 차단됩니다."
    p_sub.font.size = Pt(13)
    p_sub.font.color.rgb = COLOR_TEXT_MUTED

    admin_role_cards = [
        ("1. 총괄 슈퍼관리자 권한", "전 부서의 특근 승인, 소속팀 등록/삭제, 팀관리자 권한 토글 및 전사 정산 통계를 총괄 제어합니다.", "#fef3c7", "#d97706"),
        ("2. 팀관리자 전담 관리 & 격리", "본인 소속팀의 특근만 승인/관리하며, 총괄관리자의 특근 일정은 목록·달력·정산에서 일체 노출되지 않습니다.", "#e0f2fe", "#0284c7"),
        ("3. 동시 접속 실시간 동기화", "여러 관리자가 동시에 접속해 승인/신청할 경우 상단 [🔄 새로고침] 버튼을 눌러 즉시 최신 상태로 동기화합니다.", "#ecfdf5", "#10b981")
    ]

    for i, (c_title, c_desc, c_bg, c_line) in enumerate(admin_role_cards):
        cx = Inches(0.8 + i * 4.0)
        c_box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), Inches(3.7), Inches(4.8))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        c_box.line.color.rgb = RGBColor(203, 213, 225)
        c_box.line.width = Pt(2)

        c_tb = s2.shapes.add_textbox(cx + Inches(0.2), Inches(2.0), Inches(3.3), Inches(4.3))
        ctf = c_tb.text_frame
        ctf.word_wrap = True

        p1 = ctf.paragraphs[0]
        p1.text = f"핵심 원칙 0{i+1}"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(234, 88, 12)
        p1.space_after = Pt(8)

        p2 = ctf.add_paragraph()
        p2.text = c_title
        p2.font.size = Pt(18)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_PRIMARY
        p2.space_after = Pt(14)

        p3 = ctf.add_paragraph()
        p3.text = c_desc
        p3.font.size = Pt(13)
        p3.font.color.rgb = COLOR_TEXT_MAIN
        p3.line_spacing = 1.3

    # --- SLIDE 3: 관리자 월간 캘린더 & 색상 범례 & 일괄 승인 ---
    s3_data = [
        ("① 보너스(🎁) 식별 & 전원 노출 (v1.35 신규)", "보너스 부여 인원은 이름 앞 🎁 아이콘 및 보라색 강조 테두리, 일자 헤더에 🎁N명 뱃지가 표시되며, +N명 축약 없이 전원 노출됩니다."),
        ("② 인원수에 맞춘 달력 칸 자동 확장 (v1.35 신규)", "신청 인원이 많아지면 달력 셀 높이가 자동으로 늘어나 글자 찌그러짐과 잘림 현상이 완전히 해소되었습니다."),
        ("③ 캘린더 색상 범례 바 (요구사항 1)", "캘린더 하단에 일반휴일, 대체근무, 법정휴일, 대체휴일 사용, 🎁 보너스를 구분하는 색상 범례 바가 상시 표시됩니다."),
        ("④ 날짜별 상세 패널 & 당일 전원 일괄 확인", "달력 일자 클릭 시 하단에 작업자 상세 목록이 펼쳐지며, [당일 전원 일괄 확인]으로 원클릭 일괄 승인합니다.")
    ]
    add_visual_slide(prs, "02", "[관리자 모드] 월간 캘린더 및 일괄 승인", "색상 범례 확인, 일자별 작업자 패널 및 원클릭 일괄 승인을 지원합니다.", s3_data, images["cal"])

    # --- SLIDE 4: 팀원 특근 대리 신청 & 비밀 보너스 ---
    s4 = prs.slides.add_slide(blank_layout)
    header_tb4 = s4.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf4 = header_tb4.text_frame
    htf4.word_wrap = True
    p_step4 = htf4.paragraphs[0]
    p_step4.text = "STEP 03  |  팀원 특근 대리 신청 및 비밀 보너스 부여"
    p_step4.font.size = Pt(12)
    p_step4.font.bold = True
    p_step4.font.color.rgb = COLOR_ACCENT

    p_title4 = htf4.add_paragraph()
    p_title4.text = "➕ 팀원 대리 신청 & 🎁 비밀 보너스 (보안 원칙)"
    p_title4.font.size = Pt(22)
    p_title4.font.bold = True
    p_title4.font.color.rgb = COLOR_PRIMARY

    p_sub4 = htf4.add_paragraph()
    p_sub4.text = "관리자가 팀원을 지정해 특근을 대리 등록하며, 부여된 보너스는 사원은 전혀 모르게 안전 은닉됩니다."
    p_sub4.font.size = Pt(13)
    p_sub4.font.color.rgb = COLOR_TEXT_MUTED

    proxy_cards = [
        ("1. 팀원 선택 및 기본 가이드", "총괄관리자는 전 사원, 팀관리자는 본인 팀원을 선택할 수 있으며, 토요일 날짜 및 프로젝트/장소/사유 예시 가이드가 제공됩니다.", "#eff6ff", "#2563eb"),
        ("2. 🎁 비밀 보너스 부여 원칙", "체크 시 보너스가 부여되지만, 사원 화면 및 사용자 API에서는 무조건 0으로 마스킹되어 사원은 부여 사실을 전혀 모릅니다.", "#ecfdf5", "#10b981"),
        ("3. 관리자 화면 & 엑셀 명시", "관리자 테이블 및 다운로드 엑셀 파일(Sheet 1, 3)에는 '보너스 부여: 부여(O)/미부여(-)' 항목이 정확히 기재됩니다.", "#fef3c7", "#ea580c")
    ]

    for i, (c_title, c_desc, c_bg, c_line) in enumerate(proxy_cards):
        cx = Inches(0.8 + i * 4.0)
        c_box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), Inches(3.7), Inches(4.8))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        c_box.line.color.rgb = RGBColor(203, 213, 225)
        c_box.line.width = Pt(2)

        c_tb = s4.shapes.add_textbox(cx + Inches(0.2), Inches(2.0), Inches(3.3), Inches(4.3))
        ctf = c_tb.text_frame
        ctf.word_wrap = True

        p1 = ctf.paragraphs[0]
        p1.text = f"대리신청 규정 0{i+1}"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(234, 88, 12)
        p1.space_after = Pt(8)

        p2 = ctf.add_paragraph()
        p2.text = c_title
        p2.font.size = Pt(18)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_PRIMARY
        p2.space_after = Pt(14)

        p3 = ctf.add_paragraph()
        p3.text = c_desc
        p3.font.size = Pt(13)
        p3.font.color.rgb = COLOR_TEXT_MAIN
        p3.line_spacing = 1.3

    # --- SLIDE 5: 팀원 명부 및 소속팀 관리 ---
    s5_data = [
        ("① 팀원 정보 원클릭 수정 (요구사항 1)", "팀원의 성명, 직급, 소속팀을 모달창에서 손쉽게 수정하고 즉시 시스템에 동기화합니다."),
        ("② 팀관리자는 본인 소속팀만 전담 제어 (요구사항 2)", "팀관리자는 본인 팀원만 수정/삭제/승인할 수 있고, 타 팀원의 정보나 총괄관리자 정보는 변경할 수 없습니다."),
        ("③ 소속팀 등록 및 미사용 팀 정리", "총괄관리자 전용 소속팀 관리 팝업을 통해 신규 팀을 등록하거나 미사용 팀을 안전하게 정리합니다."),
        ("④ 엑셀 일괄 등록 및 내보내기", "팀원 명부를 엑셀 양식으로 대량 등록하거나 최신 명부를 엑셀 파일로 다운로드합니다.")
    ]
    add_visual_slide(prs, "04", "[관리자 모드] 팀원 명부 및 소속팀 관리", "팀원 정보 수정, 팀관리자 권한 토글 및 소속팀 관리를 지원합니다.", s5_data, images["user"])

    # --- SLIDE 6: 실특근 자동 산정 및 정산표 ---
    s6_data = [
        ("① 명확한 실특근 산정 공식", "★ 실특근 인정일 = 일반휴일 근무일 - 대체휴일 사용일 (대체근무/법정휴일은 산정에서 자동 분리)"),
        ("② 부서별 실특근 요약 카드", "소속 부서별 총 근무일, 제외 일수, 대체휴일 사용일 및 최종 실특근일을 대형 통계 카드로 제공합니다."),
        ("③ 개인별 정산 명세 테이블", "사원별 4대 휴일 지표와 최종 실특근 일수를 정렬 가능한 테이블로 한눈에 검토합니다."),
        ("④ 정산표 엑셀 원클릭 추출", "[📥 정산표 엑셀 내보내기] 버튼으로 월말 보고용 정산 파일을 즉시 다운로드합니다.")
    ]
    add_visual_slide(prs, "05", "[관리자 모드] 실특근 자동 산정 및 정산표", "대체휴일 사용분을 자동 차감하여 정확한 최종 실특근일을 산출합니다.", s6_data, images["settlement"])

    # --- SLIDE 7: 엑셀 3대 양식 & [보너스 부여] 열 기재 ---
    s7_data = [
        ("① 엑셀 파일 내 [보너스 부여] 열 기재 (요구사항 4)", "휴일일자별 특근현황 및 특근신청 원장 시트에 [보너스 부여] 열이 추가되어 부여(O)/미부여(-) 상태가 명시됩니다."),
        ("② 날짜순 오름차순 전개 & 순수 숫자 셀", "모든 일자가 하루단위로 펼쳐지며, 일수 셀이 순수 숫자(int/float)로 저장되어 수식 오류가 전혀 없습니다."),
        ("③ 개인별 휴일합산 정산표 제공", "두 번째 시트에 인원별 4대 휴일수 및 최종 실특근일이 완벽히 합산 정산됩니다."),
        ("④ openpyxl 고속 서버 엑셀 엔진", "수천 건의 특근 데이터도 엑셀 수식과 서식을 온전히 유지하며 실시간 스트리밍으로 다운로드합니다.")
    ]
    add_visual_slide(prs, "06", "엑셀 [보너스 부여] 기재 및 3대 양식", "고품질 엑셀 추출과 보너스 부여 명시를 완벽히 지원합니다.", s7_data, images["excel"])

    # --- SLIDE 8: 관리자 전용 웹 스냅샷 저장 및 복원 ---
    s8 = prs.slides.add_slide(blank_layout)
    header_tb8 = s8.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf8 = header_tb8.text_frame
    htf8.word_wrap = True
    p_step8 = htf8.paragraphs[0]
    p_step8.text = "STEP 07  |  관리자 전용 웹 데이터 스냅샷 백업 & 복원"
    p_step8.font.size = Pt(12)
    p_step8.font.bold = True
    p_step8.font.color.rgb = COLOR_ACCENT

    p_title8 = htf8.add_paragraph()
    p_title8.text = "💾 웹 저장(스냅샷 백업) 및 📂 웹 열기(원클릭 복원)"
    p_title8.font.size = Pt(22)
    p_title8.font.bold = True
    p_title8.font.color.rgb = COLOR_PRIMARY

    p_sub8 = htf8.add_paragraph()
    p_sub8.text = "사원 화면에서는 숨겨져 관리자만 전담 제어하며, 전체 DB를 안전한 JSON 스냅샷으로 영구 보존합니다."
    p_sub8.font.size = Pt(13)
    p_sub8.font.color.rgb = COLOR_TEXT_MUTED

    backup_cards = [
        ("1. 💾 신규 웹 스냅샷 저장", "전체 팀원 명부, 소속팀 목록, 특근 신청 및 이력 데이터를 안전한 JSON 파일로 서버에 즉시 보관합니다.", "#eff6ff", "#2563eb"),
        ("2. 📂 저장된 스냅샷 원클릭 복원", "저장된 백업 목록에서 원하는 시점의 [열기] 버튼을 누르면 해당 시점의 데이터로 1초 만에 안전 복원됩니다.", "#ecfdf5", "#10b981"),
        ("3. 자동 안전 보조백업 생성", "복원 실행 시 현재 데이터가 사라지지 않도록 직전 상태의 보조 백업을 자동 생성하여 무소실을 보장합니다.", "#fef3c7", "#ea580c")
    ]

    for i, (c_title, c_desc, c_bg, c_line) in enumerate(backup_cards):
        cx = Inches(0.8 + i * 4.0)
        c_box = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), Inches(3.7), Inches(4.8))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        c_box.line.color.rgb = RGBColor(203, 213, 225)
        c_box.line.width = Pt(2)

        c_tb = s8.shapes.add_textbox(cx + Inches(0.2), Inches(2.0), Inches(3.3), Inches(4.3))
        ctf = c_tb.text_frame
        ctf.word_wrap = True

        p1 = ctf.paragraphs[0]
        p1.text = f"백업 관리 0{i+1}"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(234, 88, 12)
        p1.space_after = Pt(8)

        p2 = ctf.add_paragraph()
        p2.text = c_title
        p2.font.size = Pt(18)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_PRIMARY
        p2.space_after = Pt(14)

        p3 = ctf.add_paragraph()
        p3.text = c_desc
        p3.font.size = Pt(13)
        p3.font.color.rgb = COLOR_TEXT_MAIN
        p3.line_spacing = 1.3

    # --- SLIDE 9: 보안 감사 및 사용자 접속 로그 관리 (신규 v1.36 요구사항 5) ---
    s9_sec = prs.slides.add_slide(blank_layout)
    header_tb9_sec = s9_sec.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf9_sec = header_tb9_sec.text_frame
    htf9_sec.word_wrap = True
    p_step9_sec = htf9_sec.paragraphs[0]
    p_step9_sec.text = "STEP 08  |  슈퍼관리자 전용 시스템 보안 및 접속 감사 로그 (v1.36 신규)"
    p_step9_sec.font.size = Pt(12)
    p_step9_sec.font.bold = True
    p_step9_sec.font.color.rgb = COLOR_ACCENT

    p_title9_sec = htf9_sec.add_paragraph()
    p_title9_sec.text = "🛡️ 시스템 보안 및 사용자 접속 감사 로그 관리"
    p_title9_sec.font.size = Pt(22)
    p_title9_sec.font.bold = True
    p_title9_sec.font.color.rgb = COLOR_PRIMARY

    p_sub9_sec = htf9_sec.add_paragraph()
    p_sub9_sec.text = "사원 로그인 시도(성공/실패), 신규 사원 등록, IP 주소 및 기기(User-Agent) 이력을 실시간 감사·추적합니다."
    p_sub9_sec.font.size = Pt(13)
    p_sub9_sec.font.color.rgb = COLOR_TEXT_MUTED

    security_cards = [
        ("1. 실시간 접속/보안 감사 추적", "로그인 시도 성공/실패, 미등록 사번 접근 시도, 신규 사원 등록, 로그아웃 이력이 초단위로 영구 기록됩니다.", "#eff6ff", "#2563eb"),
        ("2. 다차원 필터링 & 판독 용이성", "기간(시작~종료), 액션 구분(LOGIN/REGISTER/LOGOUT), 상태(성공/실패), 사번/성명/IP 검색으로 비정상 접속을 1초 만에 식별합니다.", "#ecfdf5", "#10b981"),
        ("3. 서식화된 엑셀(.xlsx) 내보내기", "openpyxl 기반 컬러 배지(성공: 연녹색, 실패: 연빨강)가 적용된 감사 보고서 엑셀 파일을 원클릭으로 즉시 다운로드합니다.", "#fef3c7", "#ea580c")
    ]

    for i, (c_title, c_desc, c_bg, c_line) in enumerate(security_cards):
        cx = Inches(0.8 + i * 4.0)
        c_box = s9_sec.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), Inches(3.7), Inches(4.8))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        c_box.line.color.rgb = RGBColor(203, 213, 225)
        c_box.line.width = Pt(2)

        c_tb = s9_sec.shapes.add_textbox(cx + Inches(0.2), Inches(2.0), Inches(3.3), Inches(4.3))
        ctf = c_tb.text_frame
        ctf.word_wrap = True

        p1 = ctf.paragraphs[0]
        p1.text = f"보안 체계 0{i+1}"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(234, 88, 12)
        p1.space_after = Pt(8)

        p2 = ctf.add_paragraph()
        p2.text = c_title
        p2.font.size = Pt(18)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_PRIMARY
        p2.space_after = Pt(14)

        p3 = ctf.add_paragraph()
        p3.text = c_desc
        p3.font.size = Pt(13)
        p3.font.color.rgb = COLOR_TEXT_MAIN
        p3.line_spacing = 1.3

    # --- SLIDE 10: 모바일 및 외부망 접속 관리 ---
    s9 = prs.slides.add_slide(blank_layout)
    header_tb9 = s9.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf9 = header_tb9.text_frame
    htf9.word_wrap = True
    p_step9 = htf9.paragraphs[0]
    p_step9.text = "STEP 09  |  스마트폰 모바일 및 외부망 접속 관리"
    p_step9.font.size = Pt(12)
    p_step9.font.bold = True
    p_step9.font.color.rgb = COLOR_ACCENT

    p_title9 = htf9.add_paragraph()
    p_title9.text = "모바일 접속 QR 배포 & 외부망 원격 터널링"
    p_title9.font.size = Pt(22)
    p_title9.font.bold = True
    p_title9.font.color.rgb = COLOR_PRIMARY

    p_sub9 = htf9.add_paragraph()
    p_sub9.text = "사내 인원들이 외부망(LTE/5G/자택)에서도 원활히 접속할 수 있도록 QR과 URL을 관리합니다."
    p_sub9.font.size = Pt(13)
    p_sub9.font.color.rgb = COLOR_TEXT_MUTED

    mobile_cards = [
        ("1. 사내 모바일 접속 QR", "상단 [📱 모바일 접속 QR]을 띄워 회의실이나 현장에서 직원들이 스마트폰으로 바로 접속할 수 있게 안내합니다.", "#eff6ff", "#2563eb"),
        ("2. 외부망 전용 URL QR 생성", "Cloudflare 등 외부 접속 주소가 변경될 경우 팝업에서 즉시 새 주소를 입력하여 전용 QR 코드를 생성합니다.", "#ecfdf5", "#10b981"),
        ("3. 상시 백그라운드 구동", "서버 프로세스가 백그라운드에서 상시 유지되어 언제 어디서나 접속 및 데이터 동기화가 유지됩니다.", "#fef3c7", "#ea580c")
    ]

    for i, (c_title, c_desc, c_bg, c_line) in enumerate(mobile_cards):
        cx = Inches(0.8 + i * 4.0)
        c_box = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), Inches(3.7), Inches(4.8))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        c_box.line.color.rgb = RGBColor(203, 213, 225)
        c_box.line.width = Pt(2)

        c_tb = s9.shapes.add_textbox(cx + Inches(0.2), Inches(2.0), Inches(3.3), Inches(4.3))
        ctf = c_tb.text_frame
        ctf.word_wrap = True

        p1 = ctf.paragraphs[0]
        p1.text = f"운영 안내 0{i+1}"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(234, 88, 12)
        p1.space_after = Pt(8)

        p2 = ctf.add_paragraph()
        p2.text = c_title
        p2.font.size = Pt(18)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_PRIMARY
        p2.space_after = Pt(14)

        p3 = ctf.add_paragraph()
        p3.text = c_desc
        p3.font.size = Pt(13)
        p3.font.color.rgb = COLOR_TEXT_MAIN
        p3.line_spacing = 1.3

    # --- SLIDE 11: 관리자 전용 Q&A 총정리 ---
    s10 = prs.slides.add_slide(blank_layout)
    header_tb10 = s10.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf10 = header_tb10.text_frame
    htf10.word_wrap = True
    p_step10 = htf10.paragraphs[0]
    p_step10.text = "SUMMARY & Q&A  |  관리자 전용 질문 답변"
    p_step10.font.size = Pt(12)
    p_step10.font.bold = True
    p_step10.font.color.rgb = COLOR_ACCENT

    p_title10 = htf10.add_paragraph()
    p_title10.text = "자주 묻는 질문(Q&A)과 관리자 운영 꿀팁 총정리 (v1.36)"
    p_title10.font.size = Pt(22)
    p_title10.font.bold = True
    p_title10.font.color.rgb = COLOR_PRIMARY

    p_sub10 = htf10.add_paragraph()
    p_sub10.text = "궁금한 사항이 있으실 때는 언제든 상단의 [📖 관리자 매뉴얼]을 다운로드하여 확인하세요."
    p_sub10.font.size = Pt(13)
    p_sub10.font.color.rgb = COLOR_TEXT_MUTED

    admin_qa_items = [
        ("Q1. 팀관리자가 총괄관리자의 특근 일정을 볼 수 있나요?", "아닙니다. 팀관리자는 총괄관리자의 특근 일정을 일체 열람할 수 없도록 철저히 차단 격리되어 있습니다."),
        ("Q2. 팀원 대리 신청 시 비밀 보너스는 사원에게 노출되나요?", "사원 화면 및 API 응답에서는 0으로 마스킹되어 사원은 전혀 모르며, 오직 관리자와 엑셀 원장에만 기재됩니다."),
        ("Q3. 팀관리자와 총괄관리자의 권한 차이는 어떻게 되나요?", "팀관리자는 본인 팀원의 승인/수정만 가능하며, 팀원 소속팀 변경 및 관리자 권한 지정은 오직 총괄관리자만 가능합니다."),
        ("Q4. 시스템 접속 및 보안 감사 로그는 누가 볼 수 있나요?", "오직 총괄 슈퍼관리자만 열람할 수 있으며, 로그인 시도(성공/실패), 신규 등록, 접속 IP 및 User-Agent 정보를 실시간 추적하고 엑셀로 추출합니다."),
        ("Q5. 웹 저장 및 열기 기능은 일반 사원도 볼 수 있나요?", "사원 화면에서는 숨김 처리되어 일반 사원은 볼 수 없으며, 관리자 모드에서만 안전하게 조작할 수 있습니다.")
    ]

    for i, (q, a) in enumerate(admin_qa_items):
        qy = Inches(1.7 + i * 1.05)
        q_box = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), qy, Inches(11.733), Inches(0.95))
        q_box.fill.solid()
        q_box.fill.fore_color.rgb = RGBColor(248, 250, 252)
        q_box.line.color.rgb = RGBColor(226, 232, 240)

        q_tb = s10.shapes.add_textbox(Inches(1.0), qy + Inches(0.08), Inches(11.3), Inches(0.8))
        qtf = q_tb.text_frame
        qtf.word_wrap = True

        qp1 = qtf.paragraphs[0]
        qp1.text = q
        qp1.font.size = Pt(13)
        qp1.font.bold = True
        qp1.font.color.rgb = COLOR_PRIMARY
        qp1.space_after = Pt(2)

        qp2 = qtf.add_paragraph()
        qp2.text = f"👉 {a}"
        qp2.font.size = Pt(11)
        qp2.font.color.rgb = COLOR_ACCENT

    return prs


# ----------------- 실행 및 저장 통합 함수 -----------------

def create_manual():
    print("[1/3] Generating ultra-friendly visual UI mockups with speech bubbles & leader lines...")
    images = {
        "login": create_mockup_login(),
        "apply": create_mockup_calendar_apply(),
        "my": create_mockup_my_records(),
        "cal": create_mockup_admin_calendar(),
        "user": create_mockup_user_mgmt(),
        "settlement": create_mockup_settlement(),
        "excel": create_mockup_excel_27()
    }

    print("[2/3] Assembling User Manual and Admin Manual separately...")
    prs_user = build_user_presentation(images)
    prs_admin = build_admin_presentation(images)

    print("[3/3] Saving PPTX presentations to downloads directory...")
    # 1. 사용자 매뉴얼 저장
    prs_user.save(str(USER_PPTX_PATH))
    print(f"✓ Saved User Manual: {USER_PPTX_PATH}")

    # 2. 관리자 매뉴얼 저장
    prs_admin.save(str(ADMIN_PPTX_PATH))
    print(f"✓ Saved Admin Manual: {ADMIN_PPTX_PATH}")

    # 3. 통합/기존 호환 매뉴얼 저장
    prs_admin.save(str(PPTX_PATH))
    print(f"✓ Saved System Manual: {PPTX_PATH}")

    # 4. 최상위 메인 폴더(루트) 및 static/downloads 디렉토리 동기화 복사
    root_dir = Path(__file__).resolve().parent
    static_dl = root_dir / "static" / "downloads"
    static_dl.mkdir(parents=True, exist_ok=True)
    
    # 루트 메인 폴더로 복사 (GitHub 업로드 편의)
    shutil.copy2(str(USER_PPTX_PATH), str(root_dir / "Overtime_User_Manual.pptx"))
    shutil.copy2(str(ADMIN_PPTX_PATH), str(root_dir / "Overtime_Admin_Manual.pptx"))
    shutil.copy2(str(PPTX_PATH), str(root_dir / "Overtime_System_Manual.pptx"))
    print(f"✓ Copied all PPTX manuals to ROOT main directory: {root_dir}")

    # static/downloads로 복사
    shutil.copy2(str(USER_PPTX_PATH), str(static_dl / "Overtime_User_Manual.pptx"))
    shutil.copy2(str(ADMIN_PPTX_PATH), str(static_dl / "Overtime_Admin_Manual.pptx"))
    shutil.copy2(str(PPTX_PATH), str(static_dl / "Overtime_System_Manual.pptx"))
    print(f"✓ Copied all PPTX manuals to static download path: {static_dl}")
    print("🎉 Both User and Admin PPT Manuals generated and synchronized successfully!")

if __name__ == "__main__":
    create_manual()

