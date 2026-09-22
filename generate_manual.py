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

CANVAS_W = 1600
CANVAS_H = 1100

# ----------------- 지시선 및 사용자용 친절한 말풍선 핀 헬퍼 (v1.54 고해상도 & 고시인성) -----------------

def draw_browser_frame(d, title="★ 스마트 특근 관리 시스템 (v1.54)", url="https://overtime-system.internal"):
    """
    1600x1100 캔버스 상단에 실제 브라우저 형태의 프레임과 주소창, 릴리즈 배지를 렌더링
    """
    d.rounded_rectangle([20, 15, 1580, 85], radius=10, fill="#1e293b")
    # 브라우저 신호등 버튼
    d.ellipse([45, 42, 61, 58], fill="#ef4444")
    d.ellipse([70, 42, 86, 58], fill="#f59e0b")
    d.ellipse([95, 42, 111, 58], fill="#10b981")
    # URL 주소창
    d.rounded_rectangle([135, 30, 980, 70], radius=8, fill="#0f172a", outline="#334155")
    d.text((155, 40), f"[보안연결] {url}", font=get_font(16), fill="#94a3b8")
    d.text((1010, 38), title, font=get_font(18, bold=True), fill="#f8fafc")
    # 최신 버전 배지
    d.rounded_rectangle([1450, 32, 1560, 68], radius=6, fill="#0d9488")
    d.text((1505, 50), "v1.54 최신", font=get_font(16, bold=True), fill="#ffffff", anchor="mm")


def draw_smart_pin(d, target_xy, pin_xy, badge_num, label_text, color="#ea580c", text_color="#ffffff", sub_hint=None):
    """
    21pt 고대비 배지와 16pt 부가설명 힌트, 도넛형 타깃 마커 및
    지시선이 배지 박스 경계에 정확히 도킹하도록 설계된 고시인성 스마트 핀
    """
    tx, ty = target_xy
    px, py = pin_xy

    # 1. 텍스트 너비 및 높이 계산 (배지 21pt)
    font = get_font(21, bold=True)
    badge_str = f" {badge_num}  {label_text} "
    try:
        bbox = font.getbbox(badge_str)
        text_w = (bbox[2] - bbox[0]) + 32
    except Exception:
        text_w = len(badge_str) * 22 + 32
    text_h = 44

    # 핀 박스 중심 및 경계 클리핑
    bx1 = px - text_w // 2
    bx1 = max(15, min(bx1, CANVAS_W - text_w - 15))
    by1 = py - text_h // 2
    by1 = max(15, min(by1, CANVAS_H - text_h - 50))
    bx2 = bx1 + text_w
    by2 = by1 + text_h

    # 2. 지시선 도킹 포인트 산출
    if tx < bx1:
        dock_x = bx1
        dock_y = max(by1 + 6, min(by2 - 6, ty))
    elif tx > bx2:
        dock_x = bx2
        dock_y = max(by1 + 6, min(by2 - 6, ty))
    else:
        dock_x = tx
        dock_y = by1 if ty < by1 else by2

    # 3. 지시선 그리기 (두께 3px 고대비)
    d.line([(tx, ty), (dock_x, dock_y)], fill=color, width=3)
    # 타깃 마커 (외곽 도넛 + 내부 색상점)
    d.ellipse([tx - 9, ty - 9, tx + 9, ty + 9], fill="#ffffff", outline=color, width=3)
    d.ellipse([tx - 4, ty - 4, tx + 4, ty + 4], fill=color)

    # 4. 배지 박스 (그림자 + 테두리)
    d.rounded_rectangle([bx1 + 3, by1 + 4, bx2 + 3, by2 + 4], radius=10, fill="#94a3b8")
    d.rounded_rectangle([bx1, by1, bx2, by2], radius=10, fill=color, outline="#ffffff", width=2)
    d.text((bx1 + text_w // 2, by1 + text_h // 2), badge_str, font=font, fill=text_color, anchor="mm")

    # 5. 서브 힌트(부연 설명 16pt)
    if sub_hint:
        s_font = get_font(16, bold=True)
        sub_str = f" {sub_hint} "
        try:
            s_bbox = s_font.getbbox(sub_str)
            sw = (s_bbox[2] - s_bbox[0]) + 24
        except Exception:
            sw = len(sub_str) * 16 + 24
        sh = 32
        sx1 = max(15, min(bx1 + (text_w - sw) // 2, CANVAS_W - sw - 15))
        sy1 = by2 + 5
        d.rounded_rectangle([sx1 + 2, sy1 + 2, sx1 + sw + 2, sy1 + sh + 2], radius=8, fill="#cbd5e1")
        d.rounded_rectangle([sx1, sy1, sx1 + sw, sy1 + sh], radius=8, fill="#fef9c3", outline="#ca8a04", width=2)
        d.text((sx1 + sw // 2, sy1 + sh // 2), sub_str, font=s_font, fill="#854d0e", anchor="mm")


# ----------------- UI 목업 삽화 이미지 생성기들 (1600x1100 고해상도 & 고시인성) -----------------

# 0. 초기 접속 및 서버 리부팅 대기 안내 목업
def create_mockup_server_init():
    """삽화 0: 서버 초기 구동 및 리부팅 워밍업 대기 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  초기 접속 & 서버 리부팅 안내", url="https://overtime-system.trycloudflare.com")

    # 중앙 대형 카드
    d.rounded_rectangle([350, 115, 1250, 1025], radius=20, fill="#ffffff", outline="#cbd5e1", width=3)

    # 상단 서버 기동 아이콘
    d.ellipse([725, 145, 875, 295], fill="#eff6ff", outline="#3b82f6", width=4)
    d.text((800, 205), "SERVER", font=get_font(26, bold=True), fill="#2563eb", anchor="mm")
    d.text((800, 245), "CONNECTING", font=get_font(17, bold=True), fill="#3b82f6", anchor="mm")

    # 메인 타이틀 & 부제
    d.text((800, 340), "시스템 초기 구동 & 서버 준비 중...", font=get_font(34, bold=True), fill="#0f172a", anchor="mm")
    d.text((800, 385), "최초 접속 또는 서버 재부팅 직후에는 약 10초~20초간 초기화가 진행됩니다.", font=get_font(19), fill="#64748b", anchor="mm")

    # 진행 상태 프로그레스 바
    d.rounded_rectangle([430, 425, 1170, 465], radius=10, fill="#f1f5f9", outline="#cbd5e1")
    d.rounded_rectangle([430, 425, 1010, 465], radius=10, fill="#2563eb")
    d.text((800, 445), "안전한 데이터베이스 연결 & 자동 백업 확인 중... (80%)", font=get_font(16, bold=True), fill="#ffffff", anchor="mm")

    # 3대 자동 초기화 상태 표시 카드
    d.rounded_rectangle([430, 495, 1170, 565], radius=10, fill="#f8fafc", outline="#e2e8f0")
    d.ellipse([455, 520, 475, 540], fill="#10b981")
    d.text((495, 517), "[1] SQLite 데이터 무결성 검증 (WAL 모드 정상 가동)", font=get_font(18, bold=True), fill="#1e293b")
    d.text((495, 542), "DB 파일 무결성 및 테이블 스키마 자동 동기화 완료", font=get_font(14), fill="#64748b")
    d.rounded_rectangle([1060, 513, 1145, 547], radius=6, fill="#dcfce7")
    d.text((1102, 530), "완료", font=get_font(14, bold=True), fill="#16a34a", anchor="mm")

    d.rounded_rectangle([430, 580, 1170, 650], radius=10, fill="#f8fafc", outline="#e2e8f0")
    d.ellipse([455, 605, 475, 625], fill="#2563eb")
    d.text((495, 602), "[2] 스냅샷 자동 백업 생성 (data/backups/)", font=get_font(18, bold=True), fill="#1e293b")
    d.text((495, 627), "서버 시작 시점의 안전한 데이터베이스 복사본 보관 중", font=get_font(14), fill="#64748b")
    d.rounded_rectangle([1060, 598, 1145, 632], radius=6, fill="#dbeafe")
    d.text((1102, 615), "진행중", font=get_font(14, bold=True), fill="#1d4ed8", anchor="mm")

    d.rounded_rectangle([430, 665, 1170, 735], radius=10, fill="#f8fafc", outline="#e2e8f0")
    d.ellipse([455, 690, 475, 710], fill="#f59e0b")
    d.text((495, 687), "[3] 외부 보안망(Cloudflare) 터널 연결", font=get_font(18, bold=True), fill="#1e293b")
    d.text((495, 712), "외부 모바일 QR 및 24시간 전송 암호화 HTTPS 보안 터널 연결", font=get_font(14), fill="#64748b")
    d.rounded_rectangle([1060, 683, 1145, 717], radius=6, fill="#fef3c7")
    d.text((1102, 700), "준비됨", font=get_font(14, bold=True), fill="#b45309", anchor="mm")

    # 핵심 주의사항 배너
    d.rounded_rectangle([430, 755, 1170, 930], radius=12, fill="#fefce8", outline="#eab308", width=2)
    d.text((460, 775), "★ [사용자 필독] 초기 접속 시 꼭 알아두세요!", font=get_font(21, bold=True), fill="#854d0e")
    d.text((460, 815), "1. 화면이 즉시 열리지 않더라도 창을 닫거나 새로고침(F5)을 연타하지 마세요.", font=get_font(17, bold=True), fill="#a16207")
    d.text((460, 850), "   (반복 새로고침 시 서버 초기화 작업이 지연될 수 있습니다.)", font=get_font(15), fill="#a16207")
    d.text((460, 885), "2. 약 10~20초 기다리시면 [사원번호 간편 입장] 화면으로 자동 전환됩니다.", font=get_font(17, bold=True), fill="#15803d")

    # 하단 팁 바
    d.rounded_rectangle([430, 950, 1170, 995], radius=8, fill="#f1f5f9")
    d.text((800, 972), "※ 평상시 서버 가동 중 재입장 시에는 1초 만에 즉시 대시보드가 열립니다.", font=get_font(16), fill="#475569", anchor="mm")

    # 스마트 지시선 (도킹 개선: 텍스트 가림 방지)
    draw_smart_pin(d, (135, 50), (220, 125), "①", "외부 보안 접속 주소 (HTTPS)", color="#0284c7")
    draw_smart_pin(d, (1010, 445), (1350, 445), "②", "약 10~20초 초기 구동 대기", color="#2563eb", sub_hint="리부팅 직후 1회 발생")
    draw_smart_pin(d, (1145, 615), (1350, 615), "③", "DB 무결성 & 자동 백업 진행", color="#10b981", sub_hint="데이터 영구 보존")
    draw_smart_pin(d, (1170, 840), (1350, 840), "④", "새로고침 연타 금지!", color="#ea580c", sub_hint="완료 시 로그인 화면 자동 전환")

    path = IMG_DIR / "mockup_server_init.png"
    img.save(path)
    return str(path)


# 1. 사번 로그인 목업
def create_mockup_login():
    """삽화 1: 사번 입력 화면 목업 (지시선 및 번호 배지)"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  로그인", url="https://overtime.company.com/login")

    # 중앙 카드 (너비 760px)
    d.rounded_rectangle([420, 120, 1180, 1020], radius=20, fill="#ffffff", outline="#cbd5e1", width=3)
    d.ellipse([720, 155, 880, 315], fill="#eff6ff", outline="#3b82f6", width=4)
    d.text((800, 235), "TIME", font=get_font(36, bold=True), fill="#2563eb", anchor="mm")

    d.text((800, 365), "사원번호로 간편 입장", font=get_font(38, bold=True), fill="#0f172a", anchor="mm")
    d.text((800, 415), "사원번호 6자리 숫자를 입력하시면 본인 화면으로 안전하게 입장합니다.", font=get_font(20), fill="#64748b", anchor="mm")

    # 입력창
    d.text((500, 480), "사원번호 (6자리 숫자 입력)", font=get_font(21, bold=True), fill="#1e293b")
    d.rounded_rectangle([500, 515, 1100, 605], radius=12, fill="#ffffff", outline="#3b82f6", width=3)
    d.text((530, 540), "123456", font=get_font(32, bold=True), fill="#0f172a")

    # 입장 버튼
    d.rounded_rectangle([500, 640, 1100, 740], radius=12, fill="#2563eb")
    d.text((800, 690), "입장하기 →", font=get_font(28, bold=True), fill="#ffffff", anchor="mm")

    # 신규 등록 안내 배너
    d.rounded_rectangle([500, 775, 1100, 925], radius=12, fill="#eff6ff", outline="#93c5fd", width=2)
    d.text((800, 820), "[미등록 사번은 즉시 회원등록 창이 열립니다]", font=get_font(22, bold=True), fill="#1e40af", anchor="mm")
    d.text((800, 870), "성명과 소속 부서를 선택하면 즉시 등록되어 대시보드로 이동합니다.", font=get_font(18), fill="#2563eb", anchor="mm")

    d.rounded_rectangle([500, 955, 1100, 1000], radius=8, fill="#f1f5f9")
    d.text((800, 977), "※ 사원번호는 6자리 숫자로 입력해야 정상 처리됩니다 (총괄관리자 제외)", font=get_font(16), fill="#475569", anchor="mm")

    # 스마트 지시선 배치 (충돌 0%)
    draw_smart_pin(d, (1505, 70), (1350, 125), "①", "v1.53 최신 버전 배지", color="#0284c7", sub_hint="최신 릴리즈 확인")
    draw_smart_pin(d, (500, 560), (220, 560), "②", "6자리 숫자 사번 입력창", color="#ea580c", sub_hint="[안내] 6자리 숫자 필수")
    draw_smart_pin(d, (1100, 690), (1380, 690), "③", "[입장하기] 원클릭!", color="#16a34a", sub_hint="즉시 대시보드 입장")
    draw_smart_pin(d, (500, 850), (220, 850), "④", "신규 사원 즉시 등록", color="#9333ea", sub_hint="성명 및 소속 부서 선택")

    path = IMG_DIR / "mockup_login.png"
    img.save(path)
    return str(path)


# 2. 특근 신청 달력 목업
def create_mockup_calendar_apply():
    """삽화 2: 특근 신청 달력 목업 (원클릭 날짜 선택 및 폼)"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  특근 신청", url="https://overtime.company.com/apply")

    d.rounded_rectangle([40, 105, 1560, 1060], radius=16, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((80, 120), "[특근 신청하기] (달력 원클릭 저장)", font=get_font(28, bold=True), fill="#0f172a")
    d.text((80, 155), "달력에서 일할 날짜를 콕 누르고 [특근 신청 저장]만 누르면 신청 완료!", font=get_font(18), fill="#64748b")

    # 좌측: 입력 폼 섹션 (너비 600px: 80 ~ 680)
    d.text((80, 260), "1. 특근 종류 선택 (기본: 일반휴일)", font=get_font(20, bold=True), fill="#1e293b")
    d.rounded_rectangle([80, 295, 265, 360], radius=10, fill="#eff6ff", outline="#2563eb", width=3)
    d.text((172, 327), "● 일반휴일", font=get_font(19, bold=True), fill="#2563eb", anchor="mm")
    d.rounded_rectangle([285, 295, 470, 360], radius=10, fill="#ffffff", outline="#cbd5e1")
    d.text((377, 327), "대체근무", font=get_font(19), fill="#64748b", anchor="mm")
    d.rounded_rectangle([490, 295, 675, 360], radius=10, fill="#ffffff", outline="#cbd5e1")
    d.text((582, 327), "법정휴일", font=get_font(19), fill="#64748b", anchor="mm")

    d.text((80, 385), "2. 시작일 / 종료일 (상호 자동 동기화 & 토요일 자동 세팅)", font=get_font(20, bold=True), fill="#1e293b")
    d.rounded_rectangle([80, 420, 365, 485], radius=8, fill="#f8fafc", outline="#3b82f6", width=2)
    d.text((105, 440), "2026-09-12 (토)", font=get_font(20, bold=True), fill="#0f172a")
    d.rounded_rectangle([390, 420, 675, 485], radius=8, fill="#f8fafc", outline="#cbd5e1")
    d.text((415, 440), "2026-09-13 (일)", font=get_font(20, bold=True), fill="#0f172a")
    d.text((80, 500), "※ 이번 주 토요일이 자동 선택되며, 연속 근무 시 종료일을 변경합니다.", font=get_font(15), fill="#2563eb")

    d.text((80, 545), "3. 프로젝트 번호 & 장소 (가이드 예시)", font=get_font(20, bold=True), fill="#1e293b")
    d.rounded_rectangle([80, 580, 675, 645], radius=8, fill="#ffffff", outline="#cbd5e1")
    d.text((105, 600), "BT2601-L1  |  본사 5층 제어실", font=get_font(20), fill="#334155")

    d.text((80, 675), "4. 일하는 이유 (특근 상세 사유)", font=get_font(20, bold=True), fill="#1e293b")
    d.rounded_rectangle([80, 710, 675, 805], radius=8, fill="#ffffff", outline="#cbd5e1")
    d.text((105, 745), "1공장 제어설비 정기 점검 및 시운전 지원", font=get_font(20), fill="#334155")

    d.rounded_rectangle([80, 835, 675, 930], radius=12, fill="#10b981")
    d.text((377, 882), "[저장] 특근 신청 저장하기 (신청 끝!)", font=get_font(25, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([80, 955, 675, 1035], radius=10, fill="#fefce8", outline="#ca8a04", width=2)
    d.text((105, 975), "★ [특근 신청 핵심 팁]", font=get_font(18, bold=True), fill="#854d0e")
    d.text((105, 1005), "동일 날짜 중복 신청 자동 방지 / 상단 새로고침으로 1초 최신 동기화", font=get_font(15), fill="#a16207")

    # 우측: 달력 인터랙티브 목업 (너비 600px: 800 ~ 1400)
    d.rounded_rectangle([800, 220, 1400, 1035], radius=14, fill="#f8fafc", outline="#cbd5e1", width=2)
    d.text((830, 245), "2026년 9월 특근 달력", font=get_font(24, bold=True), fill="#0f172a")
    d.text((1080, 248), "▶ 날짜를 마우스/손가락으로 콕!", font=get_font(17, bold=True), fill="#2563eb")

    days = ["일", "월", "화", "수", "목", "금", "토"]
    col_w = 82
    for i, day in enumerate(days):
        col_c = "#ef4444" if i == 0 else ("#2563eb" if i == 6 else "#475569")
        d.text((820 + i * col_w + 35, 295), day, font=get_font(18, bold=True), fill=col_c, anchor="mm")

    for row in range(5):
        for col in range(7):
            day_num = row * 7 + col - 1
            if 1 <= day_num <= 30:
                cx = 820 + col * col_w
                cy = 330 + row * 125
                is_sel = (day_num in [12, 13])
                if is_sel:
                    d.rounded_rectangle([cx, cy, cx + 76, cy + 110], radius=8, fill="#dbeafe", outline="#2563eb", width=3)
                    d.text((cx + 8, cy + 8), str(day_num), font=get_font(20, bold=True), fill="#1d4ed8")
                    d.rounded_rectangle([cx + 6, cy + 55, cx + 70, cy + 95], radius=6, fill="#2563eb")
                    d.text((cx + 38, cy + 75), "선택", font=get_font(16, bold=True), fill="#ffffff", anchor="mm")
                else:
                    d.rounded_rectangle([cx, cy, cx + 76, cy + 110], radius=8, fill="#ffffff", outline="#e2e8f0")
                    num_c = "#ef4444" if col == 0 else ("#2563eb" if col == 6 else "#334155")
                    d.text((cx + 8, cy + 8), str(day_num), font=get_font(18), fill=num_c)

    # 지시선 배치 (충돌 0%)
    draw_smart_pin(d, (172, 295), (172, 215), "①", "특근 종류 선택", color="#2563eb", sub_hint="주말엔 일반휴일 기본 선택!")
    draw_smart_pin(d, (1400, 480), (1485, 480), "②", "달력 날짜 콕 찍기", color="#0284c7", sub_hint="선택 시 파란색 강조!")
    draw_smart_pin(d, (675, 750), (740, 750), "③", "사유 및 프로젝트 적기", color="#ea580c")
    draw_smart_pin(d, (675, 882), (740, 882), "④", "[특근 저장] 누르면 끝!", color="#10b981", sub_hint="관리자에게 즉시 전송")

    path = IMG_DIR / "mockup_apply.png"
    img.save(path)
    return str(path)


# 3. 내 신청 내역 목업
def create_mockup_my_records():
    """삽화 3: 내 신청 내역 목업 (확인완료 상태 & 대체휴일 쉰 날 표시)"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  내 신청 내역", url="https://overtime.company.com/my-records")

    d.rounded_rectangle([40, 105, 1560, 1060], radius=16, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((80, 135), "[내가 신청한 특근 확인하기] (초록색 도장 쾅!)", font=get_font(28, bold=True), fill="#0f172a")
    d.text((80, 175), "관리자 승인 완료 여부 및 대체휴일로 쉰 날을 실시간으로 확인하고 수정/취소합니다.", font=get_font(18), fill="#64748b")

    # 검색 툴바
    d.rounded_rectangle([80, 215, 1520, 275], radius=10, fill="#f1f5f9")
    d.text((105, 235), "조회 기간: 2026-09-01 ~ 2026-09-30  |  구분: 전체  |  상태: 전체  |  검색어: [제어실]", font=get_font(17), fill="#334155")
    d.rounded_rectangle([1180, 225, 1320, 265], radius=6, fill="#2563eb")
    d.text((1250, 245), "[새로고침]", font=get_font(16, bold=True), fill="#ffffff", anchor="mm")
    d.rounded_rectangle([1340, 225, 1500, 265], radius=6, fill="#10b981")
    d.text((1420, 245), "총 2건 (3.0일)", font=get_font(16, bold=True), fill="#ffffff", anchor="mm")

    # 카드 1: 확인완료 & 대체휴일
    d.rounded_rectangle([80, 300, 1520, 545], radius=12, fill="#ffffff", outline="#cbd5e1", width=2)
    d.rectangle([80, 300, 95, 545], fill="#10b981")
    d.text((120, 325), "2026-09-12 (토) ~ 2026-09-13 (일)", font=get_font(24, bold=True), fill="#0f172a")
    d.rounded_rectangle([590, 320, 710, 365], radius=6, fill="#dcfce7", outline="#86efac")
    d.text((650, 342), "일반휴일", font=get_font(18, bold=True), fill="#15803d", anchor="mm")

    d.rounded_rectangle([1230, 320, 1490, 370], radius=8, fill="#10b981")
    d.text((1360, 345), "[확인완료] 승인완료", font=get_font(20, bold=True), fill="#ffffff", anchor="mm")

    d.text((120, 385), "• 상세사유: 1공장 제어설비 긴급 점검 및 라인 개선  |  근무장소: 1공장 제어실  |  신청일수: 2.0일", font=get_font(18), fill="#475569")

    d.rounded_rectangle([120, 430, 880, 515], radius=8, fill="#fef3c7", outline="#f59e0b", width=2)
    d.text((145, 450), "[휴가] 대체휴일 쉰 날: 2026-09-25 (금)  (1.0일 쉼 / 자동 차감 적용)", font=get_font(19, bold=True), fill="#b45309")
    d.text((145, 480), "※ 주말 특근 대신 평일에 하루 쉰 날로, 최종 정산 시 1.0일이 자동 차감 계산됩니다.", font=get_font(15), fill="#92400e")

    # v1.53: 실특근 완료 후 확정 피드백 버튼
    d.rounded_rectangle([1080, 445, 1245, 500], radius=6, fill="#0d9488")
    d.text((1162, 472), "🎯 [특근확정]", font=get_font(17, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([1260, 445, 1365, 500], radius=6, fill="#f1f5f9", outline="#cbd5e1")
    d.text((1312, 472), "[수정]", font=get_font(18, bold=True), fill="#334155", anchor="mm")
    d.rounded_rectangle([1385, 445, 1490, 500], radius=6, fill="#fee2e2", outline="#fca5a5")
    d.text((1437, 472), "[삭제]", font=get_font(18, bold=True), fill="#dc2626", anchor="mm")

    # 카드 2: 승인대기
    d.rounded_rectangle([80, 570, 1520, 790], radius=12, fill="#ffffff", outline="#cbd5e1", width=2)
    d.rectangle([80, 570, 95, 790], fill="#f59e0b")
    d.text((120, 595), "2026-09-20 (일)", font=get_font(24, bold=True), fill="#0f172a")
    d.rounded_rectangle([390, 590, 510, 635], radius=6, fill="#eff6ff", outline="#93c5fd")
    d.text((450, 612), "대체근무", font=get_font(18, bold=True), fill="#1d4ed8", anchor="mm")

    d.rounded_rectangle([1230, 590, 1490, 640], radius=8, fill="#fef3c7", outline="#f59e0b", width=2)
    d.text((1360, 615), "● 승인대기 (검토중)", font=get_font(20, bold=True), fill="#b45309", anchor="mm")

    d.text((120, 655), "• 상세사유: 전장배선 라인 사전 포설 및 케이블 결선  |  근무장소: 배선실  |  신청일수: 1.0일", font=get_font(18), fill="#475569")

    d.rounded_rectangle([1260, 695, 1365, 750], radius=6, fill="#f1f5f9", outline="#cbd5e1")
    d.text((1312, 722), "[수정]", font=get_font(18, bold=True), fill="#334155", anchor="mm")
    d.rounded_rectangle([1385, 695, 1490, 750], radius=6, fill="#fee2e2", outline="#fca5a5")
    d.text((1437, 722), "[삭제]", font=get_font(18, bold=True), fill="#dc2626", anchor="mm")

    # 하단 가이드
    d.rounded_rectangle([80, 825, 1520, 1025], radius=12, fill="#f8fafc", outline="#cbd5e1", width=2)
    d.text((120, 850), "★ [내 신청 내역 & 4단계 라이프사이클 핵심 가이드]", font=get_font(20, bold=True), fill="#0f172a")
    d.text((120, 885), "1. 2단계 [승인완료]: 부서 관리자가 특근 계획을 승인한 상태입니다.", font=get_font(16), fill="#334155")
    d.text((120, 915), "2. 3단계 [특근확정]: 실제 특근을 마친 후 청록색 [🎯 특근확정] 버튼을 눌러 이행 피드백을 전달합니다.", font=get_font(16, bold=True), fill="#0d9488")
    d.text((120, 945), "3. 4단계 [검토완료]: 관리자가 최종 확인을 마치면 보라색 [검토완료]로 최종 정산 마감됩니다.", font=get_font(16), fill="#6d28d9")
    d.text((120, 975), "4. 상단 [💡 건의사항 소통함]: 100% 무기명으로 불편사항이나 개선요청을 자유롭게 건의할 수 있습니다.", font=get_font(16), fill="#d97706")

    # 지시선 배치 (충돌 0%)
    draw_smart_pin(d, (1230, 345), (1050, 345), "①", "초록색 승인 완료 도장!", color="#10b981", sub_hint="2단계 승인 완료")
    draw_smart_pin(d, (1162, 445), (1020, 395), "②", "실특근 후 [특근확정] 클릭!", color="#0d9488", sub_hint="3단계 확정 피드백 전달")
    draw_smart_pin(d, (880, 475), (960, 520), "③", "대체휴일 쉰 날 표시", color="#f59e0b", sub_hint="내가 쉰 날짜 확인!")
    draw_smart_pin(d, (1230, 615), (1050, 615), "④", "노란색 승인대기 표시", color="#ea580c", sub_hint="1단계 관리자 검토 중")

    path = IMG_DIR / "mockup_my.png"
    img.save(path)
    return str(path)


# 4. 관리자 월간 캘린더 목업
def create_mockup_admin_calendar():
    """삽화 4: 관리자 월간 캘린더 & 부서별 격리 관리 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  관리자 모드 (월간 캘린더)", url="https://overtime.company.com/admin/calendar")

    d.rounded_rectangle([40, 105, 1560, 1060], radius=16, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((80, 135), "[관리자 모드] 월간 캘린더 및 일괄 승인 (도장 쾅쾅!)", font=get_font(28, bold=True), fill="#0f172a")

    d.text((80, 185), "소속팀 필터 선택:", font=get_font(18, bold=True), fill="#64748b")
    teams = ["[전체보기]", "제어실", "전장배선팀", "전장설계팀", "PLC제어팀"]
    for i, t in enumerate(teams):
        bg = "#2563eb" if i == 1 else "#f1f5f9"
        tc = "#ffffff" if i == 1 else "#475569"
        bx = 240 + i * 155
        d.rounded_rectangle([bx, 175, bx + 140, 220], radius=8, fill=bg)
        d.text((bx + 70, 197), t, font=get_font(17, bold=True), fill=tc, anchor="mm")

    # 좌측: 캘린더 뷰 (너비 720px)
    d.rounded_rectangle([80, 240, 800, 940], radius=12, fill="#f8fafc", outline="#cbd5e1")
    d.text((105, 265), "2026년 9월 [제어실] 특근 달력", font=get_font(22, bold=True), fill="#0f172a")

    d.rounded_rectangle([105, 310, 390, 480], radius=8, fill="#ffffff", outline="#cbd5e1")
    d.text((120, 325), "5 (토)", font=get_font(18, bold=True), fill="#2563eb")
    d.rounded_rectangle([120, 360, 375, 410], radius=6, fill="#10b981")
    d.text((247, 385), "정진규 (일반휴일 1일)", font=get_font(16, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([420, 310, 705, 480], radius=8, fill="#ffffff", outline="#cbd5e1")
    d.text((435, 325), "6 (일)", font=get_font(18, bold=True), fill="#ef4444")
    d.text((435, 385), "- 근무자 없음 -", font=get_font(16), fill="#94a3b8")

    d.rounded_rectangle([105, 510, 705, 730], radius=10, fill="#eff6ff", outline="#2563eb", width=3)
    d.text((125, 525), "12 (토)  [보너스 1명 부여 일자]", font=get_font(20, bold=True), fill="#1d4ed8")
    d.rounded_rectangle([125, 570, 390, 620], radius=6, fill="#10b981")
    d.text((257, 595), "정진규 (일반휴일 1일)", font=get_font(17, bold=True), fill="#ffffff", anchor="mm")
    d.rounded_rectangle([420, 570, 685, 620], radius=6, fill="#f59e0b", outline="#7c3aed", width=2)
    d.text((552, 595), "[보너스] 김철수 (대휴)", font=get_font(17, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([105, 760, 390, 915], radius=8, fill="#ffffff", outline="#cbd5e1")
    d.text((120, 775), "13 (일)", font=get_font(18, bold=True), fill="#ef4444")
    d.rounded_rectangle([120, 810, 375, 860], radius=6, fill="#10b981")
    d.text((247, 835), "정진규 (일반휴일 1일)", font=get_font(16, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([420, 760, 705, 915], radius=8, fill="#ffffff", outline="#cbd5e1")
    d.text((435, 775), "20 (일)", font=get_font(18, bold=True), fill="#ef4444")
    d.rounded_rectangle([435, 810, 690, 860], radius=6, fill="#3b82f6")
    d.text((562, 835), "이영희 (대체근무 1일)", font=get_font(16, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([80, 960, 800, 1025], radius=8, fill="#f1f5f9")
    d.text((100, 985), "범례: ● 일반휴일(초록)  ● 대체근무(파랑)  ● 대휴사용(주황)  ● 보너스(보라테두리)", font=get_font(15, bold=True), fill="#334155")

    # 우측: 당일 명단 및 확인 패널
    d.rounded_rectangle([830, 240, 1520, 1025], radius=12, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((865, 265), "9월 12일 (토) 제어실 근무자 명단 (2명)", font=get_font(22, bold=True), fill="#0f172a")

    d.rounded_rectangle([865, 310, 1485, 395], radius=10, fill="#10b981")
    d.text((1175, 352), "[일괄승인] 당일 전원 일괄 확인 (원클릭 승인!)", font=get_font(24, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([865, 420, 1485, 525], radius=8, fill="#f8fafc", outline="#e2e8f0")
    d.text((890, 445), "정진규 (실장) - 1공장 제어설비 정기 점검", font=get_font(20, bold=True), fill="#0f172a")
    d.text((890, 485), "특근분류: 일반휴일 (1.0일)  |  상태: [확인완료] 승인완료", font=get_font(17), fill="#10b981")

    d.rounded_rectangle([865, 550, 1485, 655], radius=8, fill="#f8fafc", outline="#e2e8f0")
    d.text((890, 575), "김철수 (선임) - [보너스 부여 대상] + 대체휴무", font=get_font(20, bold=True), fill="#7c3aed")
    d.text((890, 615), "특근분류: 대체근무 (1.0일)  |  보너스: 부여됨  |  대휴: 1.0일 차감", font=get_font(17), fill="#b45309")

    d.rounded_rectangle([865, 685, 1485, 1000], radius=10, fill="#fefce8", outline="#facc15", width=2)
    d.text((895, 715), "★ [관리자 승인 및 보안 격리 운영 팁]", font=get_font(19, bold=True), fill="#854d0e")
    d.text((895, 755), "1. 팀관리자는 본인 부서 일정만 관리하며, 총괄관리자 일정은 완벽히 격리 차단됩니다.", font=get_font(16), fill="#a16207")
    d.text((895, 795), "2. 초록색 [당일 전원 일괄 확인]을 누르면 해당 날짜 근무자 전원이 즉시 승인됩니다.", font=get_font(16), fill="#a16207")
    d.text((895, 835), "3. 부여된 보너스는 사원 화면에는 마스킹되며, 관리자 화면 및 엑셀에만 명기됩니다.", font=get_font(16), fill="#a16207")

    # 지시선 배치 (충돌 0%)
    draw_smart_pin(d, (315, 175), (550, 135), "①", "우리 팀만 쏙 골라보기", color="#2563eb")
    draw_smart_pin(d, (552, 620), (552, 680), "②", "보너스 인원 달력 표시", color="#7c3aed", sub_hint="보라색 강조 테두리")
    draw_smart_pin(d, (1175, 310), (1175, 215), "③", "[당일 전원 일괄 확인] 쾅!", color="#10b981", sub_hint="원클릭 일괄 승인")
    draw_smart_pin(d, (1485, 470), (1350, 560), "④", "당일 일한 상세 내용 확인", color="#0284c7")

    path = IMG_DIR / "mockup_admin_cal.png"
    img.save(path)
    return str(path)


# 5. 팀원 및 소속팀 관리 목업
def create_mockup_user_mgmt():
    """삽화 5: 팀원 관리 및 소속팀 추가/삭제 완전 영구정리 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  팀원 및 소속팀 관리", url="https://overtime.company.com/admin/users")

    d.rounded_rectangle([40, 105, 1560, 1060], radius=16, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((80, 125), "[팀원 및 소속팀 척척 관리하기]", font=get_font(28, bold=True), fill="#0f172a")

    # 상단 툴바 (너비 축소로 여백 확보)
    d.rounded_rectangle([80, 175, 320, 235], radius=8, fill="#f59e0b")
    d.text((200, 205), "[[팀] 소속팀 관리]", font=get_font(18, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([340, 175, 580, 235], radius=8, fill="#2563eb")
    d.text((460, 205), "[[엑셀] 대량 등록]", font=get_font(18, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([600, 175, 840, 235], radius=8, fill="#10b981")
    d.text((720, 205), "[[+] 신규 팀원 추가]", font=get_font(18, bold=True), fill="#ffffff", anchor="mm")

    # 좌측: 테이블 목업
    d.rounded_rectangle([80, 270, 860, 1010], radius=10, fill="#f8fafc", outline="#cbd5e1")
    d.rectangle([80, 270, 860, 330], fill="#e2e8f0")
    d.text((105, 290), "사번", font=get_font(18, bold=True), fill="#334155")
    d.text((230, 290), "이름", font=get_font(18, bold=True), fill="#334155")
    d.text((340, 290), "소속팀", font=get_font(18, bold=True), fill="#334155")
    d.text((490, 290), "관리자권한", font=get_font(18, bold=True), fill="#334155")
    d.text((690, 290), "관리", font=get_font(18, bold=True), fill="#334155")

    members = [
        ("ADMIN01", "슈퍼관리자", "제어실", "총괄 ON", "[수정]"),
        ("113019", "정진규", "제어실", "OFF", "[수정] [삭제]"),
        ("2024001", "김철수", "기술연구팀", "팀관리 ON", "[수정] [삭제]"),
        ("2026002", "이영희", "전장배선팀", "OFF", "[수정] [삭제]")
    ]
    for idx, (m_id, m_name, m_team, m_adm, m_act) in enumerate(members):
        my = 360 + idx * 85
        d.text((105, my), m_id, font=get_font(18, bold=True), fill="#0f172a")
        d.text((230, my), m_name, font=get_font(18, bold=True), fill="#0f172a")
        d.text((340, my), m_team, font=get_font(18), fill="#2563eb")
        d.text((495, my), m_adm, font=get_font(18, bold=True), fill="#16a34a" if "ON" in m_adm else "#64748b")
        d.rounded_rectangle([680, my - 6, 820, my + 34], radius=6, fill="#fee2e2")
        d.text((750, my + 14), m_act, font=get_font(16, bold=True), fill="#dc2626", anchor="mm")

    # 우측: 모달창
    d.rounded_rectangle([900, 270, 1520, 1010], radius=12, fill="#ffffff", outline="#f59e0b", width=3)
    d.text((935, 305), "소속팀 관리 (새로운 팀 만들기 / 정리)", font=get_font(24, bold=True), fill="#b45309")
    d.text((935, 355), "새 팀 이름 입력:", font=get_font(18), fill="#475569")
    d.rounded_rectangle([935, 395, 1340, 460], radius=8, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((955, 415), "PLC제어팀", font=get_font(20, bold=True), fill="#0f172a")
    d.rounded_rectangle([1360, 395, 1490, 460], radius=8, fill="#f59e0b")
    d.text((1425, 427), "+ 추가", font=get_font(19, bold=True), fill="#ffffff", anchor="mm")

    d.text((935, 495), "현재 등록된 소속팀 목록: (삭제된 팀은 영구 정리)", font=get_font(18, bold=True), fill="#334155")
    teams_list = ["• 제어실", "• 전장배선팀", "• 전장설계팀", "• 기술연구팀", "• PLC제어팀"]
    for i, tm in enumerate(teams_list):
        d.text((955, 540 + i * 45), tm, font=get_font(19, bold=True), fill="#2563eb")

    d.rounded_rectangle([935, 800, 1490, 970], radius=10, fill="#ecfdf5", outline="#10b981", width=2)
    d.text((955, 825), "★ [소속팀 및 권한 정리 보증]", font=get_font(19, bold=True), fill="#065f46")
    d.text((955, 870), "1. 삭제된 팀은 시스템에서 완전히 소멸되어 다시 생겨나지 않습니다.", font=get_font(16), fill="#047857")
    d.text((955, 915), "2. 팀관리자는 본인 팀원 정보만 수정 가능하며 타 팀 통제는 차단됩니다.", font=get_font(16), fill="#047857")

    # 지시선 배치 (충돌 0%)
    draw_smart_pin(d, (200, 235), (200, 260), "①", "[소속팀 관리] 버튼", color="#f59e0b")
    draw_smart_pin(d, (1425, 395), (1425, 230), "②", "새로운 팀 1초 만에 추가", color="#0284c7")
    draw_smart_pin(d, (1150, 720), (1150, 770), "③", "삭제된 팀은 다시 안 나와요!", color="#10b981", sub_hint="완벽한 부서 정리")
    draw_smart_pin(d, (750, 640), (750, 715), "④", "팀원 수정 및 삭제 & 권한", color="#dc2626")

    path = IMG_DIR / "mockup_user_mgmt.png"
    img.save(path)
    return str(path)


# 6. 실특근 산정 목업
def create_mockup_settlement():
    """삽화 6: 실특근 정산 및 자동 계산 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  실특근 정산 및 자동 계산 (v1.53)", url="https://overtime.company.com/admin/settlement")

    d.rounded_rectangle([40, 105, 1560, 1060], radius=16, fill="#ffffff", outline="#cbd5e1", width=2)

    # 공식 배너
    d.rounded_rectangle([80, 125, 1520, 235], radius=12, fill="#ecfdf5", outline="#10b981", width=3)
    d.text((115, 145), "★ 컴퓨터가 1초 만에 자동 산출하는 진짜 일한 날(최종 실특근) 공식! (v1.53 개편)", font=get_font(22, bold=True), fill="#065f46")
    d.text((115, 185), "최종 실특근일 = 일반특근 - 사전차감 - (대체휴무 - 대체휴무 시 출장기간 내 사전차감)", font=get_font(20, bold=True), fill="#047857")

    # 팀 카드 (y축 280부터 시작하여 상단 여백 확보)
    d.rounded_rectangle([80, 280, 780, 440], radius=12, fill="#ffffff", outline="#cbd5e1", width=2)
    d.rectangle([80, 280, 95, 440], fill="#2563eb")
    d.text((120, 295), "[팀] 제어실", font=get_font(24, bold=True), fill="#0f172a")
    d.text((650, 300), "팀원: 3명", font=get_font(18), fill="#64748b")
    d.rounded_rectangle([120, 340, 750, 380], radius=6, fill="#f8fafc")
    d.text((135, 350), "총신청 7일  |  법정제외 0일  |  일반휴일 7일  |  대휴 0.0일", font=get_font(17), fill="#475569")
    d.text((120, 400), "★ 우리 팀 최종 실특근:", font=get_font(20, bold=True), fill="#065f46")
    d.text((370, 395), "7.0일 (일반 7 - 대휴 0)", font=get_font(24, bold=True), fill="#059669")

    d.rounded_rectangle([820, 280, 1520, 440], radius=12, fill="#ffffff", outline="#cbd5e1", width=2)
    d.rectangle([820, 280, 835, 440], fill="#f59e0b")
    d.text((860, 295), "[팀] 기술연구팀", font=get_font(24, bold=True), fill="#0f172a")
    d.text((1390, 300), "팀원: 1명", font=get_font(18), fill="#64748b")
    d.rounded_rectangle([860, 340, 1490, 380], radius=6, fill="#f8fafc")
    d.text((875, 350), "총신청 6일  |  법정제외 1일  |  일반휴일 5일  |  대휴 1.0일", font=get_font(17), fill="#475569")
    d.text((860, 400), "★ 우리 팀 최종 실특근:", font=get_font(20, bold=True), fill="#065f46")
    d.text((1110, 395), "4.0일 (일반 5 - 대휴 1.0)", font=get_font(24, bold=True), fill="#059669")

    # 테이블
    d.rounded_rectangle([80, 465, 1520, 890], radius=12, fill="#ffffff", outline="#cbd5e1", width=2)
    d.rectangle([80, 465, 1520, 525], fill="#f1f5f9")
    d.text((105, 485), "사번", font=get_font(18, bold=True), fill="#334155")
    d.text((230, 485), "이름", font=get_font(18, bold=True), fill="#334155")
    d.text((350, 485), "소속팀", font=get_font(18, bold=True), fill="#334155")
    d.text((500, 485), "일반휴일", font=get_font(18, bold=True), fill="#2563eb")
    d.text((640, 485), "대휴사용", font=get_font(18, bold=True), fill="#ea580c")
    d.text((790, 485), "사전차감잔여", font=get_font(18, bold=True), fill="#64748b")
    d.text((950, 485), "보너스", font=get_font(18, bold=True), fill="#7c3aed")
    d.text((1070, 485), "★ 최종실특근", font=get_font(19, bold=True), fill="#059669")
    d.text((1280, 485), "계산 산출 공식", font=get_font(18, bold=True), fill="#64748b")

    d.text((105, 560), "113019", font=get_font(19), fill="#0f172a")
    d.text((230, 560), "정진규", font=get_font(19, bold=True), fill="#0f172a")
    d.text((350, 560), "제어실", font=get_font(19), fill="#475569")
    d.text((530, 560), "7", font=get_font(20, bold=True), fill="#2563eb")
    d.text((670, 560), "0.0", font=get_font(20), fill="#64748b")
    d.text((830, 560), "0", font=get_font(20), fill="#64748b")
    d.text((970, 560), "0건", font=get_font(19), fill="#64748b")
    d.text((1090, 560), "7.0일", font=get_font(22, bold=True), fill="#059669")
    d.text((1280, 560), "7 - 0 = 7.0일", font=get_font(18), fill="#059669")

    d.text((105, 660), "2024001", font=get_font(19), fill="#0f172a")
    d.text((230, 660), "김철수", font=get_font(19, bold=True), fill="#0f172a")
    d.text((350, 660), "기술연구팀", font=get_font(19), fill="#475569")
    d.text((530, 660), "5", font=get_font(20, bold=True), fill="#2563eb")
    d.text((670, 660), "1.0", font=get_font(20, bold=True), fill="#ea580c")
    d.text((830, 660), "0", font=get_font(20), fill="#64748b")
    d.text((970, 660), "1건", font=get_font(19, bold=True), fill="#7c3aed")
    d.text((1090, 660), "4.0일", font=get_font(22, bold=True), fill="#059669")
    d.text((1280, 660), "5 - 1.0 = 4.0일 (대휴 차감!)", font=get_font(18, bold=True), fill="#059669")

    d.rounded_rectangle([1180, 770, 1490, 840], radius=8, fill="#2563eb")
    d.text((1335, 805), "[[엑셀] 정산표 내보내기]", font=get_font(20, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([80, 920, 1520, 1025], radius=10, fill="#fefce8", outline="#facc15", width=2)
    d.text((115, 945), "★ [실특근 정산 및 자동 계산 핵심 원리]", font=get_font(19, bold=True), fill="#854d0e")
    d.text((115, 985), "총 특근일수 = 대체근무 + 법정휴일 + 일반휴일 (대체휴무 제외 표준화) / 엑셀 보너스 건수 독립 산출", font=get_font(16), fill="#a16207")

    # 지시선 배치 (충돌 0%)
    draw_smart_pin(d, (800, 125), (800, 95), "①", "직관적인 실특근 공식", color="#10b981")
    draw_smart_pin(d, (450, 280), (450, 245), "②", "팀별 실특근 합계 요약", color="#2563eb")
    draw_smart_pin(d, (1090, 660), (1090, 730), "③", "개인별 뺄셈 자동 계산", color="#ea580c", sub_hint="5 - 1.0 = 4.0일")
    draw_smart_pin(d, (1335, 840), (1335, 890), "④", "엑셀 파일 즉시 저장", color="#0284c7")

    path = IMG_DIR / "mockup_settlement.png"
    img.save(path)
    return str(path)


# 7. 엑셀 27 개편 목업
def create_mockup_excel_27():
    """삽화 7: 개편된 엑셀 보고서 목업 (수식 자동합계 & 4대 휴일수)"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  개편된 엑셀 보고서 (.xlsx)", url="https://overtime.company.com/export/excel")

    d.rounded_rectangle([40, 105, 1560, 1060], radius=16, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((80, 125), "[개편된 엑셀 보고서] (수식 자동합계 & 4대 휴일수)", font=get_font(28, bold=True), fill="#0f172a")

    # 탭
    d.rounded_rectangle([80, 175, 450, 230], radius=8, fill="#1e3a8a")
    d.text((265, 202), "[시트1: 휴일일자별_특근현황]", font=get_font(18, bold=True), fill="#ffffff", anchor="mm")
    d.rounded_rectangle([470, 175, 840, 230], radius=8, fill="#0284c7")
    d.text((655, 202), "[시트2: 개인별_휴일합산_정산표]", font=get_font(18, bold=True), fill="#ffffff", anchor="mm")
    d.rounded_rectangle([860, 175, 1180, 230], radius=8, fill="#475569")
    d.text((1020, 202), "[시트3: 보너스_부여내역]", font=get_font(18, bold=True), fill="#ffffff", anchor="mm")

    # 시트 1 테이블
    d.rounded_rectangle([80, 260, 1520, 540], radius=10, fill="#ffffff", outline="#cbd5e1")
    d.rectangle([80, 260, 1520, 315], fill="#1e3a8a")
    d.text((105, 277), "휴일일자", font=get_font(17, bold=True), fill="#ffffff")
    d.text((265, 277), "요일", font=get_font(17, bold=True), fill="#ffffff")
    d.text((345, 277), "성명", font=get_font(17, bold=True), fill="#ffffff")
    d.text((475, 277), "사번", font=get_font(17, bold=True), fill="#ffffff")
    d.text((615, 277), "소속팀", font=get_font(17, bold=True), fill="#ffffff")
    d.text((775, 277), "특근분류", font=get_font(17, bold=True), fill="#ffffff")
    d.text((945, 277), "휴일일수(숫자만!)", font=get_font(17, bold=True), fill="#fef08a")
    d.text((1165, 277), "대체휴가일수", font=get_font(17, bold=True), fill="#fed7aa")
    d.text((1355, 277), "대휴사용일", font=get_font(17, bold=True), fill="#ffffff")

    d.text((105, 335), "2026-09-05", font=get_font(18), fill="#0f172a")
    d.text((275, 335), "토", font=get_font(18, bold=True), fill="#2563eb")
    d.text((345, 335), "정진규", font=get_font(18, bold=True), fill="#0f172a")
    d.text((475, 335), "113019", font=get_font(18), fill="#475569")
    d.text((615, 335), "제어실", font=get_font(18), fill="#475569")
    d.text((775, 335), "일반휴일", font=get_font(18), fill="#2563eb")
    d.text((1015, 335), "1", font=get_font(22, bold=True), fill="#b45309")
    d.text((1205, 335), "0.0", font=get_font(18), fill="#64748b")
    d.text((1375, 335), "-", font=get_font(18), fill="#64748b")

    d.text((105, 395), "2026-09-12", font=get_font(18), fill="#0f172a")
    d.text((275, 395), "토", font=get_font(18, bold=True), fill="#2563eb")
    d.text((345, 395), "김철수", font=get_font(18, bold=True), fill="#0f172a")
    d.text((475, 395), "2024001", font=get_font(18), fill="#475569")
    d.text((615, 395), "기술연구팀", font=get_font(18), fill="#475569")
    d.text((775, 395), "일반휴일", font=get_font(18), fill="#2563eb")
    d.text((1015, 395), "1", font=get_font(22, bold=True), fill="#b45309")
    d.text((1205, 395), "1.0", font=get_font(20, bold=True), fill="#ea580c")
    d.text((1355, 395), "2026-09-25", font=get_font(18, bold=True), fill="#059669")

    d.rectangle([80, 455, 1520, 515], fill="#fef3c7")
    d.text((105, 475), "합계 (엑셀 수식 자동계산)", font=get_font(18, bold=True), fill="#92400e")
    d.text((965, 475), "=SUM(G4:G5) -> 2", font=get_font(20, bold=True), fill="#b45309")
    d.text((1175, 475), "=SUM(H4:H5) -> 1.0", font=get_font(20, bold=True), fill="#ea580c")

    # 시트 2 테이블
    d.rounded_rectangle([80, 570, 1520, 820], radius=10, fill="#ffffff", outline="#0284c7")
    d.rectangle([80, 570, 1520, 625], fill="#0284c7")
    d.text((105, 587), "사번", font=get_font(17, bold=True), fill="#ffffff")
    d.text((235, 587), "성명", font=get_font(17, bold=True), fill="#ffffff")
    d.text((345, 587), "소속팀", font=get_font(17, bold=True), fill="#ffffff")
    d.text((495, 587), "대체근무(일)", font=get_font(17, bold=True), fill="#ffffff")
    d.text((675, 587), "법정휴일(일)", font=get_font(17, bold=True), fill="#ffffff")
    d.text((855, 587), "일반휴일(일)", font=get_font(17, bold=True), fill="#ffffff")
    d.text((1035, 587), "대휴사용(일)", font=get_font(17, bold=True), fill="#ffffff")
    d.text((1215, 587), "사전차감잔여", font=get_font(17, bold=True), fill="#ffffff")
    d.text((1375, 587), "★ 최종실특근", font=get_font(18, bold=True), fill="#fef08a")

    d.text((105, 655), "2024001", font=get_font(18), fill="#0f172a")
    d.text((235, 655), "김철수", font=get_font(18, bold=True), fill="#0f172a")
    d.text((345, 655), "기술연구팀", font=get_font(18), fill="#475569")
    d.text((545, 655), "0", font=get_font(19), fill="#64748b")
    d.text((725, 655), "1", font=get_font(19), fill="#64748b")
    d.text((905, 655), "5", font=get_font(20, bold=True), fill="#2563eb")
    d.text((1085, 655), "1.0", font=get_font(20, bold=True), fill="#ea580c")
    d.text((1265, 655), "0", font=get_font(19), fill="#64748b")
    d.rounded_rectangle([1365, 642, 1490, 690], radius=6, fill="#dcfce7")
    d.text((1427, 666), "4.0일", font=get_font(20, bold=True), fill="#15803d", anchor="mm")

    d.rounded_rectangle([80, 850, 1520, 1025], radius=10, fill="#fefce8", outline="#facc15", width=2)
    d.text((115, 875), "★ [개편된 엑셀 양식 핵심 활용법]", font=get_font(19, bold=True), fill="#854d0e")
    d.text((115, 915), "1. [27-1] 일한 날짜 순서대로 차례차례 펼쳐지며 대휴 사용일이 바로 연결됩니다.", font=get_font(16), fill="#a16207")
    d.text((115, 955), "2. [27-2] 문자 '일' 대신 순수 숫자 '1'만 입력되어 마우스 드래그 합계(SUM)가 즉시 계산됩니다.", font=get_font(16), fill="#a16207")
    d.text((115, 990), "3. [27-3] Sheet 2에 4가지 휴일수와 최종 실특근일, Sheet 3에 보너스 내역이 완벽히 정리됩니다.", font=get_font(16), fill="#a16207")

    # 지시선 배치 (충돌 0%)
    draw_smart_pin(d, (265, 175), (265, 135), "①", "[27-1] 날짜순 쫙 전개", color="#0284c7")
    draw_smart_pin(d, (1015, 277), (1350, 205), "②", "[27-2] 순수 숫자 '1'만 쏙", color="#b45309", sub_hint="수식 합계 자동화")
    draw_smart_pin(d, (655, 570), (655, 535), "③", "[27-3] 개인별 4대 휴일수 합산", color="#16a34a")
    draw_smart_pin(d, (1427, 690), (1427, 755), "④", "일반휴일 - 대휴 = 실특근", color="#10b981")

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
    """
    슬라이드 너비 13.333" x 7.5" 와이드스크린 레이아웃:
    - 좌측: 설명 카드 (너비 3.8인치, 폰트 12.5pt/10.5pt)
    - 우측: 대형 고화질 화면 목업 삽화 (너비 8.25인치, 높이 5.65인치, 화면 면적 +72% 대폭 확대)
    """
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 상단 헤더 (너비 확장)
    header_tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12.2), Inches(1.1))
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

    # 좌측: 지시선 번호 대응 설명 카드 (너비 3.8인치, 높이 5.65인치)
    card_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.55), Inches(3.8), Inches(5.65))
    card_box.fill.solid()
    card_box.fill.fore_color.rgb = COLOR_BG_CARD
    card_box.line.color.rgb = COLOR_BORDER

    card_tb = slide.shapes.add_textbox(Inches(0.75), Inches(1.7), Inches(3.5), Inches(5.35))
    ctf = card_tb.text_frame
    ctf.word_wrap = True

    for idx, (item_title, item_desc) in enumerate(items):
        p_it = ctf.paragraphs[0] if idx == 0 else ctf.add_paragraph()
        p_it.text = item_title
        p_it.font.size = Pt(12.5)
        p_it.font.bold = True
        p_it.font.color.rgb = COLOR_ACCENT
        p_it.space_after = Pt(2)

        p_id = ctf.add_paragraph()
        p_id.text = item_desc
        p_id.font.size = Pt(10.5)
        p_id.font.color.rgb = COLOR_TEXT_MAIN
        p_id.space_after = Pt(8)

    # 우측: 실제 화면 목업 삽화 이미지 삽입 (너비 8.25인치, 높이 5.65인치: 기존 6.6인치 대비 72% 면적 확대)
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path, Inches(4.55), Inches(1.55), width=Inches(8.25), height=Inches(5.65))
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

    tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.4), Inches(11.3), Inches(4.7))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p1_tag = tf1.paragraphs[0]
    p1_tag.text = "SMART OVERTIME SYSTEM v1.53  |  일반 사원 전용 공식 매뉴얼"
    p1_tag.font.size = Pt(14)
    p1_tag.font.bold = True
    p1_tag.font.color.rgb = RGBColor(253, 186, 116)
    p1_tag.space_after = Pt(14)

    p1_title = tf1.add_paragraph()
    p1_title.text = "스마트 특근 관리 시스템\n사용자 전용 기능 매뉴얼 (v1.53)"
    p1_title.font.size = Pt(36)
    p1_title.font.bold = True
    p1_title.font.color.rgb = RGBColor(255, 255, 255)
    p1_title.space_after = Pt(18)

    p1_sub = tf1.add_paragraph()
    p1_sub.text = "승인/확정 독립 피드백, 100% 무기명 건의 소통함, 접속 주소 및 접속 방법, 서버 초기화 대기 안내, 사원번호 간편 입장, 스마트 달력 신청까지 누구나 쉽게 따라 할 수 있는 사용자 공식 가이드입니다."
    p1_sub.font.size = Pt(15)
    p1_sub.font.color.rgb = RGBColor(203, 213, 225)
    p1_sub.space_after = Pt(24)

    p1_auth = tf1.add_paragraph()
    p1_auth.text = "배포 버전: v1.53 (2026-09-21)  |  PC 모니터 & 스마트폰(모바일) 완벽 지원"
    p1_auth.font.size = Pt(13)
    p1_auth.font.color.rgb = RGBColor(148, 163, 184)

    # --- SLIDE 2: 1단계 - 접속 주소 체계 및 접속 방법 ---
    s2 = prs.slides.add_slide(blank_layout)
    header_tb2 = s2.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf2 = header_tb2.text_frame
    htf2.word_wrap = True
    p2_step = htf2.paragraphs[0]
    p2_step.text = "STEP 01  |  외부 접속망(HTTPS) 및 모바일 접속 가이드"
    p2_step.font.size = Pt(12)
    p2_step.font.bold = True
    p2_step.font.color.rgb = COLOR_ACCENT

    p2_title = htf2.add_paragraph()
    p2_title.text = "외부 접속망(HTTPS) 및 모바일 전용 QR 접속 가이드"
    p2_title.font.size = Pt(22)
    p2_title.font.bold = True
    p2_title.font.color.rgb = COLOR_PRIMARY

    p2_sub = htf2.add_paragraph()
    p2_sub.text = "사내외 어디서나(LTE/5G/자택/PC) 지정된 외부 보안 링크와 전용 QR 코드로 간편하게 접속합니다."
    p2_sub.font.size = Pt(13)
    p2_sub.font.color.rgb = COLOR_TEXT_MUTED

    access_methods = [
        ("1. 외부 보안 링크 접속 (PC / 태블릿)",
         "https://xxxx.trycloudflare.com\n(관리자 발급 공식 외부 보안 링크)",
         "• 사내외 어디서나 PC 및 노트북 웹브라우저(Chrome, Edge 권장) 주소창에 보안 HTTPS 링크를 입력하여 접속합니다.\n"
         "• 별도의 사내망 연결이나 VPN 설치 없이 24시간 안전한 전송 구간 암호화(HTTPS) 터널을 통해 어디서든 열립니다.\n"
         "• 브라우저 상단 즐겨찾기(북마크 ★)에 등록해 두시면 매번 주소를 입력하지 않고 원클릭으로 즉시 접속됩니다.",
         "#eff6ff", "#2563eb"),
        ("2. 외부 모바일 QR 스캔 접속",
         "화면 상단 [📱 모바일 QR] 스캔\n(스마트폰 LTE / 5G 전용)",
         "• 스마트폰 기본 카메라 앱을 켜고 화면 상단 [📱 모바일 접속 QR]을 비추면 외부망 접속 링크가 즉시 나타납니다.\n"
         "• 사내 Wi-Fi 연결 여부와 상관없이 LTE/5G 데이터망으로 언제 어디서나 이동 중에도 특근 신청 및 조회가 가능합니다.\n"
         "• 스마트폰에서 신청한 내역은 PC 관리자 화면 및 대시보드와 100% 실시간 동일하게 동기화됩니다.",
         "#ecfdf5", "#10b981"),
        ("3. 스마트폰 홈 화면 추가 (앱 모드)",
         "브라우저 메뉴 ▶ [홈 화면에 추가]\n(바탕화면 원클릭 간편 실행)",
         "• 스마트폰 브라우저 메뉴에서 [홈 화면에 추가]를 터치하시면 스마트폰 바탕화면에 전용 앱 아이콘이 생성됩니다.\n"
         "• 매번 주소를 치거나 QR을 찍을 필요 없이 앱처럼 터치 한 번으로 1초 만에 시스템에 바로 입장할 수 있습니다.\n"
         "• 로그인 세션이 안전하게 유지되어 언제든 빠르게 본인 특근 일정과 승인 도장을 확인하실 수 있습니다.",
         "#fefce8", "#d97706")
    ]

    for i, (m_title, m_addr, m_desc, m_bg, m_border) in enumerate(access_methods):
        mx = Inches(0.8 + i * 4.0)
        m_box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, mx, Inches(1.65), Inches(3.7), Inches(5.4))
        m_box.fill.solid()
        m_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        m_box.line.color.rgb = RGBColor(203, 213, 225)
        m_box.line.width = Pt(2)

        # 상단 뱃지
        b_box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, mx + Inches(0.25), Inches(1.85), Inches(3.2), Inches(0.45))
        b_box.fill.solid()
        b_box.fill.fore_color.rgb = RGBColor(241, 245, 249)
        b_box.line.color.rgb = RGBColor(203, 213, 225)
        btf = b_box.text_frame
        bp = btf.paragraphs[0]
        bp.text = f"접속 방법 0{i+1}"
        bp.font.size = Pt(11)
        bp.font.bold = True
        bp.font.color.rgb = RGBColor(234, 88, 12)
        bp.alignment = PP_ALIGN.CENTER

        # 카드 텍스트 박스
        m_tb = s2.shapes.add_textbox(mx + Inches(0.2), Inches(2.4), Inches(3.3), Inches(4.5))
        mtf = m_tb.text_frame
        mtf.word_wrap = True

        mp1 = mtf.paragraphs[0]
        mp1.text = m_title
        mp1.font.size = Pt(17)
        mp1.font.bold = True
        mp1.font.color.rgb = COLOR_PRIMARY
        mp1.space_after = Pt(10)

        # 주소 하이라이트 박스
        mp_addr = mtf.add_paragraph()
        mp_addr.text = f"▶ {m_addr}"
        mp_addr.font.size = Pt(12)
        mp_addr.font.bold = True
        mp_addr.font.color.rgb = RGBColor(37, 99, 235)
        mp_addr.space_after = Pt(12)

        mp2 = mtf.add_paragraph()
        mp2.text = m_desc
        mp2.font.size = Pt(11.5)
        mp2.font.color.rgb = COLOR_TEXT_MAIN
        mp2.line_spacing = 1.35

    # --- SLIDE 3: 2단계 - 최초 접속 및 서버 리부팅 시 초기화 대기 안내 ---
    s3_data = [
        ("① 외부 보안 접속 주소 입력 & 연결", "배포된 외부 보안 링크(HTTPS)를 브라우저에 입력하여 첫 접속을 시작합니다."),
        ("② 약 10~20초 초기 구동 대기", "서버 최초 기동 또는 PC 리부팅 직후 1회에 한해 보안 터널 및 DB 초기화(약 10~20초) 대기 시간이 발생합니다."),
        ("③ DB 무결성 & 자동 백업 진행", "데이터 유실 방지를 위한 SQLite WAL 무결성 검증 및 백업이 안전하게 수행됩니다."),
        ("④ 새로고침(F5) 연타 금지!", "외부망 연결 초기화 중 새로고침을 연타하지 마시고 잠시 대기하시면 로그인 화면으로 자동 전환됩니다.")
    ]
    add_visual_slide(prs, "02", "최초 접속 및 서버 리부팅 시 초기화 대기 안내", "서버 최초 기동 또는 리부팅 직후에는 약 10~20초 동안 안전 초기화가 진행되므로 잠시 기다려주세요.", s3_data, images["init"])

    # --- SLIDE 4: 3단계 - 사번만 넣고 슝 들어가기 ---
    s4_data = [
        ("① v1.53 최신 버전 배지 확인", "화면 오른쪽 위에 청록색 [v1.53] 배지가 보이면 최신 버전입니다. 클릭 시 신규 릴리즈 이력이 표시됩니다."),
        ("② 사원번호 6자리 숫자 입력", "본인의 6자리 사원번호를 입력창에 기재합니다. 복잡한 비밀번호 없이 빠르게 입장 가능합니다."),
        ("③ [입장하기] 원클릭 이동", "입력 후 [입장하기]를 누르면 등록된 이름과 소속팀이 자동 확인되어 대시보드로 즉시 입장합니다."),
        ("④ 미등록 사번 1초 즉시 등록", "처음 방문한 사번은 신규 등록창이 나타나며, 성명과 소속 부서를 선택하면 즉시 등록되어 입장합니다.")
    ]
    add_visual_slide(prs, "03", "사원번호로 간편 입장하기", "비밀번호 없이 사원번호 6자리 입력만으로 빠르고 안전하게 입장합니다.", s4_data, images["login"])

    # --- SLIDE 5: 4단계 - 재입장 시 핵심 유의사항 가이드 ---
    s5 = prs.slides.add_slide(blank_layout)
    header_tb5 = s5.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf5 = header_tb5.text_frame
    htf5.word_wrap = True
    p5_step = htf5.paragraphs[0]
    p5_step.text = "STEP 04  |  재입장(재접속) 시 핵심 유의사항 가이드"
    p5_step.font.size = Pt(12)
    p5_step.font.bold = True
    p5_step.font.color.rgb = COLOR_ACCENT

    p5_title = htf5.add_paragraph()
    p5_title.text = "다시 접속(재입장)할 때 꼭 알아두어야 할 4대 핵심 가이드"
    p5_title.font.size = Pt(22)
    p5_title.font.bold = True
    p5_title.font.color.rgb = COLOR_PRIMARY

    p5_sub = htf5.add_paragraph()
    p5_sub.text = "사번 자동 완성, 멀티 디바이스 실시간 동기화, 최신 상태 갱신 및 공용 PC 보안 수칙을 확인하세요."
    p5_sub.font.size = Pt(13)
    p5_sub.font.color.rgb = COLOR_TEXT_MUTED

    reentry_cards = [
        ("1. 사원번호 자동 완성 (원클릭 재입장)",
         "• 브라우저 로컬 저장소(LocalStorage)에 직전에 로그인했던 본인 사번 6자리가 안전하게 기억됩니다.\n"
         "• 창을 닫았다가 다시 열었을 때 사번을 재입력할 필요 없이 [입장하기] 버튼만 누르면 1초 만에 바로 진입합니다.\n"
         "• 다른 사번으로 로그인해야 할 때는 입력창의 숫자를 지우고 새 사번을 입력하시면 됩니다.",
         "#eff6ff", "#2563eb"),
        ("2. PC ↔ 스마트폰 실시간 100% 동기화",
         "• 스마트폰(모바일)에서 신청한 내역을 회사 PC에서 열어도, 반대로 PC에서 신청한 내역을 폰에서 열어도 동일합니다.\n"
         "• 중앙 데이터베이스(data/overtime.db)를 실시간 공유하므로 어떤 기기에서 재입장하더라도 본인 데이터가 완벽 유지됩니다.\n"
         "• 기기를 교체하더라도 회원가입 없이 본인 사번만 입력하면 기존 내역을 즉시 조회할 수 있습니다.",
         "#ecfdf5", "#10b981"),
        ("3. 재입장 시 [새로고침] 1회 권장 (상태 동기화)",
         "• 브라우저 탭을 백그라운드에 켜두었거나 다시 접속했을 때는 화면 상단 [🔄 새로고침] 버튼 또는 F5를 1회 눌러주세요.\n"
         "• 부서 관리자가 승인 도장을 새로 찍었거나 팀원 일정이 변경되었을 때 1초 만에 최신 상태로 화면이 새로고침됩니다.\n"
         "• 네트워크 일시 지연이 발생할 때도 상단 새로고침 버튼을 누르면 즉시 정상 회복됩니다.",
         "#fefce8", "#d97706"),
        ("4. 공용 PC 보안 수칙 & 중복 신청 방지 유의",
         "• 회의실, 공용 작업장 PC 등 여러 사람이 함께 쓰는 PC에서는 본인 신청 후 반드시 브라우저 탭을 닫아주세요.\n"
         "• 다른 동료가 연이어 사용할 때는 사원번호 입력창에 본인 사번을 새로 덮어쓰고 입장하도록 안내합니다.\n"
         "• 동일한 날짜에는 시스템이 중복 신청을 사전에 원천 차단하므로, 날짜 변경 시에는 '내 신청 내역'에서 [수정] 버튼을 이용하세요.",
         "#fdf2f8", "#db2777")
    ]

    positions = [
        (Inches(0.8), Inches(1.65)),
        (Inches(6.8), Inches(1.65)),
        (Inches(0.8), Inches(4.45)),
        (Inches(6.8), Inches(4.45))
    ]

    for idx, ((r_title, r_desc, r_bg, r_line), (rx, ry)) in enumerate(zip(reentry_cards, positions)):
        r_box = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, rx, ry, Inches(5.7), Inches(2.6))
        r_box.fill.solid()
        r_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        r_box.line.color.rgb = RGBColor(203, 213, 225)
        r_box.line.width = Pt(1.5)

        r_tb = s5.shapes.add_textbox(rx + Inches(0.2), ry + Inches(0.15), Inches(5.3), Inches(2.3))
        rtf = r_tb.text_frame
        rtf.word_wrap = True

        rp_tag = rtf.paragraphs[0]
        rp_tag.text = f"유의사항 0{idx+1}  |  {r_title}"
        rp_tag.font.size = Pt(15)
        rp_tag.font.bold = True
        rp_tag.font.color.rgb = COLOR_PRIMARY
        rp_tag.space_after = Pt(8)

        rp_desc = rtf.add_paragraph()
        rp_desc.text = r_desc
        rp_desc.font.size = Pt(11.5)
        rp_desc.font.color.rgb = COLOR_TEXT_MAIN
        rp_desc.line_spacing = 1.3

    # --- SLIDE 6: 5단계 - 달력 콕 찍어서 특근 신청하기 ---
    s6_data = [
        ("① [일반휴일] 기본 선택 & 토요일 자동 세팅", "주말 특근에 가장 많이 쓰이는 '일반휴일'이 디폴트로 선택되며, 시작일/종료일이 이번 주 토요일로 자동 설정됩니다."),
        ("② 시작일/종료일 상호 자동 동기화", "시작일(또는 종료일)을 선택하면 다른 날짜도 동일하게 자동 변경되며, 이후 두 번째 날짜를 변경해 2단계로 자유롭게 범위를 설정합니다."),
        ("③ 프로젝트번호/장소/사유 입력 가이드", "프로젝트번호(예: BT2601-L1), 근무장소(예: 본사5층), 특근사유(예: 프로그램 개발) 예시 가이드가 기본 제공됩니다."),
        ("④ 달력 인터랙티브 날짜 선택 & 중복 방지", "오른쪽 달력에서 날짜를 터치하여 선택할 수 있으며, 동일한 날짜에 중복 신청하는 실수를 사전에 원천 차단합니다.")
    ]
    add_visual_slide(prs, "05", "스마트 달력 기반 특근 신청", "일반휴일 및 이번 주 토요일이 기본 선택되어 몇 초 만에 신청이 완료됩니다.", s6_data, images["apply"])

    # --- SLIDE 7: 6단계 - 나의 특근 신청 내역 및 다차원 검색 ---
    s7_data = [
        ("① 실시간 검색 결과 통계 요약 바", "검색 조건에 맞춰 총 건수(총 일수), 승인완료/대기 건수, 일반/법정/대체근무별 수량이 미니 통계 칩으로 실시간 집계 표시됩니다."),
        ("② 다차원 검색 필터", "조회기간, 특근구분, 승인상태 및 프로젝트/장소/사유 통합 키워드 검색으로 원하는 특근을 즉시 찾습니다."),
        ("③ 목록 보기 ↔ 개인 달력 보기 전환", "신청 내역을 카드 리스트 형태뿐만 아니라 개인 전용 월간 달력 형태로도 한눈에 확인할 수 있습니다."),
        ("④ 승인 상태 및 대체휴일 명확 표시 & 수정/취소", "관리자 승인 시 [✓ 확인완료] 표시, 대휴 일수 안내, 일정 변경 시 [수정], 취소 시 [삭제] 및 이력이 안전 보존됩니다.")
    ]
    add_visual_slide(prs, "06", "나의 특근 신청 내역 및 다차원 검색", "다양한 검색 조건과 개인 달력 뷰로 본인의 특근 일정을 편리하게 확인합니다.", s7_data, images["my"])

    # --- SLIDE 8: 7단계 - 4단계 특근 라이프사이클 & 특근 완료 후 [확정] 피드백 (v1.53 신규) ---
    s8 = prs.slides.add_slide(blank_layout)
    header_tb8 = s8.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf8 = header_tb8.text_frame
    htf8.word_wrap = True
    p8_step = htf8.paragraphs[0]
    p8_step.text = "STEP 07  |  특근 4단계 라이프사이클 & 사원 [확정] 피드백 (v1.53 최신)"
    p8_step.font.size = Pt(12)
    p8_step.font.bold = True
    p8_step.font.color.rgb = COLOR_ACCENT

    p8_title = htf8.add_paragraph()
    p8_title.text = "특근 4단계 라이프사이클 & 승인/확정 독립 운영"
    p8_title.font.size = Pt(22)
    p8_title.font.bold = True
    p8_title.font.color.rgb = COLOR_PRIMARY

    p8_sub = htf8.add_paragraph()
    p8_sub.text = "신청 후 승인이 없어도 사원이 즉시 확정 가능! 승인과 확정이 분리된 유연한 피드백 체계"
    p8_sub.font.size = Pt(13)
    p8_sub.font.color.rgb = COLOR_TEXT_MUTED

    lifecycle_cards = [
        ("1. 승인과 확정의 독립 분리 체계",
         "• 1단계 [신청]: 본인이 특근 일정을 등록한 대기 상태\n"
         "• [승인]과 [확정]은 별개로 독립 관리되어 사전 승인이 없어도 사원이 직접 완료 확정 가능!\n"
         "• 승인 상태(✅ 승인/⏳ 대기)와 확정 상태(🎯 확정/⚪ 미확정)가 개별 독립 배지로 동시 표시\n"
         "• 관리자가 최종 검토 마감하면 🟣 [검토완료] 상태로 안전하게 확정됩니다.",
         "#eff6ff", "#2563eb"),
        ("2. 사원의 [🎯 특근완료 확정] 피드백",
         "• 실제 특근을 마치면 '내 신청 내역'으로 이동합니다.\n"
         "• 관리자 사전 승인 유무와 상관없이 [🎯 특근완료 확정] 버튼을 클릭!\n"
         "• 확정 즉시 [🎯 특근확정] 배지가 부여되며, 잘못 누른 경우 [🎯 확정취소]로 되돌릴 수 있습니다.\n"
         "• 팝업 확인 시 오류 없이 한글 완료 메시지가 안내됩니다.",
         "#ecfdf5", "#10b981"),
        ("3. 진행상태 배지 & 최종특근일 반영",
         "• 확정 즉시 카드 상단에 [🎯 특근확정] 배지가 부여되어 진행 현황을 즉시 파악합니다.\n"
         "• 정산 엑셀 파일의 [확정단계 (일)]에 실시간 반영되어 급여/휴무 정산 누락을 방지합니다.\n"
         "• 관리자의 최종 검토가 완료되면 [🟣 검토완료]로 최종 종결됩니다.",
         "#fef3c7", "#ea580c")
    ]

    for i, (c_title, c_desc, c_bg, c_line) in enumerate(lifecycle_cards):
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
        p1.text = f"핵심 절차 0{i+1}"
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
        p3.font.size = Pt(12.5)
        p3.font.color.rgb = COLOR_TEXT_MAIN
        p3.line_spacing = 1.35

    # --- SLIDE 9: 8단계 - 사원 전용 100% 무기명 건의사항 소통함 (v1.53 신규) ---
    s9_sug = prs.slides.add_slide(blank_layout)
    header_tb9_sug = s9_sug.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf9_sug = header_tb9_sug.text_frame
    htf9_sug.word_wrap = True
    p9_step = htf9_sug.paragraphs[0]
    p9_step.text = "STEP 08  |  사원 전용 100% 무기명 건의사항 소통함 (v1.53)"
    p9_step.font.size = Pt(12)
    p9_step.font.bold = True
    p9_step.font.color.rgb = COLOR_ACCENT

    p9_title = htf9_sug.add_paragraph()
    p9_title.text = "100% 무기명 건의사항 소통함 및 공식 답변 피드백"
    p9_title.font.size = Pt(22)
    p9_title.font.bold = True
    p9_title.font.color.rgb = COLOR_PRIMARY

    p9_sub = htf9_sug.add_paragraph()
    p9_sub.text = "사번/성명 일체 미수집 원칙! 사원들의 솔직한 불편사항과 개선 요청을 자유롭게 나눕니다."
    p9_sub.font.size = Pt(13)
    p9_sub.font.color.rgb = COLOR_TEXT_MUTED

    suggestion_cards = [
        ("1. 100% 완전 무기명(익명) 보장",
         "• 시스템 DB에 사원번호, 성명, IP 등 작성자 식별 정보를 일체 저장하지 않습니다.\n"
         "• 누구의 눈치도 볼 필요 없이 시스템 오류, 근무환경 개선, 건의사항을 자유롭게 작성할 수 있습니다.\n"
         "• 오직 건의 내용과 접수 일시만 등록되므로 안심하고 이용하세요.",
         "#eff6ff", "#2563eb"),
        ("2. [💡 건의사항 소통함] 원클릭 등록",
         "• 화면 우측 상단의 [💡 건의사항 소통함] 버튼을 클릭합니다.\n"
         "• 분류(시스템 개선, 오류/버그 신고, 근무환경/복지, 기타)를 선택하고 제목과 내용을 작성합니다.\n"
         "• [💡 무기명 건의 접수하기] 버튼을 누르면 1초 만에 즉시 등록됩니다.",
         "#ecfdf5", "#10b981"),
        ("3. 실시간 접수 목록 & 관리자 답변",
         "• 등록된 모든 건의사항은 소통함 모달 하단 목록에서 실시간으로 투명하게 열람됩니다.\n"
         "• 관리자가 조치 상태(접수완료, 검토중, 처리완료, 보류)를 갱신하고 공식 답변을 등록합니다.\n"
         "• 사원들의 목소리가 실제 시스템 및 부서 운영 개선에 즉각 반영됩니다.",
         "#fef3c7", "#ea580c")
    ]

    for i, (c_title, c_desc, c_bg, c_line) in enumerate(suggestion_cards):
        cx = Inches(0.8 + i * 4.0)
        c_box = s9_sug.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), Inches(3.7), Inches(4.8))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        c_box.line.color.rgb = RGBColor(203, 213, 225)
        c_box.line.width = Pt(2)

        c_tb = s9_sug.shapes.add_textbox(cx + Inches(0.2), Inches(2.0), Inches(3.3), Inches(4.3))
        ctf = c_tb.text_frame
        ctf.word_wrap = True

        p1 = ctf.paragraphs[0]
        p1.text = f"소통 원칙 0{i+1}"
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
        p3.font.size = Pt(12.5)
        p3.font.color.rgb = COLOR_TEXT_MAIN
        p3.line_spacing = 1.35

    # --- SLIDE 10: 9단계 - 스마트폰 모바일 & 외부망 원격 접속 ---
    s10 = prs.slides.add_slide(blank_layout)
    header_tb10 = s10.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf10 = header_tb10.text_frame
    htf10.word_wrap = True
    p10_step = htf10.paragraphs[0]
    p10_step.text = "STEP 09  |  스마트폰 모바일 및 외부망 접속 심화"
    p10_step.font.size = Pt(12)
    p10_step.font.bold = True
    p10_step.font.color.rgb = COLOR_ACCENT

    p10_title = htf10.add_paragraph()
    p10_title.text = "모바일 및 외부망(LTE/5G/사외) 스마트 활용"
    p10_title.font.size = Pt(22)
    p10_title.font.bold = True
    p10_title.font.color.rgb = COLOR_PRIMARY

    p10_sub = htf10.add_paragraph()
    p10_sub.text = "PC 앞이 아니어도 스마트폰으로 특근을 신청하고 실시간 승인 현황을 확인할 수 있습니다."
    p10_sub.font.size = Pt(13)
    p10_sub.font.color.rgb = COLOR_TEXT_MUTED

    step_cards = [
        ("1. 외부 접속용 모바일 QR 확인", "화면 상단 [📱 모바일 접속 QR]을 누르면 스마트폰 카메라(LTE/5G)로 즉시 스캔 가능한 외부망 전용 QR 코드가 표시됩니다.", "#eff6ff", "#2563eb"),
        ("2. 암호화(HTTPS) 보안 터널", "외부 보안 터널(Cloudflare)을 통해 모든 통신이 전송 암호화(SSL/TLS)되어 사외에서도 안전하게 신청할 수 있습니다.", "#ecfdf5", "#10b981"),
        ("3. 세션 유지 & 홈 화면 추가", "스마트폰 브라우저에서 '홈 화면에 추가'를 누르면 전용 앱처럼 등록되며 24시간 언제 어디서든 즉시 접속됩니다.", "#fef3c7", "#ea580c")
    ]

    for i, (c_title, c_desc, c_bg, c_line) in enumerate(step_cards):
        cx = Inches(0.8 + i * 4.0)
        c_box = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), Inches(3.7), Inches(4.8))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        c_box.line.color.rgb = RGBColor(203, 213, 225)
        c_box.line.width = Pt(2)

        c_tb = s10.shapes.add_textbox(cx + Inches(0.2), Inches(2.0), Inches(3.3), Inches(4.3))
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

    # --- SLIDE 11: 사용자용 Q&A 총정리 (8문항 확장) ---
    s11 = prs.slides.add_slide(blank_layout)
    header_tb11 = s11.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.7), Inches(1.0))
    htf11 = header_tb11.text_frame
    htf11.word_wrap = True
    p_step11 = htf11.paragraphs[0]
    p_step11.text = "SUMMARY & Q&A  |  사용자 자주 묻는 질문 답변 (v1.53)"
    p_step11.font.size = Pt(12)
    p_step11.font.bold = True
    p_step11.font.color.rgb = COLOR_ACCENT

    p_title11 = htf11.add_paragraph()
    p_title11.text = "자주 묻는 질문(Q&A)과 사용자 꿀팁 총정리"
    p_title11.font.size = Pt(22)
    p_title11.font.bold = True
    p_title11.font.color.rgb = COLOR_PRIMARY

    p_sub11 = htf11.add_paragraph()
    p_sub11.text = "궁금한 사항이 있으실 때는 언제든 상단의 [📖 사용자 매뉴얼]을 다운로드하여 확인하세요."
    p_sub11.font.size = Pt(12.5)
    p_sub11.font.color.rgb = COLOR_TEXT_MUTED

    user_qa_items = [
        ("Q1. 외부 접속망 주소와 모바일 접속 방법은 어떻게 되나요?", "관리자가 배포한 외부 보안 주소(https://xxxx.trycloudflare.com)로 접속하거나, 화면 상단 [모바일 접속 QR]을 스마트폰 카메라로 스캔하여 LTE/5G 어디서나 간편하게 접속합니다."),
        ("Q2. 서버 리부팅 직후 접속이 왜 바로 안 되고 로딩이 도나요?", "서버 최초 기동 및 리부팅 시에는 데이터베이스 무결성 검증과 자동 백업 생성을 위해 약 10~20초 초기화가 진행되니 새로고침 연타 없이 잠시 대기해주세요."),
        ("Q3. 재입장할 때 사번을 매번 다시 입력해야 하나요?", "아닙니다. 브라우저가 사번을 안전하게 기억하고 있으므로 [입장하기] 버튼만 누르면 1초 만에 바로 입장됩니다."),
        ("Q4. 스마트폰 모바일에서 신청해도 회사 PC와 연동되나요?", "네, 모바일과 PC는 동일한 중앙 데이터베이스를 공유하므로 사번만 넣으면 100% 실시간 자동 연동됩니다."),
        ("Q5. 실수로 같은 날짜에 중복 신청하면 어떻게 되나요?", "시스템에서 동일 날짜 중복 신청을 사전에 감지하여 경고창과 함께 안전하게 차단하므로 안심하셔도 됩니다."),
        ("Q6. 관리자 승인 여부를 실시간으로 어떻게 확인하나요?", "'내 신청 내역' 카드에서 초록색 [✓ 확인완료] 도장을 확인하실 수 있으며, 상단 [새로고침]을 누르면 실시간 반영됩니다."),
        ("Q7. 실제 특근 후 [🎯 특근완료 확정] 버튼은 언제 누르나요?", "관리자가 사전 승인한 특근에 대해 실제 휴일 근무를 마친 후 '내 신청 내역'에서 [🎯 특근완료 확정]을 누르면 관리자에게 완료 피드백이 전송됩니다."),
        ("Q8. 건의사항 소통함은 정말로 작성자가 누구인지 알 수 없나요?", "네, 100% 무기명입니다! 데이터베이스에 사번, 성명, IP 등 사용자 식별 컬럼이 아예 존재하지 않도록 설계되어 절대 추적되지 않습니다.")
    ]

    for i, (q, a) in enumerate(user_qa_items):
        qy = Inches(1.50 + i * 0.70)
        q_box = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), qy, Inches(11.733), Inches(0.66))
        q_box.fill.solid()
        q_box.fill.fore_color.rgb = RGBColor(248, 250, 252)
        q_box.line.color.rgb = RGBColor(226, 232, 240)

        q_tb = s11.shapes.add_textbox(Inches(1.0), qy + Inches(0.04), Inches(11.3), Inches(0.58))
        qtf = q_tb.text_frame
        qtf.word_wrap = True

        qp1 = qtf.paragraphs[0]
        qp1.text = q
        qp1.font.size = Pt(11.5)
        qp1.font.bold = True
        qp1.font.color.rgb = COLOR_PRIMARY
        qp1.space_after = Pt(1)

        qp2 = qtf.add_paragraph()
        qp2.text = f"👉 {a}"
        qp2.font.size = Pt(10)
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
    p1_tag.text = "SMART OVERTIME SYSTEM v1.53  |  관리자 및 운영자 전용 가이드"
    p1_tag.font.size = Pt(14)
    p1_tag.font.bold = True
    p1_tag.font.color.rgb = RGBColor(253, 186, 116)
    p1_tag.space_after = Pt(14)

    p1_title = tf1.add_paragraph()
    p1_title.text = "스마트 특근 관리 시스템\n관리자 모드 운영 매뉴얼 (v1.53)"
    p1_title.font.size = Pt(36)
    p1_title.font.bold = True
    p1_title.font.color.rgb = RGBColor(255, 255, 255)
    p1_title.space_after = Pt(20)

    p1_sub = tf1.add_paragraph()
    p1_sub.text = "승인/확정/검토 독립 제어 및 수정 모달 관리, 4단계 일수 구분 집계, 100% 무기명 소통함 운영까지 완벽 지원하는 관리자 공식 가이드입니다."
    p1_sub.font.size = Pt(15)
    p1_sub.font.color.rgb = RGBColor(203, 213, 225)
    p1_sub.space_after = Pt(28)

    p1_auth = tf1.add_paragraph()
    p1_auth.text = "배포 버전: v1.53 (2026-09-21)  |  총괄 슈퍼관리자 및 부서 팀관리자 전용"
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

    # --- SLIDE 6: 실특근 자동 산정 및 정산표 (v1.53 개편) ---
    s6_data = [
        ("① 최신 실특근 산정 공식 (v1.53)", "★ 총특근일수 = 대체근무 + 법정휴일 + 일반휴일 (대체휴무 제외 표준화)\n★ 최종 실특근일 = 일반특근 - 사전차감 - (대체휴무 - 대체휴무시 출장기간 내 사전차감)"),
        ("② 사전차감 잔여수 & 보너스 건수", "★ 사전차감 잔여수 = 총 사전차감수 - 출장기간 내 사전차감수 / 사원별 [보너스 부여 (건)] 개수 독립 산출"),
        ("③ 개인별 정산 및 통계 팝업 개편", "사원별 최근 3개월(7/8/9월), 1~4분기, 상/하반기, 연도별 집계와 사전차감 잔여수 독립 제공"),
        ("④ 정산표 엑셀 원클릭 추출 & 수식 명시", "엑셀 중복 '대체휴가 사용일수' 제거, 헤더에 산출 수식 및 취합 기간/일시 자동 명기")
    ]
    add_visual_slide(prs, "05", "[관리자 모드] 실특근 자동 산정 및 정산표", "대체휴일 사용분을 자동 차감하여 정확한 최종 실특근일을 산출합니다.", s6_data, images["settlement"])

    # --- SLIDE 7: 관리자 보너스 수정 및 일괄 삭제 기능 ---
    s7_data = [
        ("① 관리자 모드 [보너스 부여 여부] 수정", "관리자 특근 수정 모달에서 [🎁 보너스 부여 여부] 토글 스위치를 통해 보너스 부여 상태를 직접 ON/OFF 변경 가능"),
        ("② 관리자 테이블 [선택 일괄 삭제] 신설", "테이블 목록에서 체크박스로 여러 특근 건을 선택 후 [🗑️ 선택 일괄 삭제] 클릭 시 안전 컨펌 후 일괄 삭제"),
        ("③ 사용자 / 관리자 수정 모달 완전 분리", "일반 사원 수정 창에는 사전차감/보너스 항목이 원천 미노출되며, 관리자 창에서만 전담 제어"),
        ("④ openpyxl 고속 서버 엑셀 엔진 동기화", "엑셀 Sheet 2에 [총 특근일수], [보너스 부여 (건)] 반영 및 클라이언트 SheetJS와 100% 동일 포맷 유지")
    ]
    add_visual_slide(prs, "06", "보너스 수정 · 일괄 삭제 및 엑셀 정산", "관리자 보너스 관리, 일괄 삭제 및 엑셀 정산표를 완벽히 지원합니다.", s7_data, images["excel"])

    # --- SLIDE 8: 특근 4단계 검증 체계 및 관리자 [검토완료] (v1.53 신규) ---
    s8_admin = prs.slides.add_slide(blank_layout)
    header_tb8_adm = s8_admin.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf8_adm = header_tb8_adm.text_frame
    htf8_adm.word_wrap = True
    p8_step_adm = htf8_adm.paragraphs[0]
    p8_step_adm.text = "STEP 07  |  특근 확정 · 검토 종합 관리 및 개별 제어 (v1.53 최신)"
    p8_step_adm.font.size = Pt(12)
    p8_step_adm.font.bold = True
    p8_step_adm.font.color.rgb = COLOR_ACCENT

    p8_title_adm = htf8_adm.add_paragraph()
    p8_title_adm.text = "승인/확정/검토 독립 제어 및 수정 모달 종합 관리"
    p8_title_adm.font.size = Pt(22)
    p8_title_adm.font.bold = True
    p8_title_adm.font.color.rgb = COLOR_PRIMARY

    p8_sub_adm = htf8_adm.add_paragraph()
    p8_sub_adm.text = "테이블/캘린더 원클릭 버튼 세트, 수정 모달 3대 상태 토글 및 일괄 처리 강력 지원"
    p8_sub_adm.font.size = Pt(13)
    p8_sub_adm.font.color.rgb = COLOR_TEXT_MUTED

    admin_stage_cards = [
        ("1. 원클릭 승인/확정/검토 개별 제어",
         "• 관리자 테이블 및 캘린더 작업자 패널에 각 단계별 원클릭 버튼 세트 완비\n"
         "• 승인: [✓ 승인] ↔ [✕ 취소] 자유 전환\n"
         "• 확정: [🎯 확정] ↔ [✕ 확정취소] 독립 제어\n"
         "• 검토: [🟣 검토완료] ↔ [✕ 검토취소] 최종 마감\n"
         "• 대체휴무(휴가): 4단계 절차가 자동 생략되어 등록 즉시 [🌿 대체휴무]로 자동 반영됩니다.",
         "#eff6ff", "#2563eb"),
        ("2. 관리자 특근 수정 모달 상태 직접 설정",
         "• 테이블에서 [수정] 버튼 클릭 시 나타나는 [관리자 특근 수정 모달]에 '진행 상태 관리' 신설\n"
         "• 1) 관리자 승인 (⏳ 대기 ↔ ✅ 승인)\n"
         "• 2) 특근완료 확정 (⚪ 미확정 ↔ 🎯 확정)\n"
         "• 3) 관리자 최종 검토 (⚪ 미검토 ↔ 🟣 검토)\n"
         "• 관리자가 세 플래그를 자유롭게 지정하여 안전 저장 가능!",
         "#ecfdf5", "#10b981"),
        ("3. 상단 [일괄 확정] & [일괄 검토완료]",
         "• 관리자 테이블 목록에서 체크박스로 여러 특근 건을 복수 선택합니다.\n"
         "• 상단의 [🎯 일괄 확정] 또는 [🟣 일괄 검토완료] 버튼을 눌러 대량의 특근을 한 번에 처리!\n"
         "• 일괄 확정 및 일괄 검토완료 시 진행 카운트 및 감사 로그(Audit Log)가 자동 기록됩니다.",
         "#fef3c7", "#ea580c")
    ]

    for i, (c_title, c_desc, c_bg, c_line) in enumerate(admin_stage_cards):
        cx = Inches(0.8 + i * 4.0)
        c_box = s8_admin.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), Inches(3.7), Inches(4.8))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        c_box.line.color.rgb = RGBColor(203, 213, 225)
        c_box.line.width = Pt(2)

        c_tb = s8_admin.shapes.add_textbox(cx + Inches(0.2), Inches(2.0), Inches(3.3), Inches(4.3))
        ctf = c_tb.text_frame
        ctf.word_wrap = True

        p1 = ctf.paragraphs[0]
        p1.text = f"검증 원칙 0{i+1}"
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
        p3.font.size = Pt(12.5)
        p3.font.color.rgb = COLOR_TEXT_MAIN
        p3.line_spacing = 1.35

    # --- SLIDE 9: 특근정보 엑셀 내보내기 4단계 구분 집계 (v1.53 신규) ---
    s9_admin_excel = prs.slides.add_slide(blank_layout)
    header_tb9_excel = s9_admin_excel.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf9_excel = header_tb9_excel.text_frame
    htf9_excel.word_wrap = True
    p9_step_excel = htf9_excel.paragraphs[0]
    p9_step_excel.text = "STEP 08  |  특근정보 엑셀 내보내기 4단계 구분 집계 (v1.53)"
    p9_step_excel.font.size = Pt(12)
    p9_step_excel.font.bold = True
    p9_step_excel.font.color.rgb = COLOR_ACCENT

    p9_title_excel = htf9_excel.add_paragraph()
    p9_title_excel.text = "엑셀 보고서 4단계(신청·승인·확정·검토) 일수 완벽 분리 집계"
    p9_title_excel.font.size = Pt(22)
    p9_title_excel.font.bold = True
    p9_title_excel.font.color.rgb = COLOR_PRIMARY

    p9_sub_excel = htf9_excel.add_paragraph()
    p9_sub_excel.text = "신청, 승인, 확정, 검토완료 4단계에 따라 최종특근일을 명확히 구분 산출합니다."
    p9_sub_excel.font.size = Pt(13)
    p9_sub_excel.font.color.rgb = COLOR_TEXT_MUTED

    excel_stage_cards = [
        ("1. Sheet 1 [휴일일자별_특근현황]",
         "• 4단계 진행상태 열(신청:호박, 승인:청색, 확정:초록, 검토완료:보라) 컬러 스타일 적용\n"
         "• 승인자, 확정자, 검토완료자 일시 정보가 정확하게 기록됩니다.\n"
         "• 특정 일자에 누가 확정하고 검토했는지 한눈에 추적 가능합니다.",
         "#eff6ff", "#2563eb"),
        ("2. Sheet 2 [개인별_휴일합산_정산표]",
         "• [신청/승인/확정/검토완료] 4개 단계별 '최종 실특근일+보너스 [일]'이 각각 별도 계산!\n"
         "• 신청단계 헤더를 짙은 호박색(Amber)으로 변경하여 글자 시인성을 극대화(고대비 보장)!\n"
         "• 각 단계별 합계(SUM 수식)가 하단에 자동 계산되어 원하는 단계의 실특근일만 즉시 확인 가능합니다.",
         "#ecfdf5", "#10b981"),
        ("3. Sheet 4 [부서별_요약] 및 정산표",
         "• 부서별 요약 시트에도 4단계별 최종 실특근일+보너스 합계가 자동 집계됩니다.\n"
         "• 관리자 정산표 모달 및 정산 엑셀 파일에도 동일한 4단계 분리 및 고대비 서식이 완벽 적용됩니다.\n"
         "• 정산의 신뢰성과 데이터 정밀도가 비약적으로 향상되었습니다.",
         "#fef3c7", "#ea580c")
    ]

    for i, (c_title, c_desc, c_bg, c_line) in enumerate(excel_stage_cards):
        cx = Inches(0.8 + i * 4.0)
        c_box = s9_admin_excel.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), Inches(3.7), Inches(4.8))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        c_box.line.color.rgb = RGBColor(203, 213, 225)
        c_box.line.width = Pt(2)

        c_tb = s9_admin_excel.shapes.add_textbox(cx + Inches(0.2), Inches(2.0), Inches(3.3), Inches(4.3))
        ctf = c_tb.text_frame
        ctf.word_wrap = True

        p1 = ctf.paragraphs[0]
        p1.text = f"집계 규격 0{i+1}"
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
        p3.font.size = Pt(12.5)
        p3.font.color.rgb = COLOR_TEXT_MAIN
        p3.line_spacing = 1.35

    # --- SLIDE 10: 사원 무기명 건의사항 소통함 운영 및 공식 답변 (v1.53 신규) ---
    s10_admin_sug = prs.slides.add_slide(blank_layout)
    header_tb10_sug = s10_admin_sug.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf10_sug = header_tb10_sug.text_frame
    htf10_sug.word_wrap = True
    p10_step_sug = htf10_sug.paragraphs[0]
    p10_step_sug.text = "STEP 09  |  사원 무기명 건의사항 소통함 운영 & 답변 (v1.53)"
    p10_step_sug.font.size = Pt(12)
    p10_step_sug.font.bold = True
    p10_step_sug.font.color.rgb = COLOR_ACCENT

    p10_title_sug = htf10_sug.add_paragraph()
    p10_title_sug.text = "사원 무기명 건의사항 소통함 운영 및 공식 피드백 관리"
    p10_title_sug.font.size = Pt(22)
    p10_title_sug.font.bold = True
    p10_title_sug.font.color.rgb = COLOR_PRIMARY

    p10_sub_sug = htf10_sug.add_paragraph()
    p10_sub_sug.text = "100% 무기명으로 접수된 사원 건의사항을 검토하고 조치 상태와 공식 답변을 피드백합니다."
    p10_sub_sug.font.size = Pt(13)
    p10_sub_sug.font.color.rgb = COLOR_TEXT_MUTED

    admin_sug_cards = [
        ("1. 실시간 무기명 접수 모니터링",
         "• 화면 우측 상단 [💡 건의사항 소통함]을 클릭해 사원들이 남긴 건의사항을 확인합니다.\n"
         "• 사번/성명 없는 100% 무기명 원칙으로 사원들이 진솔한 의견을 개진할 수 있습니다.\n"
         "• 분류(시스템개선, 버그신고, 근무환경 등)별로 정리되어 현안 파악이 용이합니다.",
         "#eff6ff", "#2563eb"),
        ("2. 조치 상태 변경 & 공식 답변 등록",
         "• 관리자 로그인 상태에서는 각 건의 카드 하단에 [답변 작성] 버튼이 활성화됩니다.\n"
         "• 조치 상태(접수완료, 검토중, 처리완료, 보류)를 선택하고 공식 답변 내용을 입력 후 등록합니다.\n"
         "• 등록된 답변은 즉시 사원들에게 공유되어 조직 소통 만족도를 극대화합니다.",
         "#ecfdf5", "#10b981"),
        ("3. 건전한 게시판 유지 관리",
         "• 중복 등록되거나 부적절한 게시글은 관리자 전용 [삭제] 버튼으로 안전하게 정비 가능합니다.\n"
         "• 사원들의 건의를 바탕으로 시스템 기능 개선 로드맵을 수립할 수 있습니다.",
         "#fef3c7", "#ea580c")
    ]

    for i, (c_title, c_desc, c_bg, c_line) in enumerate(admin_sug_cards):
        cx = Inches(0.8 + i * 4.0)
        c_box = s10_admin_sug.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), Inches(3.7), Inches(4.8))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        c_box.line.color.rgb = RGBColor(203, 213, 225)
        c_box.line.width = Pt(2)

        c_tb = s10_admin_sug.shapes.add_textbox(cx + Inches(0.2), Inches(2.0), Inches(3.3), Inches(4.3))
        ctf = c_tb.text_frame
        ctf.word_wrap = True

        p1 = ctf.paragraphs[0]
        p1.text = f"운영 원칙 0{i+1}"
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
        p3.font.size = Pt(12.5)
        p3.font.color.rgb = COLOR_TEXT_MAIN
        p3.line_spacing = 1.35

    # --- SLIDE 11: 관리자 전용 웹 스냅샷 저장 및 복원 ---
    s8 = prs.slides.add_slide(blank_layout)
    header_tb8 = s8.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    htf8 = header_tb8.text_frame
    htf8.word_wrap = True
    p_step8 = htf8.paragraphs[0]
    p_step8.text = "STEP 10  |  관리자 전용 웹 데이터 스냅샷 백업 & 복원"
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
    p_step9_sec.text = "STEP 11  |  슈퍼관리자 전용 시스템 보안 및 접속 감사 로그"
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
    p_step9.text = "STEP 12  |  스마트폰 모바일 및 외부망 접속 관리"
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
        ("1. 외부망 기본 접속 QR 배포", "상단 [📱 모바일 접속 QR]을 띄워 직원들이 스마트폰(LTE/5G)으로 외부망에 즉시 접속할 수 있는 기본 QR 코드를 공유합니다.", "#eff6ff", "#2563eb"),
        ("2. 외부 도메인 변경 시 원클릭 갱신", "Cloudflare 등 외부 접속 도메인이나 전용 주소가 변경될 경우 팝업에서 새 주소를 입력하고 [QR 갱신]을 눌러 즉시 반영합니다.", "#ecfdf5", "#10b981"),
        ("3. 24시간 안전 보안 터널 구동", "Cloudflare Tunnel(start_external_access.bat)을 통해 포트포워딩 없이도 외부 접속이 전송 구간 암호화(HTTPS)로 안전 유지됩니다.", "#fef3c7", "#ea580c")
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
    p_title10.text = "자주 묻는 질문(Q&A)과 관리자 운영 꿀팁 총정리 (v1.53)"
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
        ("Q3. 팀관리자와 총괄관리자의 권한 차이는 어떻게 되나요?", "팀관리자는 본인 팀원의 승인/수정만 가능하며, 팀원 소속팀 변경, 슈퍼관리자 승격/하야 지정은 오직 총괄관리자만 가능합니다."),
        ("Q4. 승인 완료된 특근을 사원이 임의로 수정하거나 삭제할 수 있나요?", "아닙니다. v1.42부터 승인된 특근은 일반 사원의 수정/삭제가 원천 차단되며, 오직 관리자 모드에서 관리자만 수정 또는 삭제할 수 있습니다."),
        ("Q5. 서버가 재부팅되어도 인원과 특근 데이터가 영구 보존되나요?", "네! Turso 클라우드 영구 DB 연동으로 모든 데이터가 안전 보존되며, v1.53의 안전 스마트 폴백과 스레드 보호로 특근 신청과 데이터 무결성을 완벽하게 보장합니다."),
        ("Q6. 특근 승인과 확정은 어떻게 운영되나요?", "v1.53부터 특근 승인과 확정은 완전히 독립 분리되었습니다. 관리자 사전 승인이 없어도 사원이 특근 완료 후 즉시 [특근완료 확정]을 할 수 있으며, 관리자는 테이블/캘린더/수정모달에서 승인, 확정, 검토완료를 각각 개별 제어 및 취소할 수 있습니다."),
        ("Q7. 엑셀 내보내기에서 4단계 특근일수는 어떻게 구분되어 나오나요?", "Sheet 1에는 4단계 진행상태와 승인/확정/검토자가 명시되며, Sheet 2 개인별 정산표 및 Sheet 4 부서별 요약에 신청·승인·확정·검토완료 일수가 각각 독립 열로 자동 집계됩니다."),
        ("Q8. 무기명 건의사항 소통함에서 관리자는 어떤 관리를 하나요?", "사원들이 무기명으로 접수한 모든 건의사항에 대해 조치 상태(접수완료/검토중/처리완료/보류)를 갱신하고 공식 답변을 등록하거나 부적절한 게시물을 삭제 관리합니다.")
    ]

    for i, (q, a) in enumerate(admin_qa_items):
        qy = Inches(1.50 + i * 0.70)
        q_box = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), qy, Inches(11.733), Inches(0.66))
        q_box.fill.solid()
        q_box.fill.fore_color.rgb = RGBColor(248, 250, 252)
        q_box.line.color.rgb = RGBColor(226, 232, 240)

        q_tb = s10.shapes.add_textbox(Inches(1.0), qy + Inches(0.04), Inches(11.3), Inches(0.58))
        qtf = q_tb.text_frame
        qtf.word_wrap = True

        qp1 = qtf.paragraphs[0]
        qp1.text = q
        qp1.font.size = Pt(11.5)
        qp1.font.bold = True
        qp1.font.color.rgb = COLOR_PRIMARY
        qp1.space_after = Pt(1)

        qp2 = qtf.add_paragraph()
        qp2.text = f"👉 {a}"
        qp2.font.size = Pt(10)
        qp2.font.color.rgb = COLOR_ACCENT

    return prs


# ----------------- 실행 및 저장 통합 함수 -----------------

def create_manual():
    print("[1/3] Generating ultra-friendly visual UI mockups with speech bubbles & leader lines...")
    images = {
        "init": create_mockup_server_init(),
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

    def copy_and_verify(src, dst):
        try:
            shutil.copy2(str(src), str(dst))
            if dst.exists() and dst.stat().st_size > 0:
                print(f"✓ Copied & Verified: {dst.name} ({dst.stat().st_size:,} bytes) -> {dst}")
                return True
            else:
                raise RuntimeError(f"Failed to verify copied file: {dst}")
        except PermissionError:
            print(f"⚠️ [NOTICE] Cannot overwrite '{dst.name}' because it is currently open. Retrying or saving updated version in downloads/.")
            return False
        except Exception as e:
            print(f"⚠️ [WARNING] Failed to copy '{src.name}' to '{dst.name}': {e}")
            return False

    # 4. 최상위 메인 폴더(루트) 및 static/downloads 디렉토리 동기화 복사
    root_dir = Path(__file__).resolve().parent
    static_dl = root_dir / "static" / "downloads"
    static_dl.mkdir(parents=True, exist_ok=True)
    
    # 루트 메인 폴더로 복사 (GitHub 업로드 및 메인 루트 동기화)
    copy_and_verify(USER_PPTX_PATH, root_dir / "Overtime_User_Manual.pptx")
    copy_and_verify(ADMIN_PPTX_PATH, root_dir / "Overtime_Admin_Manual.pptx")
    copy_and_verify(PPTX_PATH, root_dir / "Overtime_System_Manual.pptx")
    print(f"✓ Synchronized PPTX manuals with ROOT main directory: {root_dir}")

    # static/downloads로 복사
    copy_and_verify(USER_PPTX_PATH, static_dl / "Overtime_User_Manual.pptx")
    copy_and_verify(ADMIN_PPTX_PATH, static_dl / "Overtime_Admin_Manual.pptx")
    copy_and_verify(PPTX_PATH, static_dl / "Overtime_System_Manual.pptx")
    print(f"✓ Synchronized PPTX manuals with static download path: {static_dl}")
    print("🎉 Both User and Admin PPT Manuals generated and synchronized successfully!")

if __name__ == "__main__":
    create_manual()

