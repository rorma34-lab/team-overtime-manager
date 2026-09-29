import os
import sys
import shutil
from pathlib import Path

# 콘솔 UTF-8 인코딩 설정 (Windows cp949 환경 대응)
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

def get_font(size, bold=False):
    font_name = "malgunbd.ttf" if bold else "malgun.ttf"
    try:
        return ImageFont.truetype(font_name, size)
    except Exception:
        try:
            return ImageFont.truetype("arial.ttf", size)
        except Exception:
            return ImageFont.load_default()

CANVAS_W = 1600
CANVAS_H = 1100

COLOR_PRIMARY = RGBColor(30, 58, 138)     # Deep Royal Navy (#1E3A8A)
COLOR_ACCENT = RGBColor(234, 88, 12)     # Vivid Orange (#EA580C)
COLOR_SUCCESS = RGBColor(16, 185, 129)   # Emerald (#10B981)
COLOR_TEXT_MAIN = RGBColor(15, 23, 42)   # Dark Slate (#0F172A)
COLOR_TEXT_MUTED = RGBColor(100, 116, 139) # Slate Grey (#64748B)
COLOR_BG_CARD = RGBColor(248, 250, 252)  # Light Soft Grey (#F8FAFC)
COLOR_BORDER = RGBColor(226, 232, 240)   # Light Border (#E2E8F0)

# ----------------- UI 렌더링 헬퍼 함수 -----------------

def draw_browser_frame(d, title="스마트 특근 관리 시스템 (v1.54)", url="https://overtime-system.internal"):
    """상단 브라우저 프레임, 주소창, 최신 릴리즈 배지 렌더링"""
    d.rounded_rectangle([20, 15, 1580, 85], radius=10, fill="#1e293b")
    # 신호등 버튼
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
    """지시선 및 고시인성 번호 배지 렌더링"""
    tx, ty = target_xy
    px, py = pin_xy

    font = get_font(21, bold=True)
    badge_str = f" {badge_num}  {label_text} "
    try:
        bbox = font.getbbox(badge_str)
        text_w = (bbox[2] - bbox[0]) + 32
    except Exception:
        text_w = len(badge_str) * 22 + 32
    text_h = 44

    bx1 = px - text_w // 2
    bx1 = max(15, min(bx1, CANVAS_W - text_w - 15))
    by1 = py - text_h // 2
    by1 = max(15, min(by1, CANVAS_H - text_h - 50))
    bx2 = bx1 + text_w
    by2 = by1 + text_h

    if tx < bx1:
        dock_x = bx1
        dock_y = max(by1 + 6, min(by2 - 6, ty))
    elif tx > bx2:
        dock_x = bx2
        dock_y = max(by1 + 6, min(by2 - 6, ty))
    else:
        dock_x = tx
        dock_y = by1 if ty < by1 else by2

    d.line([(tx, ty), (dock_x, dock_y)], fill=color, width=3)
    d.ellipse([tx - 9, ty - 9, tx + 9, ty + 9], fill="#ffffff", outline=color, width=3)
    d.ellipse([tx - 4, ty - 4, tx + 4, ty + 4], fill=color)

    d.rounded_rectangle([bx1 + 3, by1 + 4, bx2 + 3, by2 + 4], radius=10, fill="#94a3b8")
    d.rounded_rectangle([bx1, by1, bx2, by2], radius=10, fill=color, outline="#ffffff", width=2)
    d.text((bx1 + text_w // 2, by1 + text_h // 2), badge_str, font=font, fill=text_color, anchor="mm")

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


# ----------------- UI 목업 삽화 이미지 생성기들 (1600x1100 고해상도) -----------------

# 1. 초기 접속 주소 및 웹 단일 접속 안내 목업 (PC 및 모바일)
def create_mockup_access_methods():
    """삽화: 공식 웹주소 단일화 및 PC / 모바일 접속 가이드 시각화"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  공식 웹 접속 안내", url="https://team-overtime-manager.onrender.com/")

    # 상단 안내 및 공식 웹주소 하이라이트 박스
    d.rounded_rectangle([40, 100, 1560, 240], radius=16, fill="#ffffff", outline="#2563eb", width=3)
    d.text((70, 120), "🌐 스마트 특근 관리 시스템 공식 단일 웹 접속 주소", font=get_font(26, bold=True), fill="#0f172a")
    d.text((70, 155), "별도 프로그램이나 VPN 설치 없이, PC 브라우저 및 스마트폰 모바일에서 동일한 공식 주소로 즉시 접속하세요.", font=get_font(16), fill="#64748b")

    # 웹주소 터치/클릭 배너
    d.rounded_rectangle([70, 182, 1150, 226], radius=8, fill="#0f172a")
    d.text((95, 204), "🔗 공식 웹주소:  https://team-overtime-manager.onrender.com/", font=get_font(18, bold=True), fill="#38bdf8", anchor="lm")
    d.rounded_rectangle([1170, 182, 1530, 226], radius=8, fill="#10b981")
    d.text((1350, 204), "🚀 클릭 시 즉시 연결 (PC / 모바일 공용)", font=get_font(16, bold=True), fill="#ffffff", anchor="mm")

    cards = [
        ("🖥️ 사내 / 자택 PC 브라우저 접속 안내",
         "#eff6ff", "#3b82f6",
         "사내 업무 PC, 자택 PC, 외부 노트북 (Chrome / Edge / Whale)",
         "https://team-overtime-manager.onrender.com/",
         ["① PC 웹 브라우저(Chrome, Edge 등)를 실행합니다.",
          "② 주소창에 공식 웹주소를 입력하거나 즐겨찾기를 클릭합니다.",
          "③ 브라우저 [즐겨찾기 ★] 등록 시 매번 원클릭으로 바로 진입!",
          "④ 회사 사내망은 물론 자택/출장지 어디서나 자유롭게 연결",
          "※ 별도 프로그램이나 플러그인 설치 없이 1초 만에 로딩 완료"],
         "① PC는 브라우저 주소 입력 및 즐겨찾기 등록"),

        ("📱 스마트폰 모바일 접속 안내 (Android / iPhone)",
         "#ecfdf5", "#10b981",
         "갤럭시, 아이폰, 아이패드, 태블릿 (모바일 반응형 자동 최적화)",
         "https://team-overtime-manager.onrender.com/",
         ["① 스마트폰 기본 브라우저(Safari, Chrome, 삼성인터넷) 실행",
          "② 주소창에 동일한 공식 웹주소 입력 (또는 전달받은 링크 터치)",
          "③ 모바일 터치에 최적화된 깔끔한 화면이 자동으로 펼쳐집니다.",
          "★ 브라우저 메뉴 [홈 화면에 추가] 터치 시 전용 앱 아이콘 생성!",
          "※ 언제 어디서나 출퇴근길/휴일에도 간편하게 특근 신청 및 확인"],
         "② 모바일은 스마트폰 브라우저 접속 후 홈 화면 추가")
    ]

    card_w = 735
    card_gap = 30
    start_x = 45

    for i, (title, bg_c, border_c, sub_head, url_sample, steps, pin_label) in enumerate(cards):
        cx = start_x + i * (card_w + card_gap)
        cy = 260
        ch = 700

        d.rounded_rectangle([cx, cy, cx + card_w, cy + ch], radius=16, fill="#ffffff", outline=border_c, width=3)
        d.rounded_rectangle([cx, cy, cx + card_w, cy + 85], radius=14, fill=bg_c)
        d.text((cx + card_w // 2, cy + 42), title, font=get_font(22, bold=True), fill="#0f172a", anchor="mm")

        d.rounded_rectangle([cx + 25, cy + 100, cx + card_w - 25, cy + 195], radius=10, fill="#f8fafc", outline="#cbd5e1", width=2)
        d.text((cx + 40, cy + 118), sub_head, font=get_font(15, bold=True), fill="#64748b")
        d.rounded_rectangle([cx + 35, cy + 142, cx + card_w - 35, cy + 185], radius=6, fill="#0f172a")
        d.text((cx + card_w // 2, cy + 163), url_sample, font=get_font(17, bold=True), fill="#38bdf8", anchor="mm")

        for s_idx, step_text in enumerate(steps):
            sy = cy + 210 + s_idx * 90
            d.rounded_rectangle([cx + 25, sy, cx + card_w - 25, sy + 75], radius=8, fill="#f8fafc", outline="#e2e8f0")
            if "★" in step_text:
                d.text((cx + 40, sy + 37), step_text, font=get_font(16, bold=True), fill="#b45309", anchor="lm")
            elif "※" in step_text:
                d.text((cx + 40, sy + 37), step_text, font=get_font(15), fill="#64748b", anchor="lm")
            else:
                d.text((cx + 40, sy + 37), step_text, font=get_font(16, bold=True), fill="#1e293b", anchor="lm")

        draw_smart_pin(d, (cx + card_w // 2, cy + 163), (cx + card_w // 2, cy + 630), f"0{i+1}", pin_label, color=border_c)

    d.rounded_rectangle([40, 980, 1560, 1070], radius=12, fill="#eff6ff", outline="#93c5fd", width=2)
    d.text((70, 1008), "★ 핵심 알림: PC와 스마트폰은 동일한 클라우드 DB로 100% 실시간 연동됩니다.", font=get_font(20, bold=True), fill="#1d4ed8")
    d.text((70, 1038), "어디서 접속하든 본인 사원번호 6자리만 입력하면 모든 신청 내역 및 승인 상태가 실시간으로 일치합니다.", font=get_font(16), fill="#2563eb")

    path = IMG_DIR / "mockup_access_methods.png"
    img.save(path)
    return str(path)


# 2. 초간편 사원번호 로그인 목업
def create_mockup_login():
    """삽화: 사원번호 6자리 간편 입장 및 자동 기억 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  사원번호 간편 입장", url="https://overtime.company.com/login")

    # 좌측: 메인 로그인 카드
    d.rounded_rectangle([60, 115, 800, 1045], radius=18, fill="#ffffff", outline="#cbd5e1", width=3)
    d.ellipse([380, 155, 480, 255], fill="#eff6ff", outline="#3b82f6", width=3)
    d.text((430, 205), "TIME", font=get_font(28, bold=True), fill="#2563eb", anchor="mm")

    d.text((430, 285), "사원번호로 초간편 1초 입장", font=get_font(32, bold=True), fill="#0f172a", anchor="mm")
    d.text((430, 325), "복잡한 비밀번호 없이 본인 사원번호 6자리만 쏙 입력하세요.", font=get_font(18), fill="#64748b", anchor="mm")

    d.text((120, 385), "1. 사원번호 (6자리 숫자 입력)", font=get_font(20, bold=True), fill="#1e293b")
    d.rounded_rectangle([120, 420, 740, 505], radius=10, fill="#ffffff", outline="#2563eb", width=3)
    d.text((150, 445), "123456", font=get_font(32, bold=True), fill="#0f172a")
    d.text((620, 450), "✓ 정상", font=get_font(18, bold=True), fill="#10b981")

    d.rounded_rectangle([120, 535, 740, 625], radius=12, fill="#2563eb")
    d.text((430, 580), "2. 입장하기 → (원클릭!)", font=get_font(26, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([120, 660, 740, 810], radius=12, fill="#f0fdf4", outline="#86efac", width=2)
    d.text((145, 685), "★ [사번 자동 기억 기능 탑재]", font=get_font(20, bold=True), fill="#166534")
    d.text((145, 725), "• 브라우저가 사번을 안전하게 기억하므로 매번 재입력할 필요 없습니다.", font=get_font(16), fill="#15803d")
    d.text((145, 765), "• 다음 접속 시에는 바로 [입장하기 →] 버튼만 누르면 1초 만에 입장 완료!", font=get_font(16, bold=True), fill="#166534")

    d.rounded_rectangle([120, 840, 740, 990], radius=10, fill="#fefce8", outline="#facc15")
    d.text((145, 865), "※ 공용 PC 및 스마트폰 접속 팁", font=get_font(18, bold=True), fill="#854d0e")
    d.text((145, 905), "• 공용 PC에서 다른 사번으로 로그인할 때는 사번을 지우고 새로 입력합니다.", font=get_font(15), fill="#a16207")
    d.text((145, 945), "• 퇴근 시에는 상단 [로그아웃]을 누르면 사번이 안전하게 분리됩니다.", font=get_font(15), fill="#a16207")

    # 우측: 신규 방문자 1초 자동 가입 창
    d.rounded_rectangle([840, 115, 1540, 1045], radius=18, fill="#ffffff", outline="#cbd5e1", width=3)
    d.rounded_rectangle([840, 115, 1540, 205], radius=16, fill="#f1f5f9")
    d.text((1190, 160), "✨ 미등록 사번 최초 방문 시 (1회만 진행)", font=get_font(24, bold=True), fill="#0f172a", anchor="mm")

    d.text((890, 245), "사원번호가 아직 등록되지 않은 분은 성명과 소속팀만 고르면 즉시 완료!", font=get_font(18), fill="#64748b")

    d.text((890, 310), "사원번호 (자동 입력됨)", font=get_font(19, bold=True), fill="#334155")
    d.rounded_rectangle([890, 345, 1490, 420], radius=8, fill="#f8fafc", outline="#cbd5e1")
    d.text((920, 368), "123456  (입력한 사번 고정)", font=get_font(20, bold=True), fill="#64748b")

    d.text((890, 455), "성명 (이름 입력)", font=get_font(19, bold=True), fill="#334155")
    d.rounded_rectangle([890, 490, 1490, 565], radius=8, fill="#ffffff", outline="#3b82f6", width=2)
    d.text((920, 513), "홍길동", font=get_font(22, bold=True), fill="#0f172a")

    d.text((890, 600), "소속팀 선택 (드롭다운)", font=get_font(19, bold=True), fill="#334155")
    d.rounded_rectangle([890, 635, 1490, 710], radius=8, fill="#ffffff", outline="#3b82f6", width=2)
    d.text((920, 658), "제어실 ▼", font=get_font(22, bold=True), fill="#0f172a")

    d.rounded_rectangle([890, 755, 1490, 845], radius=12, fill="#10b981")
    d.text((1190, 800), "[ ✓ 등록 및 입장하기 ] (1초 만에 완료)", font=get_font(24, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([890, 885, 1490, 1000], radius=12, fill="#eff6ff", outline="#bfdbfe")
    d.text((920, 915), "★ 한 번만 등록해두시면 이후에는 모든 기기(폰/PC)에서", font=get_font(18, bold=True), fill="#1e40af")
    d.text((920, 955), "사원번호 6자리만 넣으면 회원가입 없이 바로 대시보드로 입장됩니다.", font=get_font(16), fill="#2563eb")

    draw_smart_pin(d, (430, 462), (430, 395), "①", "사번 6자리 입력", color="#2563eb")
    draw_smart_pin(d, (430, 580), (430, 645), "②", "[입장하기] 원클릭!", color="#10b981")
    draw_smart_pin(d, (1190, 800), (1190, 725), "③", "신규 사원 즉시 자동 가입", color="#ea580c")

    path = IMG_DIR / "mockup_login.png"
    img.save(path)
    return str(path)


# 3. 화면 기능 및 선택버튼 상세 안내 목업
def create_mockup_dashboard_buttons():
    """삽화: 상단 툴바 및 주요 선택버튼의 상세 설명(어떤 때 해당 버튼을 쓰는지)"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  상단 버튼 상세 가이드", url="https://overtime.company.com/dashboard")

    # 상단 툴바
    d.rounded_rectangle([40, 105, 1560, 220], radius=14, fill="#1e293b", outline="#334155")
    d.text((70, 140), "TIME  스마트 특근 관리 시스템", font=get_font(24, bold=True), fill="#ffffff")
    d.text((70, 175), "홍길동 사원 (소속: 제어실) 님 환영합니다.", font=get_font(17), fill="#94a3b8")

    btn_defs = [
        (820, "🔄 새로고침", "#3b82f6"),
        (975, "📱 모바일 접속 QR", "#10b981"),
        (1160, "💡 건의사항 소통함", "#f59e0b"),
        (1345, "💾 웹 저장", "#64748b"),
        (1455, "📂 열기", "#64748b")
    ]
    for bx, b_label, b_color in btn_defs:
        d.rounded_rectangle([bx, 140, bx + 130, 185], radius=8, fill=b_color)
        d.text((bx + 65, 162), b_label, font=get_font(15, bold=True), fill="#ffffff", anchor="mm")

    # 하단 뷰 전환 버튼
    d.rounded_rectangle([40, 240, 1560, 310], radius=10, fill="#ffffff", outline="#cbd5e1")
    d.text((70, 260), "화면 전환: ", font=get_font(18, bold=True), fill="#334155")
    d.rounded_rectangle([170, 252, 290, 298], radius=6, fill="#2563eb")
    d.text((230, 275), "📋 목록 보기", font=get_font(16, bold=True), fill="#ffffff", anchor="mm")
    d.rounded_rectangle([305, 252, 425, 298], radius=6, fill="#f1f5f9", outline="#cbd5e1")
    d.text((365, 275), "📅 달력 보기", font=get_font(16, bold=True), fill="#475569", anchor="mm")

    d.text((500, 263), "신청 내역 액션 버튼: ", font=get_font(18, bold=True), fill="#334155")
    d.rounded_rectangle([680, 252, 750, 298], radius=6, fill="#f1f5f9", outline="#cbd5e1")
    d.text((715, 275), "수정", font=get_font(15, bold=True), fill="#334155", anchor="mm")
    d.rounded_rectangle([765, 252, 835, 298], radius=6, fill="#fee2e2", outline="#fca5a5")
    d.text((800, 275), "취소", font=get_font(15, bold=True), fill="#b91c1c", anchor="mm")
    d.rounded_rectangle([850, 252, 1020, 298], radius=6, fill="#10b981")
    d.text((935, 275), "🎯 특근완료 확정", font=get_font(15, bold=True), fill="#ffffff", anchor="mm")

    button_cards = [
        ("🔄 [새로고침] 버튼", "#eff6ff", "#2563eb",
         "어떤 때 누르나요?",
         "• 관리자가 내 특근을 승인했는지 실시간 상태를 즉시 확인할 때\n"
         "• 다른 기기(스마트폰)에서 신청한 내역을 회사 PC에서 동기화할 때\n"
         "• 화면을 켜둔 채 시간이 오래 지나 최신 정보를 불러오고 싶을 때"),

        ("📱 [모바일 접속 QR] 버튼", "#ecfdf5", "#10b981",
         "어떤 때 누르나요?",
         "• 스마트폰으로 특근을 신청하고 싶을 때 (PC 화면 QR을 폰 카메라로 찰칵!)\n"
         "• 퇴근길, 자택, 휴일에 이동 중 모바일로 간편 신청하고 싶을 때\n"
         "• 동료 사원에게 모바일 접속 링크를 보여주거나 공유할 때"),

        ("💡 [건의사항 소통함] 버튼", "#fef3c7", "#f59e0b",
         "어떤 때 누르나요?",
         "• 사번이나 이름 노출 없이 100% 무기명으로 회사/부서에 의견 낼 때\n"
         "• 휴일 구내식당, 휴게실 냉난방, 근무환경 개선 등을 익명 건의할 때\n"
         "• 관리자가 남겨준 공식 검토 답변을 확인하고 싶을 때"),

        ("📋 [목록보기] vs 📅 [달력보기]", "#f8fafc", "#64748b",
         "어떤 때 누르나요?",
         "• [목록보기]: 내가 신청한 모든 특근 내역과 승인 상태를 표 형태로 한눈에 점검할 때\n"
         "• [달력보기]: 달력 위에 내가 일하는 날짜를 직관적으로 확인하고 날짜를 콕 찍어 신청할 때"),

        ("🎯 [특근완료 확정] 버튼 (핵심!)", "#fdf2f8", "#ec4899",
         "어떤 때 누르나요?",
         "• 관리자가 사전 승인한 특근에 대해 실제 휴일 근무를 마쳤을 때!\n"
         "• '나 오늘 실제로 출근해서 근무 끝냈습니다'라고 완료 피드백을 보낼 때\n"
         "• 이 버튼을 누르면 상태가 [🎯 확정완료]로 바뀌며 최종 정산에 반영됩니다!"),

        ("💾 [웹 저장] / 📂 [웹 열기] 버튼", "#f1f5f9", "#475569",
         "어떤 때 누르나요?",
         "• [웹 저장]: 내 신청 내역을 내 컴퓨터에 JSON 파일로 다운로드 백업할 때\n"
         "• [웹 열기]: 과거에 백업해둔 신청 내역 파일을 다시 불러올 때")
    ]

    grid_positions = [
        (45, 335, 480, 340),
        (560, 335, 480, 340),
        (1075, 335, 480, 340),
        (45, 700, 480, 350),
        (560, 700, 480, 350),
        (1075, 700, 480, 350)
    ]

    for (b_name, b_bg, b_line, when_title, when_desc), (gx, gy, gw, gh) in zip(button_cards, grid_positions):
        d.rounded_rectangle([gx, gy, gx + gw, gy + gh], radius=14, fill="#ffffff", outline=b_line, width=2)
        d.rounded_rectangle([gx, gy, gx + gw, gy + 55], radius=12, fill=b_bg)
        d.text((gx + 20, gy + 27), b_name, font=get_font(19, bold=True), fill="#0f172a", anchor="lm")

        d.text((gx + 20, gy + 80), f"👉 {when_title}", font=get_font(16, bold=True), fill=b_line)
        d.text((gx + 20, gy + 115), when_desc, font=get_font(15), fill="#334155", line_spacing=1.3)

    path = IMG_DIR / "mockup_dashboard_buttons.png"
    img.save(path)
    return str(path)


# 4. 특근 신청 순서 & 선택 버튼 상세 설명 목업
def create_mockup_calendar_apply():
    """삽화: 접속 후 순서에 맞는 순차적인 설명 & 선택 버튼 상세 설명 & 선택 항목 예시"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  특근 신청 순차 가이드", url="https://overtime.company.com/apply")

    # 상단 4단계 순서도 바
    d.rounded_rectangle([40, 105, 1560, 175], radius=12, fill="#ffffff", outline="#cbd5e1", width=2)
    step_bar = [
        ("1. 달력 날짜 콕!", "#3b82f6"),
        ("➔", "#94a3b8"),
        ("2. 특근 분류 버튼 선택", "#10b981"),
        ("➔", "#94a3b8"),
        ("3. 항목 입력 (예시 참고)", "#f59e0b"),
        ("➔", "#94a3b8"),
        ("4. [특근 신청 저장] 클릭!", "#ea580c")
    ]
    cur_x = 70
    for txt, col in step_bar:
        d.text((cur_x, 140), txt, font=get_font(20, bold=True), fill=col, anchor="lm")
        cur_x += len(txt) * 16 + 30

    # 좌측: 입력 폼 섹션
    d.rounded_rectangle([40, 195, 800, 1055], radius=16, fill="#ffffff", outline="#cbd5e1", width=2)

    d.text((70, 220), "STEP 2. 특근 분류 선택 (★어떤 때 누르나요?)", font=get_font(21, bold=True), fill="#0f172a")

    type_buttons = [
        ("🟢 일반휴일 (가장 많이 씀!)",
         "#eff6ff", "#2563eb",
         "👉 어떤 때 누르나요?\n"
         "• 주말(토/일) 나와서 정상 특근할 때 선택합니다. (기본 선택됨)\n"
         "• 휴일 1일 근무 시 1.0일 특근으로 자동 산정됩니다."),

        ("🔵 대체근무",
         "#f8fafc", "#0284c7",
         "👉 어떤 때 누르나요?\n"
         "• 평일에 쉬고 주말에 대체 출근하거나, 주말 근무 대신 다른 날 쉴 때 선택!\n"
         "• 나중에 대체휴무(대휴)를 쓸 수 있는 특근입니다."),

        ("🟣 법정휴일",
         "#faf5ff", "#9333ea",
         "👉 어떤 때 누르나요?\n"
         "• 신정, 설날, 추석, 삼일절, 광복절 등 법정 공휴일에 출근할 때 선택!\n"
         "• 법정 유급휴일 특근으로 정산표에 별도 자동 집계됩니다.")
    ]

    for idx, (t_title, t_bg, t_border, t_desc) in enumerate(type_buttons):
        ty = 260 + idx * 115
        d.rounded_rectangle([70, ty, 770, ty + 105], radius=10, fill=t_bg, outline=t_border, width=2)
        d.text((95, ty + 25), t_title, font=get_font(18, bold=True), fill=t_border, anchor="lm")
        d.text((95, ty + 68), t_desc, font=get_font(14), fill="#334155", line_spacing=1.2)

    d.text((70, 625), "STEP 1. 일할 날짜 (달력 클릭 시 자동 세팅)", font=get_font(19, bold=True), fill="#0f172a")
    d.rounded_rectangle([70, 655, 410, 715], radius=8, fill="#eff6ff", outline="#3b82f6", width=2)
    d.text((95, 685), "시작일: 2026-09-12 (토)", font=get_font(17, bold=True), fill="#1e40af", anchor="lm")
    d.rounded_rectangle([430, 655, 770, 715], radius=8, fill="#eff6ff", outline="#3b82f6", width=2)
    d.text((455, 685), "종료일: 2026-09-12 (토)", font=get_font(17, bold=True), fill="#1e40af", anchor="lm")

    d.text((70, 735), "STEP 3. 상세 내용 입력 (선택적 항목 예시)", font=get_font(19, bold=True), fill="#0f172a")

    d.rounded_rectangle([70, 765, 410, 835], radius=8, fill="#f8fafc", outline="#cbd5e1")
    d.text((90, 785), "프로젝트 번호 (선택)", font=get_font(14, bold=True), fill="#64748b")
    d.text((90, 812), "BT2601-L1 (예: 프로젝트/라인)", font=get_font(16, bold=True), fill="#0f172a")

    d.rounded_rectangle([430, 765, 770, 835], radius=8, fill="#f8fafc", outline="#cbd5e1")
    d.text((450, 785), "근무 장소 (선택)", font=get_font(14, bold=True), fill="#64748b")
    d.text((450, 812), "본사 5층 제어실 (예: 공장/제어실)", font=get_font(16, bold=True), fill="#0f172a")

    d.rounded_rectangle([70, 850, 770, 930], radius=8, fill="#f8fafc", outline="#cbd5e1")
    d.text((90, 870), "특근 사유 (구체적 업무 내용 권장)", font=get_font(14, bold=True), fill="#64748b")
    d.text((90, 900), "1공장 제어설비 정기 점검 및 시운전 긴급 지원", font=get_font(16, bold=True), fill="#0f172a")

    d.rounded_rectangle([70, 950, 770, 1035], radius=12, fill="#10b981")
    d.text((420, 992), "STEP 4. [💾 특근 신청 저장하기] 클릭! (신청 완료)", font=get_font(22, bold=True), fill="#ffffff", anchor="mm")

    # 우측: 달력 인터랙티브 목업
    d.rounded_rectangle([840, 195, 1560, 1055], radius=16, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((870, 225), "2026년 9월 특근 달력 (날짜를 콕 누르세요!)", font=get_font(23, bold=True), fill="#0f172a")
    d.text((870, 260), "▶ 달력 날짜를 클릭하면 시작일과 종료일이 자동으로 맞춰집니다.", font=get_font(16), fill="#2563eb")

    days = ["일", "월", "화", "수", "목", "금", "토"]
    col_w = 95
    for i, day in enumerate(days):
        col_c = "#ef4444" if i == 0 else ("#2563eb" if i == 6 else "#475569")
        d.text((870 + i * col_w + 45, 305), day, font=get_font(18, bold=True), fill=col_c, anchor="mm")

    for row in range(5):
        for col in range(7):
            day_num = row * 7 + col - 1
            if 1 <= day_num <= 30:
                cx = 870 + col * col_w
                cy = 340 + row * 105
                is_sel = (day_num == 12)
                if is_sel:
                    d.rounded_rectangle([cx, cy, cx + 90, cy + 95], radius=10, fill="#dbeafe", outline="#2563eb", width=3)
                    d.text((cx + 12, cy + 12), str(day_num), font=get_font(20, bold=True), fill="#1d4ed8")
                    d.rounded_rectangle([cx + 8, cy + 50, cx + 82, cy + 85], radius=6, fill="#2563eb")
                    d.text((cx + 45, cy + 67), "선택됨!", font=get_font(13, bold=True), fill="#ffffff", anchor="mm")
                else:
                    d.rounded_rectangle([cx, cy, cx + 90, cy + 95], radius=8, fill="#f8fafc", outline="#e2e8f0")
                    num_col = "#ef4444" if col == 0 else ("#2563eb" if col == 6 else "#1e293b")
                    d.text((cx + 12, cy + 12), str(day_num), font=get_font(18), fill=num_col)

    d.rounded_rectangle([870, 880, 1530, 1025], radius=12, fill="#fefce8", outline="#facc15", width=2)
    d.text((895, 905), "💡 [연속 근무 시 간편 팁]", font=get_font(19, bold=True), fill="#854d0e")
    d.text((895, 942), "• 토/일 이틀 모두 특근할 때는 달력에서 토요일(12일)을 먼저 누르고,", font=get_font(16), fill="#a16207")
    d.text((895, 977), "• 왼쪽 폼의 [종료일]을 일요일(13일)로 변경하시면 2일 연속 신청됩니다!", font=get_font(16, bold=True), fill="#b45309")

    draw_smart_pin(d, (1335, 595), (1335, 530), "①", "일할 날짜 콕 누르기", color="#2563eb")
    draw_smart_pin(d, (420, 260), (420, 200), "②", "특근 분류 선택", color="#10b981")
    draw_smart_pin(d, (420, 950), (420, 890), "④", "[신청 저장] 클릭!", color="#ea580c")

    path = IMG_DIR / "mockup_apply_flow.png"
    img.save(path)
    return str(path)


# 5. 나의 특근 신청 내역 & [🎯 특근완료 확정] 목업
def create_mockup_my_records_guide():
    """삽화: 나의 신청 내역, 승인 확인, 수정/취소, 그리고 실제 근무 후 [확정] 버튼 가이드"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  나의 특근 내역 및 확정 가이드", url="https://overtime.company.com/my-records")

    # 상단 4단계 상태 라이프사이클 프로그레스 바
    d.rounded_rectangle([40, 105, 1560, 215], radius=14, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((70, 125), "특근 처리 4단계 라이프사이클 (안심하고 확인하세요!)", font=get_font(21, bold=True), fill="#0f172a")

    stages = [
        ("1. ⏳ 승인대기", "신청 직후 (관리자 확인 전)", "#fef3c7", "#d97706"),
        ("➔", "", "", "#94a3b8"),
        ("2. ✓ 승인완료", "관리자가 사전 승인 완료!", "#ecfdf5", "#10b981"),
        ("➔", "", "", "#94a3b8"),
        ("3. 🎯 확정완료", "실제 휴일 근무 후 [확정] 클릭!", "#eff6ff", "#2563eb"),
        ("➔", "", "", "#94a3b8"),
        ("4. 🟣 검토완료", "관리자 최종 마감 검토 완료", "#faf5ff", "#9333ea")
    ]
    cur_x = 70
    for s_title, s_desc, s_bg, s_col in stages:
        if s_desc:
            d.rounded_rectangle([cur_x, 155, cur_x + 280, 202], radius=8, fill=s_bg, outline=s_col)
            d.text((cur_x + 15, 178), s_title, font=get_font(15, bold=True), fill=s_col, anchor="lm")
            d.text((cur_x + 130, 178), f"| {s_desc}", font=get_font(13), fill="#475569", anchor="lm")
            cur_x += 300
        else:
            d.text((cur_x + 10, 178), s_title, font=get_font(18, bold=True), fill=s_col, anchor="lm")
            cur_x += 45

    # 내 신청 내역 테이블 카드
    d.rounded_rectangle([40, 235, 1560, 680], radius=16, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((70, 260), "📋 나의 특근 신청 내역 (최신 등록순)", font=get_font(23, bold=True), fill="#0f172a")

    d.rounded_rectangle([60, 295, 1540, 345], radius=8, fill="#1e293b")
    headers = [
        (80, "일할 날짜 / 요일"),
        (280, "특근분류"),
        (420, "시간"),
        (540, "프로젝트 / 근무장소"),
        (850, "특근 사유 (업무 내용)"),
        (1160, "진행 상태"),
        (1360, "작업 (수정/취소/확정)")
    ]
    for hx, h_txt in headers:
        d.text((hx, 320), h_txt, font=get_font(15, bold=True), fill="#ffffff", anchor="lm")

    d.rounded_rectangle([60, 360, 1540, 460], radius=10, fill="#f8fafc", outline="#3b82f6", width=2)
    d.text((80, 395), "2026-09-12 (토)", font=get_font(17, bold=True), fill="#0f172a")
    d.text((80, 425), "08:00 ~ 17:00", font=get_font(14), fill="#64748b")

    d.rounded_rectangle([270, 390, 370, 430], radius=6, fill="#dbeafe")
    d.text((320, 410), "일반휴일", font=get_font(15, bold=True), fill="#1d4ed8", anchor="mm")

    d.text((430, 410), "1.0일 (8h)", font=get_font(16, bold=True), fill="#0f172a", anchor="lm")
    d.text((540, 410), "BT2601-L1  |  5층 제어실", font=get_font(15), fill="#334155", anchor="lm")
    d.text((850, 410), "1공장 제어설비 정기 점검 및 시운전 지원", font=get_font(15), fill="#334155", anchor="lm")

    d.rounded_rectangle([1150, 390, 1260, 430], radius=6, fill="#dcfce7")
    d.text((1205, 410), "✓ 승인완료", font=get_font(15, bold=True), fill="#15803d", anchor="mm")

    d.rounded_rectangle([1280, 385, 1345, 435], radius=6, fill="#f1f5f9", outline="#cbd5e1")
    d.text((1312, 410), "수정", font=get_font(14, bold=True), fill="#334155", anchor="mm")

    d.rounded_rectangle([1355, 385, 1420, 435], radius=6, fill="#fee2e2", outline="#fca5a5")
    d.text((1387, 410), "취소", font=get_font(14, bold=True), fill="#b91c1c", anchor="mm")

    d.rounded_rectangle([1430, 385, 1530, 435], radius=8, fill="#10b981")
    d.text((1480, 410), "🎯 확정", font=get_font(15, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([60, 480, 1540, 560], radius=10, fill="#ffffff", outline="#e2e8f0")
    d.text((80, 520), "2026-09-05 (토)", font=get_font(16, bold=True), fill="#0f172a", anchor="lm")
    d.text((280, 520), "일반휴일", font=get_font(15), fill="#475569", anchor="lm")
    d.text((430, 520), "1.0일 (8h)", font=get_font(15), fill="#475569", anchor="lm")
    d.text((540, 520), "BT2601-L1  |  5층 제어실", font=get_font(15), fill="#475569", anchor="lm")
    d.text((850, 520), "전산 시스템 정기 백업 작업", font=get_font(15), fill="#475569", anchor="lm")
    d.rounded_rectangle([1150, 500, 1260, 540], radius=6, fill="#eff6ff")
    d.text((1205, 520), "🎯 확정완료", font=get_font(15, bold=True), fill="#1d4ed8", anchor="mm")
    d.text((1360, 520), "(근무완료 확인됨)", font=get_font(14), fill="#64748b", anchor="lm")

    d.rounded_rectangle([40, 705, 1560, 1055], radius=16, fill="#fdf2f8", outline="#f472b6", width=2)
    d.text((70, 735), "★ [사용자 필독] [🎯 특근완료 확정] 버튼은 언제 누르나요?", font=get_font(23, bold=True), fill="#9d174d")

    confirm_tips = [
        ("1. 실제 휴일 근무를 마친 후 클릭!",
         "관리자가 사전에 승인해 준 특근 일정에 대해, 실제로 휴일에 출근하여 근무를 마친 당일 퇴근 시(또는 다음 날) 클릭합니다."),
        ("2. 관리자에게 '출근 완료' 피드백 전송",
         "이 버튼을 누르면 상태가 [🎯 확정완료]로 바뀌며, 관리자는 해당 사원이 실제로 출근하여 일했음을 즉시 확인하게 됩니다."),
        ("3. 잘못 신청했거나 취소된 경우엔 [취소] 클릭",
         "휴일 근무가 취소되었거나 나가지 않은 경우에는 [확정]을 누르지 마시고 [취소] 버튼을 눌러 일정을 삭제합니다.")
    ]
    for c_idx, (c_tit, c_sub) in enumerate(confirm_tips):
        cy = 780 + c_idx * 85
        d.rounded_rectangle([70, cy, 1530, cy + 70], radius=10, fill="#ffffff", outline="#fbcfe8")
        d.text((95, cy + 22), c_tit, font=get_font(17, bold=True), fill="#be185d", anchor="lm")
        d.text((95, cy + 48), c_sub, font=get_font(15), fill="#334155", anchor="lm")

    draw_smart_pin(d, (1205, 410), (1205, 345), "①", "관리자 승인 확인", color="#15803d")
    draw_smart_pin(d, (1312, 410), (1312, 470), "②", "[수정/취소]", color="#475569")
    draw_smart_pin(d, (1480, 410), (1480, 345), "③", "근무 후 [🎯 확정] 클릭!", color="#db2777")

    path = IMG_DIR / "mockup_my_records_guide.png"
    img.save(path)
    return str(path)


# 6. 스마트폰 모바일 1분 퀵 가이드 목업
def create_mockup_mobile_flow():
    """삽화: 스마트폰 형태 3단계 (QR 스캔 -> 사번 로그인 -> 원클릭 신청) 및 앱 등록 가이드"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  스마트폰 모바일 1분 퀵 가이드", url="https://overtime.company.com/mobile")

    d.rounded_rectangle([40, 105, 1560, 180], radius=12, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((70, 125), "📱 스마트폰으로 10초 만에 끝내는 초간편 모바일 특근 신청", font=get_font(26, bold=True), fill="#0f172a")
    d.text((70, 155), "별도 앱 설치 없이 스마트폰 카메라로 QR만 찍으면 즉시 신청 가능하며, 홈 화면에 추가 시 전용 앱으로 등록됩니다.", font=get_font(17), fill="#64748b")

    phones = [
        ("STEP 1. 스마트폰 카메라로 QR 스캔",
         "#3b82f6",
         [("화면 상단 [모바일 접속 QR]을", 30),
          ("스마트폰 기본 카메라로 비추기만 하면", 30),
          ("1초 만에 노란색 접속 링크가 뜹니다.", 30),
          ("터치하면 모바일 브라우저로 즉시 연결!", 30)],
         "① 카메라로 QR 스캔"),

        ("STEP 2. 사원번호 6자리 터치 로그인",
         "#10b981",
         [("본인의 사원번호 6자리를", 30),
          ("스마트폰 화면에서 가볍게 터치 입력 후", 30),
          ("[입장하기 →] 버튼을 터치합니다.", 30),
          ("※ 모바일에서도 사번이 안전하게 자동 기억됨!", 30)],
         "② 사번 6자리 터치"),

        ("STEP 3. 달력 날짜 터치 & 저장",
         "#f59e0b",
         [("모바일 달력에서 일할 날짜를 콕!", 30),
          ("특근 분류(일반휴일/대체/법정) 선택 후", 30),
          ("[💾 특근 신청 저장하기] 터치!", 30),
          ("언제 어디서나 10초 만에 신청 끝!", 30)],
         "③ 날짜 터치 후 저장")
    ]

    p_w = 460
    p_gap = 50
    start_x = 55

    for idx, (p_title, p_color, p_lines, p_badge) in enumerate(phones):
        px = start_x + idx * (p_w + p_gap)
        py = 205
        ph = 750

        d.rounded_rectangle([px, py, px + p_w, py + ph], radius=32, fill="#0f172a", outline="#334155", width=4)
        d.rounded_rectangle([px + p_w // 2 - 60, py + 12, px + p_w // 2 + 60, py + 26], radius=7, fill="#1e293b")
        d.ellipse([px + p_w // 2 + 45, py + 15, px + p_w // 2 + 55, py + 25], fill="#3b82f6")

        sx1, sy1 = px + 16, py + 40
        sx2, sy2 = px + p_w - 16, py + ph - 25
        d.rounded_rectangle([sx1, sy1, sx2, sy2], radius=20, fill="#ffffff")

        d.rounded_rectangle([sx1, sy1, sx2, sy1 + 65], radius=18, fill=p_color)
        d.text((px + p_w // 2, sy1 + 32), p_title, font=get_font(18, bold=True), fill="#ffffff", anchor="mm")

        if idx == 0:
            d.rounded_rectangle([sx1 + 30, sy1 + 95, sx2 - 30, sy1 + 360], radius=16, fill="#0f172a")
            d.rounded_rectangle([sx1 + 70, sy1 + 135, sx2 - 70, sy1 + 320], radius=12, fill="#ffffff", outline="#38bdf8", width=3)
            d.text((px + p_w // 2, sy1 + 225), "📱 QR CODE\nSCANNING...", font=get_font(22, bold=True), fill="#0f172a", anchor="mm")
            d.rounded_rectangle([sx1 + 45, sy1 + 380, sx2 - 45, sy1 + 440], radius=10, fill="#fef08a", outline="#eab308")
            d.text((px + p_w // 2, sy1 + 410), "👉 overtime-system.trycloudflare.com", font=get_font(14, bold=True), fill="#854d0e", anchor="mm")
        elif idx == 1:
            d.rounded_rectangle([sx1 + 25, sy1 + 100, sx2 - 25, sy1 + 440], radius=14, fill="#f8fafc", outline="#cbd5e1")
            d.text((px + p_w // 2, sy1 + 140), "사원번호 입력", font=get_font(20, bold=True), fill="#0f172a", anchor="mm")
            d.rounded_rectangle([sx1 + 50, sy1 + 175, sx2 - 50, sy1 + 240], radius=8, fill="#ffffff", outline="#10b981", width=2)
            d.text((px + p_w // 2, sy1 + 208), "123456", font=get_font(26, bold=True), fill="#0f172a", anchor="mm")
            d.rounded_rectangle([sx1 + 50, sy1 + 265, sx2 - 50, sy1 + 335], radius=10, fill="#10b981")
            d.text((px + p_w // 2, sy1 + 300), "입장하기 →", font=get_font(22, bold=True), fill="#ffffff", anchor="mm")
            d.text((px + p_w // 2, sy1 + 380), "★ 사번 자동 기억 탑재", font=get_font(15, bold=True), fill="#15803d", anchor="mm")
        else:
            d.rounded_rectangle([sx1 + 25, sy1 + 85, sx2 - 25, sy1 + 440], radius=14, fill="#f8fafc", outline="#cbd5e1")
            d.text((px + p_w // 2, sy1 + 115), "9월 12일 (토) 특근 선택", font=get_font(18, bold=True), fill="#1e40af", anchor="mm")
            d.rounded_rectangle([sx1 + 45, sy1 + 145, sx2 - 45, sy1 + 205], radius=8, fill="#eff6ff", outline="#2563eb")
            d.text((px + p_w // 2, sy1 + 175), "● 일반휴일 (1.0일)", font=get_font(17, bold=True), fill="#2563eb", anchor="mm")
            d.rounded_rectangle([sx1 + 45, sy1 + 225, sx2 - 45, sy1 + 300], radius=8, fill="#ffffff", outline="#cbd5e1")
            d.text((sx1 + 60, sy1 + 262), "사유: 정기 점검", font=get_font(15), fill="#475569", anchor="lm")
            d.rounded_rectangle([sx1 + 45, sy1 + 325, sx2 - 45, sy1 + 395], radius=10, fill="#f59e0b")
            d.text((px + p_w // 2, sy1 + 360), "💾 특근 신청 저장", font=get_font(20, bold=True), fill="#ffffff", anchor="mm")

        for l_idx, (line_txt, _) in enumerate(p_lines):
            ly = sy1 + 465 + l_idx * 45
            d.text((px + p_w // 2, ly), line_txt, font=get_font(15), fill="#1e293b", anchor="mm")

        draw_smart_pin(d, (px + p_w // 2, py + ph - 25), (px + p_w // 2, py + ph + 25), f"0{idx+1}", p_badge, color=p_color)

    d.rounded_rectangle([40, 985, 1560, 1070], radius=12, fill="#eff6ff", outline="#93c5fd", width=2)
    d.text((70, 1012), "★ 모바일 전용 앱처럼 등록하기: 브라우저 메뉴에서 [홈 화면에 추가]를 터치하세요!", font=get_font(20, bold=True), fill="#1d4ed8")
    d.text((70, 1042), "스마트폰 바탕화면에 앱 아이콘이 바로 생성되어, 다음부터는 웹 주소 입력 없이 아이콘 터치 한 번으로 즉시 실행됩니다.", font=get_font(16), fill="#2563eb")

    path = IMG_DIR / "mockup_mobile_flow.png"
    img.save(path)
    return str(path)


# 7. 100% 무기명 건의사항 소통함 목업
def create_mockup_suggestion_box():
    """삽화: 100% 무기명 건의사항 소통함 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  무기명 건의사항 소통함", url="https://overtime.company.com/feedback")

    d.rounded_rectangle([250, 115, 1350, 1045], radius=20, fill="#ffffff", outline="#cbd5e1", width=3)
    d.rounded_rectangle([250, 115, 1350, 220], radius=18, fill="#1e293b")
    d.text((300, 168), "💡 사원 전용 100% 무기명 소통함", font=get_font(28, bold=True), fill="#ffffff", anchor="lm")
    d.rounded_rectangle([1180, 145, 1310, 190], radius=6, fill="#f59e0b")
    d.text((1245, 168), "완전 무기명", font=get_font(16, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([300, 250, 1300, 360], radius=12, fill="#ecfdf5", outline="#86efac", width=2)
    d.text((330, 280), "🛡️ [100% 무기명 익명성 절대 보장 안내]", font=get_font(20, bold=True), fill="#166534")
    d.text((330, 320), "시스템 DB에 사번, 성명, IP 주소 등 작성자를 식별할 수 있는 컬럼 자체가 존재하지 않아 절대 추적할 수 없습니다.", font=get_font(16), fill="#15803d")

    d.text((300, 395), "회사 및 부서에 바라는 점 / 개선 건의사항 (자유롭게 입력)", font=get_font(20, bold=True), fill="#0f172a")
    d.rounded_rectangle([300, 435, 1300, 635], radius=10, fill="#f8fafc", outline="#3b82f6", width=2)
    d.text((330, 465), "휴일 특근 시 구내식당 미운영에 따른 대체 식사 지원 방안과,", font=get_font(18), fill="#1e293b")
    d.text((330, 505), "제어실 휴게공간 냉난방 장비의 주말 탄력 가동을 검토 부탁드립니다.", font=get_font(18), fill="#1e293b")

    d.rounded_rectangle([300, 660, 1300, 745], radius=12, fill="#2563eb")
    d.text((800, 702), "📨 무기명으로 등록하기 (작성자 기록 없이 등록)", font=get_font(24, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([300, 775, 1300, 1010], radius=12, fill="#f8fafc", outline="#cbd5e1")
    d.text((330, 805), "💬 최근 등록된 건의사항 및 관리자 공식 답변 보기", font=get_font(19, bold=True), fill="#0f172a")

    d.rounded_rectangle([330, 840, 1270, 980], radius=8, fill="#ffffff", outline="#e2e8f0")
    d.text((350, 865), "건의: 주말 제어실 에어컨 자동 가동 요청의 건", font=get_font(16, bold=True), fill="#0f172a")
    d.rounded_rectangle([350, 895, 440, 925], radius=4, fill="#dcfce7")
    d.text((395, 910), "답변완료", font=get_font(13, bold=True), fill="#166534", anchor="mm")
    d.text((455, 910), "관리부서: 시설팀 협의 완료되어 금주 주말부터 08시~18시 정상 가동됩니다.", font=get_font(15), fill="#2563eb", anchor="lm")

    draw_smart_pin(d, (1245, 168), (1245, 105), "①", "완전 무기명 보장", color="#166534")
    draw_smart_pin(d, (800, 702), (800, 635), "②", "사번 노출 없이 등록", color="#2563eb")
    draw_smart_pin(d, (395, 910), (395, 845), "③", "관리자 공식 답변 열람", color="#ea580c")

    path = IMG_DIR / "mockup_suggestion_box.png"
    img.save(path)
    return str(path)


# 8. 관리자 캘린더 목업
def create_mockup_admin_calendar():
    """삽화: 관리자 월간 캘린더 & 일괄 승인 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  관리자 캘린더 & 승인", url="https://overtime.company.com/admin/calendar")

    d.rounded_rectangle([40, 105, 1560, 1060], radius=16, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((80, 125), "[관리자 모드] 월간 특근 캘린더 & 일괄 승인 관리", font=get_font(28, bold=True), fill="#0f172a")
    d.text((80, 160), "소속 팀원의 특근 일정을 한눈에 파악하고, 일자별 작업자 패널에서 당일 전원 일괄 확인을 지원합니다.", font=get_font(18), fill="#64748b")

    d.rounded_rectangle([80, 210, 820, 940], radius=12, fill="#f8fafc", outline="#cbd5e1")
    d.text((105, 235), "2026년 9월 (부서 전체 특근 현황)", font=get_font(22, bold=True), fill="#0f172a")

    days = ["일", "월", "화", "수", "목", "금", "토"]
    col_w = 98
    for i, day in enumerate(days):
        col_c = "#ef4444" if i == 0 else ("#2563eb" if i == 6 else "#475569")
        d.text((100 + i * col_w + 40, 280), day, font=get_font(16, bold=True), fill=col_c, anchor="mm")

    d.rounded_rectangle([435, 410, 690, 520], radius=8, fill="#eff6ff", outline="#3b82f6", width=2)
    d.text((445, 420), "12 (토)", font=get_font(17, bold=True), fill="#2563eb")
    d.rounded_rectangle([445, 445, 680, 475], radius=4, fill="#10b981")
    d.text((562, 460), "정진규 (일반 1일)", font=get_font(14, bold=True), fill="#ffffff", anchor="mm")
    d.rounded_rectangle([445, 482, 680, 512], radius=4, fill="#10b981")
    d.text((562, 497), "김철수 (일반 1일)", font=get_font(14, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([80, 960, 820, 1025], radius=8, fill="#f1f5f9")
    d.text((100, 985), "범례: ● 일반휴일(초록)  ● 대체근무(파랑)  ● 대휴사용(주황)  ● 보너스(보라)", font=get_font(15, bold=True), fill="#334155")

    d.rounded_rectangle([850, 210, 1530, 1025], radius=12, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((880, 235), "9월 12일 (토) 제어실 근무자 명단 (2명)", font=get_font(22, bold=True), fill="#0f172a")

    d.rounded_rectangle([880, 280, 1500, 360], radius=10, fill="#10b981")
    d.text((1190, 320), "[당일 전원 일괄 확인] (원클릭 승인!)", font=get_font(22, bold=True), fill="#ffffff", anchor="mm")

    d.rounded_rectangle([880, 385, 1500, 485], radius=8, fill="#f8fafc", outline="#e2e8f0")
    d.text((905, 410), "정진규 (실장) - 1공장 제어설비 정기 점검", font=get_font(19, bold=True), fill="#0f172a")
    d.text((905, 445), "특근분류: 일반휴일 (1.0일)  |  상태: [확인완료] 승인완료", font=get_font(16), fill="#10b981")

    d.rounded_rectangle([880, 505, 1500, 605], radius=8, fill="#f8fafc", outline="#e2e8f0")
    d.text((905, 530), "김철수 (팀원) - 전산 시스템 백업 지원", font=get_font(19, bold=True), fill="#0f172a")
    d.text((905, 565), "특근분류: 일반휴일 (1.0일)  |  상태: [대기중] 승인대기", font=get_font(16), fill="#ea580c")

    draw_smart_pin(d, (562, 460), (562, 375), "①", "팀원 일정 한눈에 파악", color="#2563eb")
    draw_smart_pin(d, (1190, 320), (1190, 250), "②", "원클릭 일괄 승인", color="#10b981")

    path = IMG_DIR / "mockup_admin_cal.png"
    img.save(path)
    return str(path)


# 9. 팀원 명부 및 권한 관리 목업
def create_mockup_user_mgmt():
    """삽화: 팀원 명부 관리 및 팀관리자 권한 토글 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  팀원 및 권한 관리", url="https://overtime.company.com/admin/users")

    d.rounded_rectangle([40, 105, 1560, 1060], radius=16, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((80, 125), "[팀원 및 권한 관리] (부서 팀관리자 지정 및 명부 관리)", font=get_font(28, bold=True), fill="#0f172a")
    d.text((80, 160), "총괄관리자는 전 사원의 소속 부서를 변경하고, 일반 팀원에게 '팀관리자' 권한을 원클릭 부여/회수합니다.", font=get_font(18), fill="#64748b")

    d.rounded_rectangle([80, 220, 1520, 940], radius=10, fill="#ffffff", outline="#cbd5e1")
    d.rectangle([80, 220, 1520, 285], fill="#1e293b")
    headers = [(105, "사번"), (265, "성명"), (465, "소속팀"), (725, "관리자 권한"), (1025, "최근 접속일시"), (1345, "관리 액션")]
    for hx, h_txt in headers:
        d.text((hx, 245), h_txt, font=get_font(17, bold=True), fill="#ffffff")

    rows = [
        ("113019", "정진규", "제어실", "총괄관리자 (슈퍼)", "2026-09-22 09:30", "보호됨 (삭제불가)"),
        ("2024001", "김철수", "기술연구팀", "팀관리자 [ON]", "2026-09-22 08:45", "[소속변경] [권한회수]"),
        ("2024002", "이영희", "품질관리팀", "일반사원 [OFF]", "2026-09-21 17:20", "[소속변경] [관리자부여]")
    ]
    for r_idx, (r_id, r_nm, r_dept, r_role, r_dt, r_act) in enumerate(rows):
        ry = 320 + r_idx * 80
        d.text((105, ry), r_id, font=get_font(17), fill="#0f172a")
        d.text((265, ry), r_nm, font=get_font(17, bold=True), fill="#0f172a")
        d.text((465, ry), r_dept, font=get_font(17), fill="#334155")
        if "총괄" in r_role:
            d.text((725, ry), r_role, font=get_font(17, bold=True), fill="#dc2626")
        elif "ON" in r_role:
            d.text((725, ry), r_role, font=get_font(17, bold=True), fill="#2563eb")
        else:
            d.text((725, ry), r_role, font=get_font(17), fill="#64748b")
        d.text((1025, ry), r_dt, font=get_font(16), fill="#64748b")
        d.text((1345, ry), r_act, font=get_font(16, bold=True), fill="#0284c7")

    draw_smart_pin(d, (785, 320), (785, 255), "①", "슈퍼관리자 영구 보호", color="#dc2626")
    draw_smart_pin(d, (785, 400), (785, 465), "②", "원클릭 팀관리자 권한 토글", color="#2563eb")

    path = IMG_DIR / "mockup_user_mgmt.png"
    img.save(path)
    return str(path)


# 10. 실특근 정산표 목업
def create_mockup_settlement():
    """삽화: 팀별/인원별 실특근 정산표 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  실특근 정산표", url="https://overtime.company.com/admin/settlement")

    d.rounded_rectangle([40, 105, 1560, 1060], radius=16, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((80, 125), "[실특근 자동 산정 및 정산표] (휴일별 완벽 구분 집계)", font=get_font(28, bold=True), fill="#0f172a")
    d.text((80, 160), "대체근무, 법정휴일, 일반휴일, 대휴사용을 독립 분리 산정하여 최종 실특근일을 1초 만에 자동 계산합니다.", font=get_font(18), fill="#64748b")

    kpis = [
        ("일반휴일 특근", "12.0일", "#eff6ff", "#2563eb"),
        ("대체근무 특근", "4.0일", "#f0fdf4", "#16a34a"),
        ("법정휴일 특근", "2.0일", "#faf5ff", "#9333ea"),
        ("★ 최종 실특근 합계", "15.0일", "#fefce8", "#ca8a04")
    ]
    for i, (k_tit, k_val, k_bg, k_col) in enumerate(kpis):
        kx = 80 + i * 360
        d.rounded_rectangle([kx, 220, kx + 330, 320], radius=12, fill=k_bg, outline=k_col, width=2)
        d.text((kx + 20, 245), k_tit, font=get_font(16, bold=True), fill="#475569")
        d.text((kx + 20, 280), k_val, font=get_font(30, bold=True), fill=k_col)

    d.rounded_rectangle([80, 360, 1520, 940], radius=10, fill="#ffffff", outline="#cbd5e1")
    d.rectangle([80, 360, 1520, 420], fill="#1e293b")
    headers = [(105, "사번"), (245, "성명"), (385, "소속팀"), (565, "대체근무"), (745, "법정휴일"), (925, "일반휴일"), (1105, "대휴사용"), (1345, "★ 최종 실특근")]
    for hx, h_txt in headers:
        d.text((hx, 380), h_txt, font=get_font(17, bold=True), fill="#ffffff")

    s_rows = [
        ("113019", "정진규", "제어실", "1.0일", "1.0일", "4.0일", "1.0일", "4.0일"),
        ("2024001", "김철수", "기술연구팀", "2.0일", "0.0일", "5.0일", "1.0일", "6.0일")
    ]
    for r_idx, row in enumerate(s_rows):
        ry = 450 + r_idx * 75
        for c_idx, val in enumerate(row):
            col_c = "#ca8a04" if c_idx == 7 else "#0f172a"
            d.text((headers[c_idx][0], ry), val, font=get_font(17, bold=(c_idx in [1, 7])), fill=col_c)

    draw_smart_pin(d, (1265, 270), (1265, 205), "①", "최종 실특근 1초 자동 산정", color="#ca8a04")

    path = IMG_DIR / "mockup_settlement.png"
    img.save(path)
    return str(path)


# 11. 엑셀 원장 27번 표준 양식 목업
def create_mockup_excel_27():
    """삽화: 엑셀 내보내기 27번 표준 서식 목업"""
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "#f8fafc")
    d = ImageDraw.Draw(img)
    draw_browser_frame(d, title="★ 스마트 특근 관리 시스템  |  엑셀 표준 서식", url="https://overtime.company.com/admin/export")

    d.rounded_rectangle([40, 105, 1560, 1060], radius=16, fill="#ffffff", outline="#cbd5e1", width=2)
    d.text((80, 125), "[엑셀(.xlsx) 내보내기] 27번 완벽 서식", font=get_font(28, bold=True), fill="#0f172a")
    d.text((80, 160), "날짜순 자동 전개, 순수 숫자 '1' 표기(마우스 드래그 SUM 수식 계산), 4대 휴일수 완벽 산출 지원", font=get_font(18), fill="#64748b")

    d.rounded_rectangle([80, 220, 1520, 520], radius=10, fill="#ffffff", outline="#16a34a", width=2)
    d.rectangle([80, 220, 1520, 280], fill="#16a34a")
    e_headers = [(105, "일자"), (275, "요일"), (385, "성명"), (515, "사번"), (655, "소속"), (795, "특근구분"), (985, "일수"), (1165, "대휴사용일")]
    for hx, h_txt in e_headers:
        d.text((hx, 240), h_txt, font=get_font(17, bold=True), fill="#ffffff")

    d.text((105, 305), "2026-09-05", font=get_font(17), fill="#0f172a")
    d.text((275, 305), "토", font=get_font(17, bold=True), fill="#2563eb")
    d.text((385, 305), "정진규", font=get_font(17, bold=True), fill="#0f172a")
    d.text((515, 305), "113019", font=get_font(17), fill="#475569")
    d.text((655, 305), "제어실", font=get_font(17), fill="#475569")
    d.text((795, 305), "일반휴일", font=get_font(17), fill="#2563eb")
    d.text((1005, 305), "1", font=get_font(20, bold=True), fill="#b45309")
    d.text((1165, 305), "-", font=get_font(17), fill="#64748b")

    d.text((105, 365), "2026-09-12", font=get_font(17), fill="#0f172a")
    d.text((275, 365), "토", font=get_font(17, bold=True), fill="#2563eb")
    d.text((385, 365), "김철수", font=get_font(17, bold=True), fill="#0f172a")
    d.text((515, 365), "2024001", font=get_font(17), fill="#475569")
    d.text((655, 365), "기술연구팀", font=get_font(17), fill="#475569")
    d.text((795, 365), "일반휴일", font=get_font(17), fill="#2563eb")
    d.text((1005, 365), "1", font=get_font(20, bold=True), fill="#b45309")
    d.text((1165, 365), "2026-09-25", font=get_font(17, bold=True), fill="#059669")

    d.rectangle([80, 425, 1520, 485], fill="#fef3c7")
    d.text((105, 445), "합계 (수식 자동 계산):", font=get_font(17, bold=True), fill="#92400e")
    d.text((980, 445), "=SUM(G4:G5) -> 2", font=get_font(19, bold=True), fill="#b45309")

    d.rounded_rectangle([80, 560, 1520, 1020], radius=12, fill="#f8fafc", outline="#cbd5e1")
    d.text((115, 595), "★ 27번 표준 엑셀 서식 3대 핵심 특징", font=get_font(21, bold=True), fill="#0f172a")
    d.text((115, 645), "1. 문자 '일' 없이 순수 숫자 1만 기재되어 엑셀에서 마우스 드래그만으로 합계(SUM)가 바로 나옵니다.", font=get_font(17), fill="#334155")
    d.text((115, 695), "2. 대휴사용 시 대휴사용일자가 같은 행에 나란히 연결되어 사전/사후 차감 관리가 투명합니다.", font=get_font(17), fill="#334155")
    d.text((115, 745), "3. Sheet 2(정산집계표) 및 Sheet 3(보너스명부)이 완벽히 분리되어 인사/경영지원팀 제출용으로 최적화되었습니다.", font=get_font(17), fill="#334155")

    draw_smart_pin(d, (1005, 305), (1005, 245), "①", "순수 숫자 1 표기 (수식 자동계산)", color="#b45309")

    path = IMG_DIR / "mockup_excel_27.png"
    img.save(path)
    return str(path)


# ----------------- 슬라이드 템플릿 헬퍼 -----------------

def add_visual_slide(prs, step_no, title, subtitle, items, img_path, hyperlink_url=None):
    """
    와이드스크린 16:9 슬라이드 레이아웃 (그림 중심 대형 배치):
    - 좌측: 핵심 요약 카드 (너비 3.5인치, 글자 수 대폭 축소 & 1~2줄 행동 요령 압축)
    - 우측: 대형 고화질 화면 목업 삽화 (너비 8.5인치, 높이 5.7인치)
    """
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    header_tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12.2), Inches(1.0))
    htf = header_tb.text_frame
    htf.word_wrap = True

    p_step = htf.paragraphs[0]
    p_step.text = f"STEP {step_no}  |  그림 보고 바로 따라 하기"
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

    card_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.45), Inches(3.5), Inches(5.7))
    card_box.fill.solid()
    card_box.fill.fore_color.rgb = COLOR_BG_CARD
    card_box.line.color.rgb = COLOR_BORDER

    card_tb_h = Inches(4.7) if hyperlink_url else Inches(5.4)
    card_tb = slide.shapes.add_textbox(Inches(0.75), Inches(1.55), Inches(3.2), card_tb_h)
    ctf = card_tb.text_frame
    ctf.word_wrap = True

    for idx, (item_title, item_desc) in enumerate(items):
        p_it = ctf.paragraphs[0] if idx == 0 else ctf.add_paragraph()
        p_it.text = item_title
        p_it.font.size = Pt(12)
        p_it.font.bold = True
        p_it.font.color.rgb = COLOR_ACCENT
        p_it.space_after = Pt(2)

        p_id = ctf.add_paragraph()
        if hyperlink_url and hyperlink_url in item_desc:
            parts = item_desc.split(hyperlink_url)
            if parts[0]:
                r0 = p_id.add_run()
                r0.text = parts[0]
                r0.font.size = Pt(10.5)
                r0.font.color.rgb = COLOR_TEXT_MAIN
            r_link = p_id.add_run()
            r_link.text = hyperlink_url
            r_link.font.size = Pt(10.5)
            r_link.font.bold = True
            r_link.font.color.rgb = RGBColor(37, 99, 235)
            r_link.font.underline = True
            r_link.hyperlink.address = hyperlink_url
            if len(parts) > 1 and parts[1]:
                r1 = p_id.add_run()
                r1.text = parts[1]
                r1.font.size = Pt(9.5)
                r1.font.color.rgb = COLOR_TEXT_MUTED
        else:
            p_id.text = item_desc
            p_id.font.size = Pt(10.5)
            p_id.font.color.rgb = COLOR_TEXT_MAIN
        p_id.space_after = Pt(6)

    # 클릭 시 웹페이지로 바로 이동하는 하이퍼링크 버튼
    if hyperlink_url:
        btn_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.75), Inches(6.38), Inches(3.2), Inches(0.62))
        btn_box.fill.solid()
        btn_box.fill.fore_color.rgb = RGBColor(30, 58, 138)
        btn_box.line.color.rgb = RGBColor(234, 88, 12)
        btn_box.line.width = Pt(1.5)
        btn_box.click_action.hyperlink.address = hyperlink_url
        btf = btn_box.text_frame
        btf.word_wrap = False
        bp = btf.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        brun = bp.add_run()
        brun.text = "🌐 웹페이지 바로 접속하기 (클릭)"
        brun.font.size = Pt(12)
        brun.font.bold = True
        brun.font.color.rgb = RGBColor(255, 255, 255)
        brun.hyperlink.address = hyperlink_url

    if os.path.exists(img_path):
        pic = slide.shapes.add_picture(img_path, Inches(4.3), Inches(1.45), width=Inches(8.5), height=Inches(5.7))
        if hyperlink_url:
            pic.click_action.hyperlink.address = hyperlink_url
    return slide


# ===========================================================================
# 1. 사용자 전용 PPTX 매뉴얼 생성 (그림 위주 & 직관적 심플 가이드)
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
    p1_tag.text = "SMART OVERTIME SYSTEM v1.54  |  초간편 사용자 매뉴얼"
    p1_tag.font.size = Pt(14)
    p1_tag.font.bold = True
    p1_tag.font.color.rgb = RGBColor(253, 186, 116)
    p1_tag.space_after = Pt(14)

    p1_title = tf1.add_paragraph()
    p1_title.text = "스마트 특근 관리 시스템\n그림만 보고 따라 하는 1분 가이드"
    p1_title.font.size = Pt(38)
    p1_title.font.bold = True
    p1_title.font.color.rgb = RGBColor(255, 255, 255)
    p1_title.space_after = Pt(20)

    p1_sub = tf1.add_paragraph()
    p1_sub.text = "접속 주소부터 사번 로그인, 버튼 용도별 상세 설명, 순차적 신청 및 [특근완료 확정]까지 그림 위주로 한눈에 정리했습니다."
    p1_sub.font.size = Pt(16)
    p1_sub.font.color.rgb = RGBColor(203, 213, 225)
    p1_sub.space_after = Pt(28)

    p1_auth = tf1.add_paragraph()
    p1_auth.text = "최신 릴리즈: v1.54  |  사내 PC 및 스마트폰 모바일(LTE/5G) 완벽 지원"
    p1_auth.font.size = Pt(13)
    p1_auth.font.color.rgb = RGBColor(148, 163, 184)

    # --- SLIDE 2: STEP 01 - 초기 접속 주소 및 방법 ---
    s2_data = [
        ("🌐 통합 웹 접속 공식 주소", "https://team-overtime-manager.onrender.com/\n(주소 클릭 시 시스템 웹페이지로 즉시 이동합니다.)"),
        ("🖥️ PC 브라우저 접속 (사내 / 자택)", "Chrome, Edge 등 웹 브라우저를 열고 위 주소를 입력하여 접속합니다. [즐겨찾기 ★] 등록 시 매번 원클릭으로 열 수 있습니다."),
        ("📱 스마트폰 모바일 접속 (LTE / 5G / Wi-Fi)", "스마트폰 기본 브라우저(Safari, Chrome, 삼성인터넷)로 접속하시면 모바일 최적화 화면이 자동으로 지원됩니다."),
        ("★ 모바일 홈 화면 바로가기 추가 팁", "스마트폰 브라우저 메뉴(⋮ 또는 공유)에서 [홈 화면에 추가]를 누르면 전용 앱 아이콘처럼 1초 만에 실행됩니다.")
    ]
    add_visual_slide(prs, "01", "스마트 특근 시스템 공식 웹 접속 안내", "PC, 노트북, 스마트폰 모바일 어디서나 공식 웹주소 하나로 간편하게 접속하세요.", s2_data, images["access"], hyperlink_url="https://team-overtime-manager.onrender.com/")

    # --- SLIDE 3: STEP 02 - 사원번호로 간편 로그인 방법 ---
    s3_data = [
        ("① 사원번호 6자리 숫자 입력", "비밀번호 없이 본인 사원번호 6자리를 입력창에 입력합니다."),
        ("② [입장하기 →] 원클릭 이동", "입력 후 파란색 버튼을 누르면 즉시 대시보드로 이동합니다."),
        ("③ 브라우저 자동 기억 기능 탑재", "다음 접속 시에는 사번이 기억되어 [입장하기]만 누르면 1초 만에 바로 진입합니다."),
        ("④ 미등록 사번 최초 1회 간편 등록", "처음 방문한 사번은 성명과 소속 부서를 선택하면 즉시 등록되어 입장합니다.")
    ]
    add_visual_slide(prs, "02", "사원번호로 초간편 1초 입장하기", "복잡한 비밀번호 없이 사번 6자리 입력만으로 빠르고 안전하게 입장합니다.", s3_data, images["login"])

    # --- SLIDE 4: STEP 03 - 화면 기능 및 선택버튼 상세 안내 ---
    s4_data = [
        ("🔄 [새로고침] 버튼", "관리자가 내 특근을 승인했는지 실시간 상태를 즉시 확인할 때 누릅니다."),
        ("📱 [모바일 접속 QR] 버튼", "스마트폰 카메라로 찍어 내 폰에서 편하게 신청하고 싶을 때 누릅니다."),
        ("💡 [건의사항 소통함] 버튼", "사번/이름 노출 없이 100% 무기명으로 회사/부서에 의견을 낼 때 누릅니다."),
        ("📋 [목록] vs 📅 [달력] 전환", "내 특근 내역을 표(리스트) 형태나 달력 형태로 바꿔보고 싶을 때 누릅니다."),
        ("🎯 [특근완료 확정] 버튼 (핵심!)", "실제 휴일 근무를 마친 후, 출근 완료 피드백을 관리자에게 보낼 때 누릅니다.")
    ]
    add_visual_slide(prs, "03", "화면 기능 및 주요 선택 버튼 상세 가이드", "어떤 상황에서 어떤 버튼을 사용하는지 직관적으로 확인하세요.", s4_data, images["buttons"])

    # --- SLIDE 5: STEP 04 - 특근 신청 순서 & 선택 버튼 상세 설명 ---
    s5_data = [
        ("1단계. 달력 일할 날짜 콕 누르기", "우측 달력에서 일할 토/일요일을 클릭하면 시작일과 종료일이 자동 세팅됩니다."),
        ("2단계. 특근 분류 버튼 선택 (★)", "● 일반휴일: 주말 나와서 정상 특근할 때\n● 대체근무: 평일 쉬고 주말 일할 때\n● 법정휴일: 설날/추석/공휴일 근무 시"),
        ("3단계. 항목 입력 (선택적 항목 예시)", "프로젝트(BT2601-L1), 장소(5층 제어실), 사유(설비 점검)를 예시 참고하여 기재합니다."),
        ("4단계. [💾 특근 신청 저장하기] 클릭", "초록색 저장 버튼을 누르면 신청 즉시 달력과 내역에 바로 등록 완료됩니다!")
    ]
    add_visual_slide(prs, "04", "특근 신청 순서 및 선택 버튼 상세 가이드", "달력 날짜 콕 ➔ 특근분류 선택 ➔ 예시 입력 ➔ 저장 원클릭으로 10초 만에 신청!", s5_data, images["apply"])

    # --- SLIDE 6: STEP 05 - 나의 특근 내역 관리 & [🎯 특근완료 확정] ---
    s6_data = [
        ("① 승인 확인: 초록색 [✓ 승인완료]", "관리자가 사전 승인하면 내 신청 내역에 [✓ 승인완료] 배지가 표시됩니다."),
        ("② 일정 변경 및 취소: [수정] / [취소]", "일정이나 사유를 고칠 때는 [수정], 특근이 취소되었을 때는 [취소]를 누릅니다."),
        ("③ ★ [🎯 특근완료 확정] (필독!)", "실제 휴일 근무를 마치고 퇴근 시 눌러 관리자에게 '출근 완료'를 최종 보고합니다."),
        ("④ 4단계 라이프사이클 안심 진행", "승인대기 ➔ 승인완료 ➔ 확정완료 ➔ 검토완료 순서로 투명하게 처리됩니다.")
    ]
    add_visual_slide(prs, "05", "나의 특근 내역 점검 및 [🎯 특근완료 확정]", "관리자 승인 확인부터 수정/취소, 실제 근무 후 완료 피드백까지 한눈에 관리!", s6_data, images["my"])

    # --- SLIDE 7: STEP 06 - 스마트폰 모바일 1분 퀵 가이드 ---
    s7_data = [
        ("① 카메라로 상단 QR 비추기", "별도 앱 설치 없이 기본 카메라로 QR 코드를 비추면 1초 만에 모바일 웹이 열립니다."),
        ("② 사번 6자리 터치 로그인", "스마트폰에서도 사번 자동 기억이 지원되어 매번 재입력할 필요 없습니다."),
        ("③ 달력 날짜 터치 & 신청 저장", "이동 중이나 자택에서도 손가락 터치 2번으로 10초 만에 특근을 신청합니다."),
        ("★ [홈 화면에 추가] (전용 앱 등록)", "브라우저 메뉴에서 홈 화면 추가 시 스마트폰 앱 아이콘으로 등록되어 바로 실행됩니다.")
    ]
    add_visual_slide(prs, "06", "스마트폰 모바일 1분 퀵 가이드", "앱 설치 없이 카메라 QR 스캔 ➔ 사번 로그인 ➔ 터치 신청으로 어디서나 간편하게!", s7_data, images["mobile"])

    # --- SLIDE 8: STEP 07 - 100% 무기명 건의사항 소통함 ---
    s8_data = [
        ("① 화면 상단 [💡 건의사항 소통함] 클릭", "언제든 상단 노란색 소통함 버튼을 누르면 무기명 팝업창이 열립니다."),
        ("② 100% 무기명 보장 (추적 불가)", "데이터베이스에 사번, 성명, IP 컬럼 자체가 존재하지 않아 완전 익명이 보장됩니다."),
        ("③ 자유로운 의견 및 건의 등록", "휴일 식사, 휴게실 냉난방, 근무환경 개선 등 바라는 점을 솔직하게 작성합니다."),
        ("④ 관리자 공식 답변 확인", "소통함 팝업 내에서 회사의 공식 검토 결과와 답변을 확인하실 수 있습니다.")
    ]
    add_visual_slide(prs, "07", "사원 전용 100% 무기명 건의사항 소통함", "사번 노출 걱정 제로! 자유로운 의견 제안과 관리자 공식 답변 확인을 지원합니다.", s8_data, images["suggestion"])

    # --- SLIDE 9: 사용자 핵심 5대 Q&A 총정리 ---
    s9 = prs.slides.add_slide(blank_layout)
    header_tb9 = s9.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.7), Inches(1.0))
    htf9 = header_tb9.text_frame
    htf9.word_wrap = True
    p_step9 = htf9.paragraphs[0]
    p_step9.text = "SUMMARY & Q&A  |  사용자 자주 묻는 질문 5문 5답 (v1.54)"
    p_step9.font.size = Pt(12)
    p_step9.font.bold = True
    p_step9.font.color.rgb = COLOR_ACCENT

    p_title9 = htf9.add_paragraph()
    p_title9.text = "자주 묻는 질문(Q&A) 핵심 요약"
    p_title9.font.size = Pt(22)
    p_title9.font.bold = True
    p_title9.font.color.rgb = COLOR_PRIMARY

    p_sub9 = htf9.add_paragraph()
    p_sub9.text = "궁금한 사항이 있으실 때는 언제든 상단 [📖 사용자 매뉴얼]을 다운로드하여 확인하세요."
    p_sub9.font.size = Pt(12.5)
    p_sub9.font.color.rgb = COLOR_TEXT_MUTED

    user_qa_items = [
        ("Q1. 사외 접속 주소와 모바일 접속 방법은 어떻게 되나요?",
         "관리자가 배포한 외부 HTTPS 보안 주소로 접속하거나, 화면 상단 [📱 모바일 접속 QR]을 스마트폰 카메라로 스캔하여 간편 접속합니다."),
        ("Q2. 재입장할 때 사번을 매번 다시 입력해야 하나요?",
         "아닙니다! 브라우저가 사번을 안전하게 기억하고 있으므로 [입장하기] 버튼만 누르면 1초 만에 바로 입장됩니다."),
        ("Q3. 실수로 같은 날짜에 중복 신청하면 어떻게 되나요?",
         "시스템이 동일 날짜 중복 신청을 사전에 감지하여 경고창과 함께 안전 차단하므로 안심하셔도 됩니다. 수정은 [수정] 버튼을 누르세요."),
        ("Q4. 실제 특근 후 [🎯 특근완료 확정] 버튼은 꼭 눌러야 하나요?",
         "네! 실제 휴일 근무를 마친 후 [확정]을 누르면 관리자에게 '출근 완료' 피드백이 전송되어 최종 정산에 정확히 반영됩니다."),
        ("Q5. 건의사항 소통함은 정말로 작성자가 누구인지 알 수 없나요?",
         "네, 100% 무기명입니다! 데이터베이스에 사번, 성명, IP 등 사용자 식별 컬럼이 아예 존재하지 않도록 설계되어 절대 추적되지 않습니다.")
    ]

    for i, (q, a) in enumerate(user_qa_items):
        qy = Inches(1.50 + i * 1.10)
        q_box = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), qy, Inches(11.733), Inches(0.95))
        q_box.fill.solid()
        q_box.fill.fore_color.rgb = RGBColor(248, 250, 252)
        q_box.line.color.rgb = RGBColor(226, 232, 240)

        q_tb = s9.shapes.add_textbox(Inches(1.0), qy + Inches(0.08), Inches(11.3), Inches(0.80))
        qtf = q_tb.text_frame
        qtf.word_wrap = True

        qp1 = qtf.paragraphs[0]
        qp1.text = q
        qp1.font.size = Pt(12.5)
        qp1.font.bold = True
        qp1.font.color.rgb = COLOR_PRIMARY
        qp1.space_after = Pt(2)

        qp2 = qtf.add_paragraph()
        qp2.text = f"👉 {a}"
        qp2.font.size = Pt(11)
        qp2.font.color.rgb = COLOR_ACCENT

    return prs


# ===========================================================================
# 2. 관리자 전용 PPTX 매뉴얼 생성 (그림 위주 & 직관적 관리 가이드)
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
    p1_tag.text = "SMART OVERTIME SYSTEM v1.54  |  관리자 및 운영자 전용 가이드"
    p1_tag.font.size = Pt(14)
    p1_tag.font.bold = True
    p1_tag.font.color.rgb = RGBColor(253, 186, 116)
    p1_tag.space_after = Pt(14)

    p1_title = tf1.add_paragraph()
    p1_title.text = "스마트 특근 관리 시스템\n관리자 모드 운영 매뉴얼 (v1.54)"
    p1_title.font.size = Pt(38)
    p1_title.font.bold = True
    p1_title.font.color.rgb = RGBColor(255, 255, 255)
    p1_title.space_after = Pt(20)

    p1_sub = tf1.add_paragraph()
    p1_sub.text = "캘린더 원클릭 승인, 팀원 대리 신청 및 비밀 보너스, 실특근 자동 정산표, 27번 표준 엑셀 내보내기까지 완벽 지원합니다."
    p1_sub.font.size = Pt(16)
    p1_sub.font.color.rgb = RGBColor(203, 213, 225)
    p1_sub.space_after = Pt(28)

    p1_auth = tf1.add_paragraph()
    p1_auth.text = "배포 버전: v1.54  |  총괄 슈퍼관리자 및 부서 팀관리자 전용"
    p1_auth.font.size = Pt(13)
    p1_auth.font.color.rgb = RGBColor(148, 163, 184)

    # --- SLIDE 2: 관리자 캘린더 & 원클릭 일괄 승인 ---
    s2_data = [
        ("① 팀원 특근 일정 한눈에 파악", "달력 칸마다 소속 팀원의 특근 구분이 색상 뱃지(초록/파랑/보라)로 즉시 표시됩니다."),
        ("② [당일 전원 일괄 확인] (원클릭!)", "날짜 클릭 후 우측 명단 패널에서 초록색 버튼 하나로 당일 신청자 전원을 즉시 일괄 승인합니다."),
        ("③ 실시간 개별 승인 토글 스위치", "작업자별 [승인] / [취소] 스위치로 개별 제어가 가능합니다."),
        ("④ 색상 범례 상시 표시", "일반휴일, 대체근무, 법정휴일, 대휴사용을 구분하는 직관적 색상 범례를 상시 제공합니다.")
    ]
    add_visual_slide(prs, "01", "월간 캘린더 및 원클릭 일괄 승인", "팀원 일정 색상 파악, 일자별 작업자 패널 및 당일 전원 일괄 승인을 지원합니다.", s2_data, images["cal"])

    # --- SLIDE 3: 팀원 명부 및 관리자 권한 관리 ---
    s3_data = [
        ("① 슈퍼관리자(총괄) 영구 보호", "총괄관리자 계정은 권한 박탈 및 삭제가 원천 차단되어 시스템이 안전하게 보호됩니다."),
        ("② 팀관리자 권한 원클릭 토글", "일반 팀원에게 '팀관리자' 권한을 토글 버튼으로 부여하여 본인 부서 특근만 승인하도록 분리합니다."),
        ("③ 소속팀 변경 및 부서 관리", "사원의 소속팀을 간편 변경하고, 새로운 부서/팀을 등록 및 삭제할 수 있습니다."),
        ("④ 엑셀 사원 명부 가져오기/내보내기", "전체 팀원 명부를 엑셀 파일로 백업하거나 대량 등록할 수 있습니다.")
    ]
    add_visual_slide(prs, "02", "팀원 명부 관리 및 팀관리자 권한 지정", "부서별 권한 분리 및 사원 명부 관리, 슈퍼관리자 보호 기능을 제공합니다.", s3_data, images["user"])

    # --- SLIDE 4: 실특근 자동 산정 및 정산표 ---
    s4_data = [
        ("① 4대 휴일수 완벽 분리 집계", "일반휴일, 대체근무, 법정휴일, 대휴사용을 각각 독립 컬럼으로 정확히 합산합니다."),
        ("② 최종 실특근 1초 자동 산정", "일반휴일에서 대휴 사용일수를 자동 차감하여 최종 실특근 일수를 1초 만에 자동 산출합니다."),
        ("③ 부서별 / 개인별 다차원 필터링", "기간 선택, 소속팀 선택으로 원하는 단위의 정산 통계를 즉시 확인합니다."),
        ("④ 정산 엑셀(.xlsx) 원클릭 다운로드", "정산 결과를 공식 엑셀 양식으로 즉시 내려받아 경영지원팀에 제출합니다.")
    ]
    add_visual_slide(prs, "03", "실특근 자동 산정 및 정산표 관리", "휴일별 독립 집계 및 최종 실특근 자동 계산, 엑셀 다운로드를 지원합니다.", s4_data, images["settlement"])

    # --- SLIDE 5: 엑셀 내보내기 27번 표준 서식 ---
    s5_data = [
        ("① 날짜순 차례차례 펼침 전개", "근무 일자 순서대로 데이터가 정리되며 대휴사용일이 같은 행에 직관적으로 연결됩니다."),
        ("② 순수 숫자 '1' 표기 (수식 자동화)", "문자 '일' 없이 순수 숫자 1만 기재되어 마우스 드래그 합계(SUM)가 바로 계산됩니다."),
        ("③ 3대 시트 분리 (원장/정산/보너스)", "Sheet 1(전체내역), Sheet 2(개인별정산), Sheet 3(보너스명부)으로 완벽 분리 제공됩니다."),
        ("④ 인사/경영지원팀 제출용 최적화", "추가 가공 없이 보고서 및 결재 문서에 즉시 첨부 가능한 표준 서식입니다.")
    ]
    add_visual_slide(prs, "04", "특근정보 엑셀(.xlsx) 27번 표준 양식", "마우스 드래그 수식 계산 및 다중 시트 완비 공식 엑셀 양식을 내보냅니다.", s5_data, images["excel"])

    # --- SLIDE 6: 무기명 소통함 운영 및 관리자 공식 답변 ---
    s6_data = [
        ("① 사원 무기명 의견 실시간 열람", "사원이 등록한 소통함 의견을 실시간 확인합니다. (사번/성명 일체 미표시)"),
        ("② 관리자 공식 답변 작성 및 게시", "검토 결과를 [답변완료] 상태로 등록하여 전 사원이 소통함에서 열람하도록 지원합니다."),
        ("③ 건전한 소통 문화 조성", "익명성을 기반으로 한 현장의 실질적 애로사항과 근무환경 개선 요구를 선제 청취합니다."),
        ("④ 완료된 의견 정리 및 보관", "답변 완료된 건의사항은 이력으로 안전하게 보존 관리됩니다.")
    ]
    add_visual_slide(prs, "05", "무기명 건의사항 소통함 운영 및 답변", "완전 무기명 의견 청취 및 관리자 공식 검토 답변 작성을 지원합니다.", s6_data, images["suggestion"])

    return prs


# ----------------- 실행 및 저장 통합 함수 -----------------

def create_manual():
    print("[1/3] Generating ultra-friendly visual UI mockups with speech bubbles & leader lines...")
    images = {
        "access": create_mockup_access_methods(),
        "login": create_mockup_login(),
        "buttons": create_mockup_dashboard_buttons(),
        "apply": create_mockup_calendar_apply(),
        "my": create_mockup_my_records_guide(),
        "mobile": create_mockup_mobile_flow(),
        "suggestion": create_mockup_suggestion_box(),
        "cal": create_mockup_admin_calendar(),
        "user": create_mockup_user_mgmt(),
        "settlement": create_mockup_settlement(),
        "excel": create_mockup_excel_27()
    }

    print("[2/3] Building User Manual, Admin Manual, and Integrated System Manual...")
    prs_user = build_user_presentation(images)
    prs_admin = build_admin_presentation(images)

    # 통합 매뉴얼 (사용자 가이드 9슬라이드 + 관리자 전환 간지 + 관리자 5슬라이드)
    prs_system = Presentation()
    prs_system.slide_width = Inches(13.333)
    prs_system.slide_height = Inches(7.5)
    blank_layout = prs_system.slide_layouts[6]

    # 시스템 종합 표지
    s_cov = prs_system.slides.add_slide(blank_layout)
    bg_cov = s_cov.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg_cov.fill.solid()
    bg_cov.fill.fore_color.rgb = COLOR_PRIMARY
    bg_cov.line.color.rgb = COLOR_PRIMARY

    tb_cov = s_cov.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.3), Inches(4.5))
    tf_cov = tb_cov.text_frame
    tf_cov.word_wrap = True

    p_cov_tag = tf_cov.paragraphs[0]
    p_cov_tag.text = "SMART OVERTIME SYSTEM v1.54  |  시스템 종합 매뉴얼"
    p_cov_tag.font.size = Pt(14)
    p_cov_tag.font.bold = True
    p_cov_tag.font.color.rgb = RGBColor(253, 186, 116)
    p_cov_tag.space_after = Pt(14)

    p_cov_tit = tf_cov.add_paragraph()
    p_cov_tit.text = "스마트 특근 관리 시스템\n사용자 및 관리자 통합 운영 가이드"
    p_cov_tit.font.size = Pt(38)
    p_cov_tit.font.bold = True
    p_cov_tit.font.color.rgb = RGBColor(255, 255, 255)
    p_cov_tit.space_after = Pt(20)

    p_cov_sub = tf_cov.add_paragraph()
    p_cov_sub.text = "일반 팀원을 위한 초간편 신청 가이드부터 총괄/팀관리자를 위한 승인, 정산, 엑셀 원장 관리까지 그림 위주로 완벽 수록했습니다."
    p_cov_sub.font.size = Pt(16)
    p_cov_sub.font.color.rgb = RGBColor(203, 213, 225)
    p_cov_sub.space_after = Pt(28)

    p_cov_auth = tf_cov.add_paragraph()
    p_cov_auth.text = "버전: v1.54 최신  |  전 사원 및 부서 관리자 공용"
    p_cov_auth.font.size = Pt(13)
    p_cov_auth.font.color.rgb = RGBColor(148, 163, 184)

    # 사용자 가이드 슬라이드 7개 추가
    add_visual_slide(prs_system, "01", "스마트 특근 시스템 공식 웹 접속 안내", "PC, 노트북, 스마트폰 모바일 어디서나 공식 웹주소 하나로 간편하게 접속하세요.", [
        ("🌐 통합 웹 접속 공식 주소", "https://team-overtime-manager.onrender.com/\n(주소 클릭 시 시스템 웹페이지로 즉시 이동합니다.)"),
        ("🖥️ PC 브라우저 접속 (사내 / 자택)", "Chrome, Edge 등 웹 브라우저를 열고 위 주소를 입력하여 접속합니다. [즐겨찾기 ★] 등록 시 매번 원클릭으로 열 수 있습니다."),
        ("📱 스마트폰 모바일 접속 (LTE / 5G / Wi-Fi)", "스마트폰 기본 브라우저(Safari, Chrome, 삼성인터넷)로 접속하시면 모바일 최적화 화면이 자동으로 지원됩니다."),
        ("★ 모바일 홈 화면 바로가기 추가 팁", "스마트폰 브라우저 메뉴(⋮ 또는 공유)에서 [홈 화면에 추가]를 누르면 전용 앱 아이콘처럼 1초 만에 실행됩니다.")
    ], images["access"], hyperlink_url="https://team-overtime-manager.onrender.com/")

    add_visual_slide(prs_system, "02", "사원번호로 초간편 1초 입장하기", "복잡한 비밀번호 없이 사번 6자리 입력만으로 빠르고 안전하게 입장합니다.", [
        ("① 사원번호 6자리 숫자 입력", "비밀번호 없이 본인 사원번호 6자리를 입력창에 입력합니다."),
        ("② [입장하기 →] 원클릭 이동", "입력 후 파란색 버튼을 누르면 즉시 대시보드로 이동합니다."),
        ("③ 브라우저 자동 기억 기능 탑재", "다음 접속 시에는 사번이 기억되어 [입장하기]만 누르면 1초 만에 바로 진입합니다."),
        ("④ 미등록 사번 최초 1회 간편 등록", "처음 방문한 사번은 성명과 소속 부서를 선택하면 즉시 등록되어 입장합니다.")
    ], images["login"])

    add_visual_slide(prs_system, "03", "화면 기능 및 주요 선택 버튼 상세 가이드", "어떤 상황에서 어떤 버튼을 사용하는지 직관적으로 확인하세요.", [
        ("🔄 [새로고침] 버튼", "관리자가 내 특근을 승인했는지 실시간 상태를 즉시 확인할 때 누릅니다."),
        ("📱 [모바일 접속 QR] 버튼", "스마트폰 카메라로 찍어 내 폰에서 편하게 신청하고 싶을 때 누릅니다."),
        ("💡 [건의사항 소통함] 버튼", "사번/이름 노출 없이 100% 무기명으로 회사/부서에 의견을 낼 때 누릅니다."),
        ("📋 [목록] vs 📅 [달력] 전환", "내 특근 내역을 표(리스트) 형태나 달력 형태로 바꿔보고 싶을 때 누릅니다."),
        ("🎯 [특근완료 확정] 버튼 (핵심!)", "실제 휴일 근무를 마친 후, 출근 완료 피드백을 관리자에게 보낼 때 누릅니다.")
    ], images["buttons"])

    add_visual_slide(prs_system, "04", "특근 신청 순서 및 선택 버튼 상세 가이드", "달력 날짜 콕 ➔ 특근분류 선택 ➔ 예시 입력 ➔ 저장 원클릭으로 10초 만에 신청!", [
        ("1단계. 달력 일할 날짜 콕 누르기", "우측 달력에서 일할 토/일요일을 클릭하면 시작일과 종료일이 자동 세팅됩니다."),
        ("2단계. 특근 분류 버튼 선택 (★)", "● 일반휴일: 주말 나와서 정상 특근할 때\n● 대체근무: 평일 쉬고 주말 일할 때\n● 법정휴일: 설날/추석/공휴일 근무 시"),
        ("3단계. 항목 입력 (선택적 항목 예시)", "프로젝트(BT2601-L1), 장소(5층 제어실), 사유(설비 점검)를 예시 참고하여 기재합니다."),
        ("4단계. [💾 특근 신청 저장하기] 클릭", "초록색 저장 버튼을 누르면 신청 즉시 달력과 내역에 바로 등록 완료됩니다!")
    ], images["apply"])

    add_visual_slide(prs_system, "05", "나의 특근 내역 점검 및 [🎯 특근완료 확정]", "관리자 승인 확인부터 수정/취소, 실제 근무 후 완료 피드백까지 한눈에 관리!", [
        ("① 승인 확인: 초록색 [✓ 승인완료]", "관리자가 사전 승인하면 내 신청 내역에 [✓ 승인완료] 배지가 표시됩니다."),
        ("② 일정 변경 및 취소: [수정] / [취소]", "일정이나 사유를 고칠 때는 [수정], 특근이 취소되었을 때는 [취소]를 누릅니다."),
        ("③ ★ [🎯 특근완료 확정] (필독!)", "실제 휴일 근무를 마치고 퇴근 시 눌러 관리자에게 '출근 완료'를 최종 보고합니다."),
        ("④ 4단계 라이프사이클 안심 진행", "승인대기 ➔ 승인완료 ➔ 확정완료 ➔ 검토완료 순서로 투명하게 처리됩니다.")
    ], images["my"])

    add_visual_slide(prs_system, "06", "스마트폰 모바일 1분 퀵 가이드", "앱 설치 없이 카메라 QR 스캔 ➔ 사번 로그인 ➔ 터치 신청으로 어디서나 간편하게!", [
        ("① 카메라로 상단 QR 비추기", "별도 앱 설치 없이 기본 카메라로 QR 코드를 비추면 1초 만에 모바일 웹이 열립니다."),
        ("② 사번 6자리 터치 로그인", "스마트폰에서도 사번 자동 기억이 지원되어 매번 재입력할 필요 없습니다."),
        ("③ 달력 날짜 터치 & 신청 저장", "이동 중이나 자택에서도 손가락 터치 2번으로 10초 만에 특근을 신청합니다."),
        ("★ [홈 화면에 추가] (전용 앱 등록)", "브라우저 메뉴에서 홈 화면 추가 시 스마트폰 앱 아이콘으로 등록되어 바로 실행됩니다.")
    ], images["mobile"])

    add_visual_slide(prs_system, "07", "사원 전용 100% 무기명 건의사항 소통함", "사번 노출 걱정 제로! 자유로운 의견 제안과 관리자 공식 답변 확인을 지원합니다.", [
        ("① 화면 상단 [💡 건의사항 소통함] 클릭", "언제든 상단 노란색 소통함 버튼을 누르면 무기명 팝업창이 열립니다."),
        ("② 100% 무기명 보장 (추적 불가)", "데이터베이스에 사번, 성명, IP 컬럼 자체가 존재하지 않아 완전 익명이 보장됩니다."),
        ("③ 자유로운 의견 및 건의 등록", "휴일 식사, 휴게실 냉난방, 근무환경 개선 등 바라는 점을 솔직하게 작성합니다."),
        ("④ 관리자 공식 답변 확인", "소통함 팝업 내에서 회사의 공식 검토 결과와 답변을 확인하실 수 있습니다.")
    ], images["suggestion"])

    # 중간 관리자 모드 간지
    s_div = prs_system.slides.add_slide(blank_layout)
    bg_div = s_div.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg_div.fill.solid()
    bg_div.fill.fore_color.rgb = COLOR_PRIMARY
    bg_div.line.color.rgb = COLOR_PRIMARY
    tb_div = s_div.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.3), Inches(3.0))
    tf_div = tb_div.text_frame
    p_div_tag = tf_div.paragraphs[0]
    p_div_tag.text = "PART 2  |  ADMINISTRATION GUIDE"
    p_div_tag.font.size = Pt(14)
    p_div_tag.font.bold = True
    p_div_tag.font.color.rgb = RGBColor(253, 186, 116)
    p_div_tit = tf_div.add_paragraph()
    p_div_tit.text = "[관리자 모드] 운영 및 정산 매뉴얼"
    p_div_tit.font.size = Pt(36)
    p_div_tit.font.bold = True
    p_div_tit.font.color.rgb = RGBColor(255, 255, 255)

    # 관리자 가이드 슬라이드 5개 추가
    add_visual_slide(prs_system, "08", "[관리자] 월간 캘린더 및 원클릭 일괄 승인", "팀원 일정 색상 파악, 일자별 작업자 패널 및 당일 전원 일괄 승인을 지원합니다.", [
        ("① 팀원 특근 일정 한눈에 파악", "달력 칸마다 소속 팀원의 특근 구분이 색상 뱃지(초록/파랑/보라)로 즉시 표시됩니다."),
        ("② [당일 전원 일괄 확인] (원클릭!)", "날짜 클릭 후 우측 명단 패널에서 초록색 버튼 하나로 당일 신청자 전원을 즉시 일괄 승인합니다."),
        ("③ 실시간 개별 승인 토글 스위치", "작업자별 [승인] / [취소] 스위치로 개별 제어가 가능합니다."),
        ("④ 색상 범례 상시 표시", "일반휴일, 대체근무, 법정휴일, 대휴사용을 구분하는 직관적 색상 범례를 상시 제공합니다.")
    ], images["cal"])

    add_visual_slide(prs_system, "09", "[관리자] 팀원 명부 관리 및 팀관리자 권한 지정", "부서별 권한 분리 및 사원 명부 관리, 슈퍼관리자 보호 기능을 제공합니다.", [
        ("① 슈퍼관리자(총괄) 영구 보호", "총괄관리자 계정은 권한 박탈 및 삭제가 원천 차단되어 시스템이 안전하게 보호됩니다."),
        ("② 팀관리자 권한 원클릭 토글", "일반 팀원에게 '팀관리자' 권한을 토글 버튼으로 부여하여 본인 부서 특근만 승인하도록 분리합니다."),
        ("③ 소속팀 변경 및 부서 관리", "사원의 소속팀을 간편 변경하고, 새로운 부서/팀을 등록 및 삭제할 수 있습니다."),
        ("④ 엑셀 사원 명부 백업/등록", "전체 팀원 명부를 엑셀 파일로 백업하거나 대량 등록할 수 있습니다.")
    ], images["user"])

    add_visual_slide(prs_system, "10", "[관리자] 실특근 자동 산정 및 정산표 관리", "휴일별 독립 집계 및 최종 실특근 자동 계산, 엑셀 다운로드를 지원합니다.", [
        ("① 4대 휴일수 완벽 분리 집계", "일반휴일, 대체근무, 법정휴일, 대휴사용을 각각 독립 컬럼으로 정확히 합산합니다."),
        ("② 최종 실특근 1초 자동 산정", "일반휴일에서 대휴 사용일수를 자동 차감하여 최종 실특근 일수를 1초 만에 자동 산출합니다."),
        ("③ 부서별 / 개인별 다차원 필터링", "기간 선택, 소속팀 선택으로 원하는 단위의 정산 통계를 즉시 확인합니다."),
        ("④ 정산 엑셀(.xlsx) 원클릭 다운로드", "정산 결과를 공식 엑셀 양식으로 즉시 내려받아 경영지원팀에 제출합니다.")
    ], images["settlement"])

    add_visual_slide(prs_system, "11", "[관리자] 특근정보 엑셀(.xlsx) 27번 표준 양식", "마우스 드래그 수식 계산 및 다중 시트 완비 공식 엑셀 양식을 내보냅니다.", [
        ("① 날짜순 차례차례 펼침 전개", "근무 일자 순서대로 데이터가 정리되며 대휴사용일이 같은 행에 직관적으로 연결됩니다."),
        ("② 순수 숫자 '1' 표기 (수식 자동화)", "문자 '일' 없이 순수 숫자 1만 기재되어 마우스 드래그 합계(SUM)가 바로 계산됩니다."),
        ("③ 3대 시트 분리 (원장/정산/보너스)", "Sheet 1(전체내역), Sheet 2(개인별정산), Sheet 3(보너스명부)으로 완벽 분리 제공됩니다."),
        ("④ 인사/경영지원팀 제출용 최적화", "추가 가공 없이 보고서 및 결재 문서에 즉시 첨부 가능한 표준 서식입니다.")
    ], images["excel"])

    add_visual_slide(prs_system, "12", "[관리자] 무기명 건의사항 소통함 운영 및 답변", "완전 무기명 의견 청취 및 관리자 공식 검토 답변 작성을 지원합니다.", [
        ("① 사원 무기명 의견 실시간 열람", "사원이 등록한 소통함 의견을 실시간 확인합니다. (사번/성명 일체 미표시)"),
        ("② 관리자 공식 답변 작성 및 게시", "검토 결과를 [답변완료] 상태로 등록하여 전 사원이 소통함에서 열람하도록 지원합니다."),
        ("③ 건전한 소통 문화 조성", "익명성을 기반으로 한 현장의 실질적 애로사항과 근무환경 개선 요구를 선제 청취합니다."),
        ("④ 완료된 의견 정리 및 보관", "답변 완료된 건의사항은 이력으로 안전하게 보존 관리됩니다.")
    ], images["suggestion"])

    print("[3/3] Saving PPTX presentations to downloads directory...")
    prs_user.save(str(USER_PPTX_PATH))
    print(f"✓ Saved User Manual: {USER_PPTX_PATH}")

    prs_admin.save(str(ADMIN_PPTX_PATH))
    print(f"✓ Saved Admin Manual: {ADMIN_PPTX_PATH}")

    prs_system.save(str(PPTX_PATH))
    print(f"✓ Saved System Manual: {PPTX_PATH}")

    def copy_and_verify(src, dst):
        try:
            shutil.copy2(str(src), str(dst))
            if dst.exists() and dst.stat().st_size > 0:
                print(f"✓ Copied & Verified: {dst.name} ({dst.stat().st_size:,} bytes) -> {dst}")
                return True
            else:
                raise RuntimeError(f"Failed to verify copied file: {dst}")
        except Exception as e:
            print(f"⚠️ [WARNING] Failed to copy '{src.name}' to '{dst.name}': {e}")
            return False

    root_dir = Path(__file__).resolve().parent
    static_dl = root_dir / "static" / "downloads"
    static_dl.mkdir(parents=True, exist_ok=True)

    copy_and_verify(USER_PPTX_PATH, root_dir / "Overtime_User_Manual.pptx")
    copy_and_verify(ADMIN_PPTX_PATH, root_dir / "Overtime_Admin_Manual.pptx")
    copy_and_verify(PPTX_PATH, root_dir / "Overtime_System_Manual.pptx")
    print(f"✓ Synchronized PPTX manuals with ROOT main directory: {root_dir}")

    copy_and_verify(USER_PPTX_PATH, static_dl / "Overtime_User_Manual.pptx")
    copy_and_verify(ADMIN_PPTX_PATH, static_dl / "Overtime_Admin_Manual.pptx")
    copy_and_verify(PPTX_PATH, static_dl / "Overtime_System_Manual.pptx")
    print(f"✓ Synchronized PPTX manuals with static download path: {static_dl}")
    print("🎉 All PPT Manuals successfully regenerated with rich pictures and intuitive guides!")

if __name__ == "__main__":
    create_manual()
