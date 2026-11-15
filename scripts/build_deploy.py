#!/usr/bin/env python3
"""
PriceGuard 배포 폴더(deploy/) 생성·검증 스크립트

사용법 (프로젝트 루트에서):
    python scripts/build_deploy.py          # deploy/ 를 처음부터 다시 만들고 검증
    python scripts/build_deploy.py --check  # 만들지 않고 현재 deploy/ 만 검증

원칙
- 아래 FILES 목록(허용 목록)에 있는 파일만 deploy/ 로 복사한다.
- 기존 프로젝트 파일은 읽기만 한다. 지우거나 바꾸는 것은 deploy/ 폴더뿐이다.
- 다시 만들 때는 임시 폴더에 먼저 만든 뒤 검증이 끝나면 deploy/ 와 바꾼다.
  (검증에 실패하면 기존 deploy/ 는 그대로 남는다.)
- 외부 패키지 없이 파이썬 표준 라이브러리만 쓴다.
"""
import hashlib
import json
import re
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEPLOY = ROOT / "deploy"
TMP = ROOT / ".deploy-build-tmp"
OLD = ROOT / ".deploy-old"

# 실제 서비스 실행에 필요한 파일 (허용 목록)
FILES = [
    "index.html",                                   # 화면 구조
    "style.css",                                    # 디자인
    "app.js",                                       # 화면 동작·업로드·다운로드
    "clean.js",                                     # 정제 규칙 실행 엔진
    "rules/병의원수가표.json",                       # 승인된 정제 규칙 (app.js 가 fetch)
    "lib/xlsx.full.min.js",                         # SheetJS (엑셀 읽기·쓰기)
    "lib/LICENSE",                                  # SheetJS Apache-2.0 라이선스 (재배포 고지)
    "lib/fonts/PretendardVariable.woff2",           # Pretendard 글꼴 (style.css 가 참조)
    "lib/fonts/LICENSE-Pretendard.txt",             # Pretendard SIL OFL 라이선스 (재배포 고지)
    "assets/brand/priceguard-logo-display.png",     # 화면 표시용 로고
    "assets/brand/favicon.png",                     # 파비콘
]

# deploy/ 안에서만 만드는 설정 파일 (Vercel 정적 사이트, 빌드 없음)
VERCEL_JSON = {
    "headers": [
        {
            "source": "/(.*)",
            "headers": [
                {"key": "X-Content-Type-Options", "value": "nosniff"},
                {"key": "X-Frame-Options", "value": "DENY"},
                {"key": "Referrer-Policy", "value": "no-referrer"},
                {"key": "Permissions-Policy", "value": "camera=(), microphone=(), geolocation=()"},
            ],
        }
    ]
}
GENERATED = {"vercel.json": json.dumps(VERCEL_JSON, ensure_ascii=False, indent=2) + "\n"}

# 절대 deploy/ 에 들어가면 안 되는 것
FORBIDDEN_TOP = {"data", "tools", "tests", "진단", "design", "docs", ".claude", "scripts", "deploy", ".git"}
FORBIDDEN_NAMES = {"CLAUDE.md", "로고.png", "priceguard-logo.png", ".env"}
FORBIDDEN_EXT = {".xlsx", ".xls", ".xlsm", ".csv", ".md", ".env", ".pem", ".key", ".log", ".py"}

# 민감 정보로 의심되는 문자열 (발견 시 보고만 하고 자동 삭제하지 않는다)
SENSITIVE = [
    ("로컬 사용자 경로", re.compile(r"[A-Za-z]:[\\/]+(?:Users|Documents and Settings)[\\/]|/home/[a-z]|AppData[\\/]", re.I)),
    ("비밀 값 할당", re.compile(r"(?:api[_-]?key|secret|access[_-]?token|auth[_-]?token|password|passwd)\s*[:=]\s*['\"][^'\"]{4,}", re.I)),
    ("개인 키", re.compile(r"BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY")),
    ("이메일 주소", re.compile(r"[\w.+-]+@[\w-]+\.(?:com|net|org|kr|co\.kr|io)\b")),
    ("휴대전화 번호", re.compile(r"\b01[016789]-?\d{3,4}-?\d{4}\b")),
    ("주민등록번호 형식", re.compile(r"\b\d{6}-[1-4]\d{6}\b")),
    ("내부 정답 자료 언급", re.compile(r"tools/answer|_정답\.csv|만든방법")),
]


def sha(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def refs_in(target: Path):
    """index.html / style.css / app.js 가 참조하는 로컬 파일 경로"""
    found = set()
    html = (target / "index.html").read_text(encoding="utf-8")
    found.update(m for m in re.findall(r'(?:src|href)="([^"#:]+)"', html))
    css = (target / "style.css").read_text(encoding="utf-8")
    found.update(m for m in re.findall(r'url\(["\']?([^)"\']+)', css) if not m.startswith("data:"))
    js = (target / "app.js").read_text(encoding="utf-8")
    found.update(re.findall(r"RULES_URL\s*=\s*'([^']+)'", js))
    return sorted(found)


def verify(target: Path, compare_source: bool = True):
    problems, warnings = [], []
    if not target.is_dir():
        return [f"{target.name}/ 폴더가 없습니다."], warnings
    present = sorted(p.relative_to(target).as_posix() for p in target.rglob("*") if p.is_file())
    expected = sorted(FILES + list(GENERATED))
    for f in expected:
        if f not in present:
            problems.append(f"필수 파일 없음: {f}")
    for f in present:
        if f not in expected:
            problems.append(f"허용 목록에 없는 파일: {f}")
        parts = f.split("/")
        if parts[0] in FORBIDDEN_TOP or parts[-1] in FORBIDDEN_NAMES or Path(f).suffix.lower() in FORBIDDEN_EXT:
            problems.append(f"금지된 파일: {f}")
    for ref in refs_in(target):
        if not (target / ref).is_file():
            problems.append(f"참조했지만 없는 파일: {ref}")
    if compare_source:
        for f in FILES:
            if (target / f).is_file() and sha(target / f) != sha(ROOT / f):
                problems.append(f"원본과 내용이 다름(다시 만들기 필요): {f}")
    for f in present:
        p = target / f
        if p.suffix.lower() in {".png", ".woff2"}:
            continue
        text = p.read_text(encoding="utf-8", errors="replace")
        for label, rx in SENSITIVE:
            for m in rx.finditer(text):
                s = max(0, m.start() - 30)
                warnings.append(f"[{label}] {f}: …{text[s:m.end() + 30]!r}…")
    return problems, warnings


def build():
    # 안전장치: deploy/ 와 임시 폴더가 정확히 프로젝트 바로 아래인지 확인
    for d in (DEPLOY, TMP, OLD):
        if d.parent != ROOT or d.name not in ("deploy", ".deploy-build-tmp", ".deploy-old"):
            sys.exit(f"중단: 예상하지 못한 경로 {d}")
        if d.is_symlink():
            sys.exit(f"중단: {d} 가 링크입니다.")
    missing = [f for f in FILES if not (ROOT / f).is_file()]
    if missing:
        sys.exit("중단: 원본 파일 없음 → " + ", ".join(missing))
    if TMP.exists():
        shutil.rmtree(TMP)
    for f in FILES:
        dst = TMP / f
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ROOT / f, dst)
    for name, content in GENERATED.items():
        (TMP / name).write_text(content, encoding="utf-8", newline="\n")
    problems, warnings = verify(TMP)
    if problems:
        shutil.rmtree(TMP)
        report(problems, warnings, title="새 배포 폴더 검증 실패 — 기존 deploy/ 는 그대로 둡니다")
        sys.exit(1)
    # 기존 deploy/ 는 지우기 전에 먼저 이름을 바꿔 둔다.
    # 다른 프로그램(로컬 서버, 탐색기, 터미널)이 사용 중이면 여기서 멈추고 기존 deploy/ 를 그대로 둔다.
    if OLD.exists():
        shutil.rmtree(OLD)
    if DEPLOY.exists():
        try:
            DEPLOY.rename(OLD)
        except OSError:
            shutil.rmtree(TMP, ignore_errors=True)
            sys.exit("중단: deploy/ 폴더를 다른 프로그램(로컬 서버, 탐색기, 터미널 등)이 사용 중입니다. "
                     "그 프로그램을 끈 뒤 다시 실행하세요. 기존 deploy/ 는 바뀌지 않았습니다.")
    TMP.rename(DEPLOY)
    shutil.rmtree(OLD, ignore_errors=True)
    return problems, warnings


def report(problems, warnings, title):
    print(f"== {title} ==")
    files = sorted(p.relative_to(DEPLOY).as_posix() for p in DEPLOY.rglob("*") if p.is_file()) if DEPLOY.exists() else []
    total = sum((DEPLOY / f).stat().st_size for f in files)
    for f in files:
        print(f"  {(DEPLOY / f).stat().st_size:>10,}  {f}")
    print(f"  파일 {len(files)}개, 합계 {total:,} bytes")
    print(f"문제: {len(problems)}건")
    for p in problems:
        print("  - " + p)
    print(f"민감 정보 의심 문자열: {len(warnings)}건 (자동 삭제하지 않음, 내용 확인 필요)")
    for w in warnings:
        print("  - " + w)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    if "--check" in sys.argv:
        problems, warnings = verify(DEPLOY)
        report(problems, warnings, title="deploy/ 검증 (--check)")
        sys.exit(1 if problems else 0)
    problems, warnings = build()
    report(problems, warnings, title="deploy/ 다시 만들기 완료")
