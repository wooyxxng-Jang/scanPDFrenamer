# v.1.0 / 2025-08-13

# =============== 프로그램 기본 설정 (수정 금지!!) ===============
import os
import sys
import shutil
import re
import logging
import json
import subprocess
import ctypes
import tempfile
import pandas as pd
from PIL import Image
from pdf2image import convert_from_path
import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
import threading
from pathlib import Path

try:
    base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) \
               else os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)
except Exception as e:
    logging.warning(f"작업 디렉토리 변경 실패(무시 가능): {e}")

# 경로 유틸
def resource_path(*relative_parts) -> str:
    if getattr(sys, 'frozen', False):
        base = Path(getattr(sys, '_MEIPASS', os.path.dirname(sys.executable)))
    else:
        base = Path(os.path.dirname(os.path.abspath(__file__)))
    return str(base.joinpath(*relative_parts))

def external_path(*parts) -> str:
    #실행 exe(또는 소스) 폴더 기준 외부 경로
    base = Path(os.path.dirname(sys.executable)) if getattr(sys, 'frozen', False) \
           else Path(os.path.dirname(os.path.abspath(__file__)))
    return str(base.joinpath(*parts))

# 기본 경로 설정(기본 후보)
TESSERACT_PATH_CAND = resource_path('tesseract', 'tesseract.exe')
TESSERACT_DIR_CAND  = resource_path('tesseract')
POPPLER_PATH_CAND   = resource_path('poppler', 'bin')

INPUT_DIR  = external_path('input_pdfs')
RESULT_DIR = external_path('result_pdfs')
APPROVAL_EXCEL_PATH = external_path('master_data.xlsx')
LOG_FILE    = external_path('automation.log')
CONFIG_PATH = external_path('config.json')

# 한글 stderr 방지(가능할 때)
os.environ['LANG'] = 'C'
os.environ['LC_ALL'] = 'C'
os.environ.pop('TESSDATA_PREFIX', None) # 충돌 방지

# 해석된 경로 (전역)
RESOLVED_TESSERACT_CMD = None
RESOLVED_TESSDATA_DIR  = None
RESOLVED_POPPLER_BIN   = None

# 윈도우 짧은 경로(8.3) 변환 - 한글/공백 경로 안전
def _short_path(p: str) -> str:
    #Windows 8.3 짧은 경로로 변환. 실패 시 원본 반환
    try:
        GetShortPathNameW = ctypes.windll.kernel32.GetShortPathNameW
        GetShortPathNameW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_uint]
        GetShortPathNameW.restype = ctypes.c_uint
        buf = ctypes.create_unicode_buffer(260)
        res = GetShortPathNameW(p, buf, 260)
        return buf.value if res else p
    except Exception:
        return p

def _exists(p: str) -> bool:
    try:
        return os.path.exists(p)
    except Exception:
        return False

# Tesseract 해석/저장 유틸
def _can_run_tesseract(cmd: str) -> bool:
    try:
        proc = subprocess.run([cmd, "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
        first = proc.stdout.decode(errors='ignore').splitlines()[0] if proc.stdout else ''
        logging.info(f"tesseract --version OK: {first}")
        return True
    except Exception as e:
        logging.error(f"tesseract 실행 실패: {e}")
        return False

def _guess_tessdata_dir(cmd_path: str) -> str | None:
    base = os.path.dirname(cmd_path)
    candidates = [
        os.path.join(base, 'tessdata'),
        os.path.join(os.path.dirname(base), 'tessdata'),
    ]
    for c in candidates:
        if os.path.isdir(c) and _exists(os.path.join(c, 'eng.traineddata')):
            return c
    return None

def _load_config() -> dict:
    if _exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def _save_config(d: dict) -> None:
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logging.error(f"설정 저장 실패: {e}")

def _candidate_tesseract_paths() -> list[str]:
    cands = [
        TESSERACT_PATH_CAND,
        external_path('tesseract', 'tesseract.exe'),
    ]
    exe_dir = Path(os.path.dirname(sys.executable)) if getattr(sys, 'frozen', False) \
              else Path(os.path.dirname(os.path.abspath(__file__)))
    cands += [
        str(exe_dir.joinpath('tesseract', 'tesseract.exe')),
        str(exe_dir.parent.joinpath('tesseract', 'tesseract.exe')),
    ]
    if getattr(sys, '_MEIPASS', None):
        cands.append(str(Path(sys._MEIPASS).joinpath('tesseract', 'tesseract.exe')))

    # 환경변수 힌트
    for k in ['TESSERACT_PATH', 'TESSERACT_CMD', 'TESSERACT_HOME']:
        v = os.environ.get(k)
        if v:
            cands.append(v if v.lower().endswith('tesseract.exe') else os.path.join(v, 'tesseract.exe'))

    # 중복 제거
    seen, uniq = set(), []
    for p in cands:
        if p and p not in seen:
            uniq.append(p); seen.add(p)
    return uniq

def resolve_tesseract_interactive() -> tuple[str, str]:
    #최종 (tesseract_cmd, tessdata_dir) 반환. 못 찾으면 사용자 선택 + 저장
    cfg = _load_config()
    if 'tesseract_cmd' in cfg and _exists(cfg['tesseract_cmd']) and _can_run_tesseract(cfg['tesseract_cmd']):
        cmd = cfg['tesseract_cmd']
        td  = cfg.get('tessdata_dir') or _guess_tessdata_dir(cmd)
        if td and os.path.isdir(td):
            return cmd, td

    for cand in _candidate_tesseract_paths():
        if _exists(cand) and _can_run_tesseract(cand):
            td = _guess_tessdata_dir(cand)
            if td:
                _save_config({'tesseract_cmd': cand, 'tessdata_dir': td})
                return cand, td

    logging.warning("tesseract.exe를 자동으로 찾지 못했습니다. 파일을 선택해주세요.")
    root = tk.Tk(); root.withdraw()
    file_path = filedialog.askopenfilename(
        title='tesseract.exe 선택',
        filetypes=[('tesseract.exe', 'tesseract.exe'), ('모든 파일', '*.*')]
    )
    root.destroy()
    if not file_path:
        raise RuntimeError("tesseract.exe를 선택하지 않았습니다.")
    if not _can_run_tesseract(file_path):
        raise RuntimeError("선택한 tesseract.exe를 실행할 수 없습니다.")

    td = _guess_tessdata_dir(file_path)
    if not td:
        raise RuntimeError("tessdata 폴더를 찾을 수 없습니다. tesseract.exe와 같은 폴더(또는 부모)에 tessdata가 있어야 합니다.")

    _save_config({'tesseract_cmd': file_path, 'tessdata_dir': td})
    return file_path, td

# Poppler 경로 해석
def resolve_poppler_bin() -> str | None:
    # 기본 후보 먼저
    if os.path.isdir(POPPLER_PATH_CAND):
        return POPPLER_PATH_CAND
    # 실행 폴더 바로 아래
    cand = external_path('poppler', 'bin')
    if os.path.isdir(cand):
        return cand
    return None

# 로그 핸들러(GUI)
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.configure(state='disabled')
            self.text_widget.yview(tk.END)
        self.text_widget.after(0, append)

# 디버그 로그
def debug_print_paths():
    logging.info(f"frozen={getattr(sys, 'frozen', False)}  MEIPASS={getattr(sys, '_MEIPASS', None)}")
    logging.info(f"exe={sys.executable}")
    logging.info(f"argv0={sys.argv[0]}")
    logging.info(f"cwd={os.getcwd()}")
    base = getattr(sys, '_MEIPASS', None)
    if base is None:
        base = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) \
               else os.path.dirname(os.path.abspath(__file__))
    logging.info(f"BASE DIR: {base}")

    logging.info(f"Tesseract exe (cand): {TESSERACT_PATH_CAND} (exists={_exists(TESSERACT_PATH_CAND)})")
    logging.info(f"Tesseract dir (cand): {TESSERACT_DIR_CAND} (exists={os.path.isdir(TESSERACT_DIR_CAND)})")
    logging.info(f"Poppler bin (cand): {POPPLER_PATH_CAND} (exists={os.path.isdir(POPPLER_PATH_CAND)})")
    logging.info(f"Input dir: {INPUT_DIR} (exists={os.path.isdir(INPUT_DIR)})")
    logging.info(f"Result dir: {RESULT_DIR} (exists={os.path.isdir(RESULT_DIR)})")
    logging.info(f"Excel path: {APPROVAL_EXCEL_PATH} (exists={os.path.exists(APPROVAL_EXCEL_PATH)})")
    if RESOLVED_TESSERACT_CMD:
        logging.info(f"[RESOLVED] Tesseract cmd: {RESOLVED_TESSERACT_CMD}")
    if RESOLVED_TESSDATA_DIR:
        logging.info(f"[RESOLVED] Tesseract tessdata: {RESOLVED_TESSDATA_DIR}")
    if RESOLVED_POPPLER_BIN:
        logging.info(f"[RESOLVED] Poppler bin: {RESOLVED_POPPLER_BIN}")

# 디렉토리 준비
def setup_directories():
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(RESULT_DIR, exist_ok=True)
    logging.info(f"폴더 확인 및 설정 완료: '{INPUT_DIR}', '{RESULT_DIR}'")

# OCR: Tesseract 직접 호출 (pytesseract 우회)
def ocr_with_tesseract(image_bgr: np.ndarray, tessdata_dir: str, psm: int = 6) -> str:
    """
    Tesseract를 직접 호출하여 stdout만 UTF-8으로 읽음.
    stderr는 DEVNULL로 버려서 윈도우 한글 로케일 디코딩 이슈 방지.
    """
    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
        temp_path = tmp.name
    try:
        cv2.imwrite(temp_path, image_bgr)
        cmd = [
            _short_path(RESOLVED_TESSERACT_CMD),
            _short_path(temp_path), 'stdout',
            '--oem', '3',
            '--psm', str(psm),
            '--tessdata-dir', _short_path(tessdata_dir),
            '-l', 'kor+eng',
            '--loglevel', 'OFF',
        ]
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, check=True)
        return proc.stdout.decode('utf-8', errors='ignore')
    finally:
        try:
            os.remove(temp_path)
        except Exception:
            pass

# ===================================================

# 전표번호 추출
def extract_doc_number(pdf_path: str) -> str | None:
    try:
        logging.info(f"[{os.path.basename(pdf_path)}] 에서 전표번호를 찾습니다...")

        # Poppler/경로 안전화
        safe_pdf_path = _short_path(pdf_path)
        safe_poppler  = _short_path(RESOLVED_POPPLER_BIN) if RESOLVED_POPPLER_BIN else None

        # 1) PDF 첫 페이지 이미지 변환
        images = convert_from_path(
            safe_pdf_path,
            first_page=1, last_page=1, dpi=300,
            poppler_path=safe_poppler
        )
        if not images:
            logging.error(f"[{os.path.basename(pdf_path)}] PDF를 이미지로 바꾸지 못했습니다.")
            return None

        # 2) 전처리(흑백+이진화)
        cv_image = np.array(images[0])
        gray = cv2.cvtColor(cv_image, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)

        # 3) OCR (PSM 6: 단일 블록)
        ocr_text = ocr_with_tesseract(binary, RESOLVED_TESSDATA_DIR, psm=6)

        # 4) 전표번호 패턴 추출 (O→0 보정 포함)
        pattern = re.compile(r"전\s*표\s*번\s*호\s*[:：.\-]?\s*([0-9O]{8}\s*[-–]?\s*[0-9O]{5})", re.IGNORECASE)
        match = pattern.search(ocr_text)
        if not match:
            logging.warning(" -> 실패: 전표번호 패턴을 찾지 못했습니다.")
            logging.debug(f"OCR 전체 텍스트:\n---\n{ocr_text}\n---")
            return None

        raw_number = match.group(1).replace('O', '0')  # O(알파벳) → 0(숫자) 보정
        cleaned = re.sub(r'[\s\-–]', '', raw_number)
        if len(cleaned) != 13:
            logging.warning(f" -> 패턴 매칭은 되었으나 자릿수 불일치: {cleaned}")
            return None

        doc_number = f"{cleaned[:8]}-{cleaned[8:]}"
        logging.info(f" -> 성공! 전표번호: '{doc_number}'")
        return doc_number

    except Exception as e:
        logging.error(f"[{os.path.basename(pdf_path)}] 처리 중 오류 발생: {e}")
        return None

# 메인 로직
def automation_main_logic():
    if not RESOLVED_TESSERACT_CMD or not RESOLVED_TESSDATA_DIR:
        logging.error("Tesseract 경로가 해석되지 않았습니다. (초기화 실패)")
        return
    setup_directories()
    logging.info("===== 스캔한 문서의 파일명을 결의적요로 변경합니다. =====")

    # 1) 입력 PDF 모으기
    try:
        pdf_files = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith('.pdf')]
        if not pdf_files:
            logging.warning(f"'{INPUT_DIR}' 폴더에 처리할 PDF 파일이 없습니다. 프로그램을 종료합니다.")
            return
        logging.info(f"총 {len(pdf_files)}개의 PDF 파일을 발견했습니다.")
    except FileNotFoundError:
        logging.error(f"입력 폴더 '{INPUT_DIR}'를 찾을 수 없습니다. 프로그램을 종료합니다.")
        return

    # 2) 각 파일 OCR
    rows = []
    for pdf_file in pdf_files:
        pdf_path = os.path.join(INPUT_DIR, pdf_file)
        doc_number = extract_doc_number(pdf_path)
        rows.append({'원본 파일명': pdf_file, '전표번호': doc_number})

    # 3) 표 생성
    source_df = pd.DataFrame(rows)

    # 4) 마스터 엑셀과 매칭
    try:
        approval_df = pd.read_excel(APPROVAL_EXCEL_PATH, engine='openpyxl', dtype={'결의서승인번호': str})
        approval_df.drop_duplicates(subset=['결의서승인번호'], keep='first', inplace=True)

        merged_df = pd.merge(
            source_df,
            approval_df[['결의서승인번호', '결의적요']],
            left_on='전표번호',
            right_on='결의서승인번호',
            how='left'
        )
        merged_df['결의적요'] = merged_df['결의적요'].fillna('제목없음')

    except FileNotFoundError:
        logging.error(f"마스터 엑셀 파일 '{APPROVAL_EXCEL_PATH}'를 찾을 수 없습니다.")
        return
    except KeyError as e:
        logging.error(f"마스터 엑셀 파일에 필요한 열이 없습니다: {e} (필수: '결의서승인번호','결의적요')")
        return

    # 5) 파일명 변경/이동
    success, fail = 0, 0
    for _, row in merged_df.iterrows():
        original = row['원본 파일명']
        title    = row['결의적요']

        if pd.isna(row['전표번호']) or title == '제목없음':
            logging.warning(f" -> [{original}] 파일은 처리할 수 없어 건너뜁니다.")
            fail += 1
            continue

        safe_title = re.sub(r'[\\/*?:"<>|]', "", str(title))
        new_name   = f"{safe_title}.pdf"

        src = os.path.join(INPUT_DIR, original)
        dst = os.path.join(RESULT_DIR, new_name)

        try:
            shutil.move(src, dst)
            logging.info(f" -> 처리 완료: '{original}' -> '{new_name}'")
            success += 1
        except Exception as e:
            logging.error(f" -> [{original}] 파일 이동 중 오류 발생: {e}")
            fail += 1

    # 6) 보고서 저장
    final_status_path = os.path.join(RESULT_DIR, '_Automation_Report.xlsx')
    merged_df.to_excel(final_status_path, index=False, engine='openpyxl')

    logging.info("===== 모든 작업이 완료되었습니다. =====")
    logging.info(f"성공: {success}건, 실패/미처리: {fail}건")
    logging.info(f"상세 결과는 '{final_status_path}' 파일에서 확인하세요.")

# 스레드/GUI
def start_automation_thread(run_button, log_widget):
    run_button.config(state="disabled", text="처리 중...")
    thread = threading.Thread(target=automation_main_logic, daemon=True)
    thread.start()
    check_thread_status(thread, run_button)

def check_thread_status(thread, run_button):
    if thread.is_alive():
        run_button.after(100, lambda: check_thread_status(thread, run_button))
    else:
        run_button.config(state="normal", text="시작")
        logging.info("===== 작업 완료. 다시 실행할 수 있습니다. =====")

def create_gui():
    global RESOLVED_TESSERACT_CMD, RESOLVED_TESSDATA_DIR, RESOLVED_POPPLER_BIN

    root = tk.Tk()
    root.title("스캔 PDF 파일명 자동 변경 프로그램 v1.0")
    root.geometry("800x600")

    top_frame = tk.Frame(root, pady=10); top_frame.pack(fill=tk.X, side=tk.TOP)
    run_button = tk.Button(top_frame, text="시작", font=("Helvetica", 12, "bold"), width=20, height=2)
    run_button.pack()

    log_frame = tk.Frame(root, padx=10, pady=10); log_frame.pack(fill=tk.BOTH, expand=True)
    log_widget = ScrolledText(log_frame, state='disabled', wrap=tk.WORD, font=("Consolas", 10))
    log_widget.pack(fill=tk.BOTH, expand=True)

    footer_frame = tk.Frame(root, pady=5); footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
    author_label = tk.Label(footer_frame, text="CONTACT: wooyxxng@gmail.com", font=("Helvetica", 9), fg="gray")
    author_label.pack()

    # 로깅 설정
    logger = logging.getLogger(); logger.setLevel(logging.INFO)
    for h in logger.handlers[:]: logger.removeHandler(h)
    fh = logging.FileHandler(LOG_FILE, encoding='utf-8')
    fh.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(fh)
    th = TextHandler(log_widget)
    th.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(th)

    # 경로 해석
    RESOLVED_POPPLER_BIN = resolve_poppler_bin()
    try:
        cmd, td = resolve_tesseract_interactive()
        RESOLVED_TESSERACT_CMD = cmd
        RESOLVED_TESSDATA_DIR  = td
        logging.info(f"[RESOLVED] tesseract_cmd: {RESOLVED_TESSERACT_CMD}")
        logging.info(f"[RESOLVED] tessdata_dir : {RESOLVED_TESSDATA_DIR}")
        if RESOLVED_POPPLER_BIN:
            logging.info(f"[RESOLVED] poppler_bin  : {RESOLVED_POPPLER_BIN}")
        else:
            logging.warning("Poppler bin 폴더를 찾지 못했습니다. (pdf2image 변환 실패 가능)")
    except Exception as e:
        logging.error(f"Tesseract 경로 해석 실패: {e}")

    debug_print_paths()
    run_button.config(command=lambda: start_automation_thread(run_button, log_widget))
    root.mainloop()

if __name__ == "__main__":
    create_gui()
