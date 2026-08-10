#!/usr/bin/env python3
"""
Advanced HWP MCP Server
고도화된 한글 MCP 서버 - 한글의 모든 기능을 제어할 수 있는 MCP 서버
"""

import logging
import sys
import os
import re
from typing import Optional, Dict, Any, List, Tuple
import json

try:
    from mcp.server.fastmcp import FastMCP
    import win32com.client
    import pythoncom
    import win32api
    import win32con
    import win32gui
    import win32process
    import ctypes
except ImportError as e:
    print(f"필수 패키지가 설치되지 않음: {e}", file=sys.stderr)
    print("다음 명령어로 패키지를 설치하세요:", file=sys.stderr)
    print("pip install mcp fastmcp pywin32", file=sys.stderr)
    sys.exit(1)

# 로깅 설정 - MCP는 stdout을 JSON-RPC로 사용하므로 stderr로만 출력
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('hwp_mcp.log'),
        logging.StreamHandler(sys.stderr)
    ]
)
logger = logging.getLogger(__name__)

# MCP 서버 초기화
mcp = FastMCP("Advanced HWP Server")

class AdvancedHwpController:
    """고급 한글 컨트롤러 클래스"""
    
    def __init__(self):
        """한글 COM 객체 초기화"""
        self.hwp = None
        self.is_initialized = False
        self.current_document = None
        
    @staticmethod
    def ensure_security_module_registry():
        """보안모듈(FilePathCheckerModule) 레지스트리 자가 복구.
        값이 실제 DLL 경로가 아니면 저장소 동봉 DLL로 재등록한다.
        이 등록이 없으면 파일 접근 시마다 '접근 허용' 대화상자가 떠서 자동화가 멈춘다."""
        try:
            import winreg
            bundled = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   "security_module", "FilePathCheckerModule.dll")
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                                   r"Software\HNC\HwpAutomation\Modules")
            try:
                cur, _ = winreg.QueryValueEx(key, "FilePathCheckerModule")
            except OSError:
                cur = None
            if (not cur or not os.path.exists(str(cur))) and os.path.exists(bundled):
                winreg.SetValueEx(key, "FilePathCheckerModule", 0,
                                  winreg.REG_SZ, bundled)
                logger.info(f"보안모듈 레지스트리 자동 등록: {bundled}")
            winreg.CloseKey(key)
        except Exception as e:
            logger.warning(f"보안모듈 레지스트리 확인 실패(무시): {e}")

    def initialize(self):
        """한글 COM 객체 초기화"""
        try:
            pythoncom.CoInitialize()

            # 보안모듈 레지스트리 선확인 (파일 접근 대화상자 방지)
            self.ensure_security_module_registry()

            # 한글 프로그램이 설치되어 있는지 확인
            try:
                self.hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
            except:
                # gencache가 실패하면 일반 Dispatch 사용
                self.hwp = win32com.client.Dispatch("HWPFrame.HwpObject")

            # ===== 자동화 모드 설정: 모든 확인 대화상자 자동 승인 =====

            # 1. 파일 경로 체크 대화상자 비활성화
            try:
                self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            except:
                pass

            # 2. 보안 경고 대화상자 비활성화
            try:
                self.hwp.RegisterModule("SecurityModule", "")
            except:
                pass

            # 3. 개인정보 보호 기능 비활성화 (대화상자 방지)
            try:
                self.hwp.SetPrivateInfoProtection(0)
            except:
                pass

            # 4. 메시지 박스 자동 응답 설정 (모든 확인창 자동 Yes)
            try:
                # MessageBoxMode: 0=자동응답, 1=대화상자표시
                self.hwp.MessageBoxMode = 0
            except:
                pass

            # 5. 편집 모드를 자동화 모드로 설정
            try:
                self.hwp.EditMode = 1  # 1=자동화모드, 0=일반모드
            except:
                pass

            # 6. 화면 업데이트 일시 중지 (성능 향상)
            try:
                self.hwp.SetMessageBoxMode(0)  # 0=자동, 1=표시
            except:
                pass

            # 한글 창을 보이게 설정
            if self.hwp.XHwpWindows.Count > 0:
                self.hwp.XHwpWindows.Item(0).Visible = True

            self.is_initialized = True
            logger.info("한글 COM 객체 초기화 완료 (자동화 모드 활성화)")
            return True

        except Exception as e:
            logger.error(f"한글 COM 객체 초기화 실패: {e}")
            return False
    
    def __del__(self):
        """리소스 정리"""
        try:
            if self.is_initialized:
                pythoncom.CoUninitialize()
        except:
            pass
    
    def check_initialization(self):
        """초기화 상태 확인"""
        if not self.is_initialized:
            if not self.initialize():
                raise Exception("한글 프로그램이 설치되지 않았거나 초기화할 수 없습니다.")
        return True

# 전역 컨트롤러 인스턴스
hwp_controller = AdvancedHwpController()


# ============================================================
# 핵심 인프라 헬퍼 (v2)
#  - 실전 검증에서 발견된 HWP COM 함정들을 한 곳에서 처리한다
# ============================================================

MM_TO_HWPUNIT = 283.465          # 1mm = 283.465 HWPUNIT (7200unit/inch)


def _mm(hwp, mm: float) -> int:
    """mm -> HWPUNIT. MiliToHwpUnit이 있으면 사용, 없으면 상수 변환."""
    try:
        return hwp.MiliToHwpUnit(mm)
    except Exception:
        return int(mm * MM_TO_HWPUNIT)


def _rgb_to_colorref(color) -> int:
    """RGB(사람이 아는 순서) -> HWP COLORREF(0x00BBGGRR).
    'red'/'#FF3B30'/'FF3B30'/0xFF3B30 모두 허용."""
    named = {
        "black": 0x000000, "white": 0xFFFFFF, "red": 0xFF0000,
        "blue": 0x0000FF, "green": 0x008000, "yellow": 0xFFD400,
        "orange": 0xFF7A00, "purple": 0x800080, "cyan": 0x00B7C3,
        "gray": 0x777777, "grey": 0x777777, "navy": 0x1A2B4A,
        "brown": 0x8B2B00,
    }
    try:
        if isinstance(color, str):
            key = color.strip().lower()
            if key in named:
                rgb = named[key]
            else:
                rgb = int(key.lstrip("#"), 16)
        else:
            rgb = int(color)
    except (ValueError, TypeError):
        logger.warning(f"알 수 없는 색상 '{color}' — 검정으로 대체")
        rgb = 0x000000
    r = (rgb >> 16) & 0xFF
    g = (rgb >> 8) & 0xFF
    b = rgb & 0xFF
    return (b << 16) | (g << 8) | r          # COLORREF = BGR


def _fix_mojibake(text: str) -> str:
    """GetTextFile 등이 CP949 바이트를 Latin-1로 잘못 해석해 반환하는 경우 복구.
    (실증: 한글이 '¡à ¿¹¼úºÐ¾ß'처럼 보이는 증상)"""
    if not text:
        return text
    sample = text[:4000]
    hangul = sum(1 for ch in sample if '가' <= ch <= '힣')
    high_latin = sum(1 for ch in sample if 0x80 <= ord(ch) <= 0xFF)
    if high_latin > 20 and high_latin > hangul * 3:
        try:
            repaired = text.encode("latin-1", "replace").decode("cp949", "replace")
            rep_hangul = sum(1 for ch in repaired[:4000] if '가' <= ch <= '힣')
            # 오탐 방지: 복구본에 '충분한' 한글이 나와야 진짜 모지바케로 판정
            # (정상 유럽어 문서를 우연히 한글 몇 자로 재해석하는 경우 배제)
            if rep_hangul >= 10 and rep_hangul > hangul * 2 and rep_hangul > high_latin * 0.3:
                return repaired
        except Exception:
            pass
    return text


def _extract_text(hwp, scope: str = "document") -> str:
    """유니코드 안전 텍스트 추출.
    scope='document' 전체 / 'selection' 현재 선택.
    1순위 InitScan/GetText(문서 전체, 커서 불변), 2순위 GetTextFile+모지바케 복구."""
    if scope == "document":
        # 1순위: SelectAll + GetTextFile — 표 안 텍스트까지 포함됨(검증됨).
        #        InitScan 방식은 표 셀을 건너뛰어 일부만 반환하는 문제가 있어 폴백으로만 사용.
        pos = None
        try:
            pos = hwp.GetPos()
        except Exception:
            pass
        hwp.HAction.Run("SelectAll")
        raw = hwp.GetTextFile("TEXT", "")
        hwp.HAction.Run("Cancel")
        if pos:
            try:
                hwp.SetPos(*pos)
            except Exception:
                pass
        text = _fix_mojibake(raw or "")
        if text:
            return text
        # 폴백: InitScan/GetText (본문 문단 위주)
        try:
            hwp.InitScan(0, 0x0077)
            chunks = []
            for _ in range(100000):
                state, txt = hwp.GetText()
                if state in (0, 1):          # 0=완료, 1=오류
                    break
                if txt:
                    chunks.append(txt)
            hwp.ReleaseScan()
            return "\r\n".join(chunks)      # 문단 경계 유지
        except Exception:
            try:
                hwp.ReleaseScan()
            except Exception:
                pass
            return ""
    else:
        # 선택 영역 전용 — 폴백으로 GetTextFile("TEXT","")를 쓰면 문서 '전체'가
        # 반환되어 호출자가 선택 내용으로 오인하므로 절대 금지 (빈 선택 = 빈 문자열이 정답)
        raw = hwp.GetTextFile("TEXT", "saveblock:true")
        return _fix_mojibake(raw or "")


def _find_once(hwp, text: str, from_start: bool = True) -> bool:
    """RepeatFind로 커서를 실제 이동시키는 찾기. 찾으면 해당 텍스트가 선택된 상태."""
    if from_start:
        hwp.HAction.Run("MoveDocBegin")
    ps = hwp.HParameterSet.HFindReplace
    hwp.HAction.GetDefault("RepeatFind", ps.HSet)
    ps.FindString = text
    ps.IgnoreMessage = 1
    ps.Direction = 0
    ps.SeveralWords = 0              # 정확 문자열 매칭 (1이면 단어 단위 오동작)
    return bool(hwp.HAction.Execute("RepeatFind", ps.HSet))


def _replace_all(hwp, find_str: str, replace_str: str) -> bool:
    """AllReplace 안전판: IgnoreMessage/정확매칭/문서 처음부터. 반환값: 실행 후 잔존 여부 검사."""
    hwp.HAction.Run("MoveDocBegin")
    ps = hwp.HParameterSet.HFindReplace
    hwp.HAction.GetDefault("AllReplace", ps.HSet)
    ps.FindString = find_str
    ps.ReplaceString = replace_str
    ps.IgnoreMessage = 1
    ps.Direction = 0
    ps.SeveralWords = 0
    ps.ReplaceMode = 1
    hwp.HAction.Execute("AllReplace", ps.HSet)
    # 검증: 아직 남아 있으면 실패로 간주.
    # 단, 치환 결과가 검색어를 (대소문자 무시로) 포함하면 재발견되므로 검증 생략
    # (HWP 기본 찾기는 대소문자를 무시함 — 'ai'→'AI' 치환의 오판 방지)
    if replace_str and find_str.lower() in replace_str.lower():
        return True
    remained = _find_once(hwp, find_str)
    hwp.HAction.Run("Cancel")
    return not remained


def _charshape_selection(hwp, size_pt=None, bold=None, italic=None,
                         underline=None, color=None, face=None):
    """현재 선택 영역에 글자 서식 적용. color는 RGB 입력."""
    ps = hwp.HParameterSet.HCharShape
    hwp.HAction.GetDefault("CharShape", ps.HSet)
    if face:
        for attr in ("FaceNameHangul", "FaceNameLatin", "FaceNameHanja",
                     "FaceNameJapanese", "FaceNameOther", "FaceNameSymbol",
                     "FaceNameUser"):
            try:
                setattr(ps, attr, face)
            except Exception:
                pass
    if size_pt is not None:
        ps.Height = int(size_pt * 100)
    if bold is not None:
        ps.Bold = 1 if bold else 0
    if italic is not None:
        ps.Italic = 1 if italic else 0
    if underline is not None:
        # HCharShape에는 'Underline' 속성이 없음(UnderlineType이 실명) — 실기기 typelib 확인
        try:
            ps.UnderlineType = 1 if underline else 0
        except Exception:
            try:
                ps.HSet.SetItem("UnderlineType", 1 if underline else 0)
            except Exception:
                pass
    if color is not None:
        ps.TextColor = _rgb_to_colorref(color)
    hwp.HAction.Execute("CharShape", ps.HSet)


def _clear_list_content(hwp, lst: int, max_paras: int = 800) -> int:
    """지정 리스트(셀) 내부를 완전히 비운다. 마지막 문단 끝까지 선택해야 잔재가 안 남고,
    자동번호(HeadingType)는 텍스트 삭제로 안 지워지므로 함께 해제한다. 반환: 지운 문단 수."""
    cleared = 0
    for _ in range(6):
        last = 0
        for i in range(0, max_paras):
            hwp.SetPos(lst, i, 0)
            cur = hwp.GetPos()
            if cur[0] == lst and cur[1] == i:
                last = i
            else:
                break
        hwp.SetPos(lst, last, 0)
        hwp.HAction.Run("MoveParaEnd")
        endc = hwp.GetPos()[2]
        if last == 0 and endc == 0:
            break
        hwp.SetPos(lst, 0, 0)
        hwp.SelectText(0, 0, last, endc)
        hwp.HAction.Run("Delete")
        cleared = max(cleared, last + 1)
        hwp.SetPos(lst, 0, 0)
    # 남은 문단들의 자동번호 제거
    for i in range(0, 30):
        hwp.SetPos(lst, i, 0)
        cur = hwp.GetPos()
        if cur[0] != lst or cur[1] != i:
            break
        try:
            ps = hwp.HParameterSet.HParaShape
            hwp.HAction.GetDefault("ParagraphShape", ps.HSet)
            if getattr(ps, "HeadingType", 0) != 0:
                ps.HeadingType = 0
                hwp.HAction.Execute("ParagraphShape", ps.HSet)
        except Exception:
            pass
    hwp.SetPos(lst, 0, 0)
    return cleared


def _iter_pictures(hwp):
    """문서 내 그림(gso) 컨트롤 목록."""
    pics = []
    ctrl = hwp.HeadCtrl
    while ctrl:
        try:
            if ctrl.CtrlID == "gso":
                pics.append(ctrl)
        except Exception:
            pass
        try:
            ctrl = ctrl.Next
        except Exception:
            break
    return pics

@mcp.tool()
def initialize_hwp() -> str:
    """한글 프로그램을 초기화합니다."""
    try:
        if hwp_controller.initialize():
            return "한글 프로그램 초기화 성공"
        else:
            return "한글 프로그램 초기화 실패"
    except Exception as e:
        logger.error(f"초기화 중 오류: {e}")
        return f"초기화 실패: {e}"

@mcp.tool()
def get_running_hwp_documents() -> str:
    """실행 중인 한글에서 열린 문서 목록을 조회합니다."""
    try:
        pythoncom.CoInitialize()
        
        hwp = None
        
        # 1. 이미 연결된 컨트롤러가 있으면 사용
        if hwp_controller.is_initialized and hwp_controller.hwp:
            hwp = hwp_controller.hwp
        else:
            # 2. GetActiveObject 시도
            try:
                hwp = win32com.client.GetActiveObject("HWPFrame.HwpObject")
            except:
                pass
            
            # 3. GetActiveObject 실패 시 Dispatch로 연결 시도
            if hwp is None:
                try:
                    hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                    if hwp.XHwpDocuments.Count == 0:
                        return "실행 중인 한글 프로그램이 없습니다. initialize_hwp()로 시작하세요."
                except:
                    return "실행 중인 한글 프로그램이 없습니다. initialize_hwp()로 시작하세요."
        
        # 열린 문서 목록 가져오기
        doc_count = hwp.XHwpDocuments.Count
        if doc_count == 0:
            return "한글이 실행 중이지만 열린 문서가 없습니다."
        
        documents = []
        for i in range(doc_count):
            doc = hwp.XHwpDocuments.Item(i)
            doc_path = doc.Path if doc.Path else "(새 문서)"
            doc_name = doc.Path.split("\\")[-1] if doc.Path else f"새 문서 {i+1}"
            documents.append({
                "index": i,
                "name": doc_name,
                "path": doc_path
            })
        
        result = f"현재 연결된 한글에서 열린 문서 ({doc_count}개):\n"
        for doc in documents:
            result += f"  [{doc['index']}] {doc['name']}\n"
            result += f"      경로: {doc['path']}\n"
        
        # 여러 한글 프로세스 확인 안내
        try:
            import subprocess
            result_proc = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq Hwp.exe', '/FO', 'CSV', '/NH'], 
                                        capture_output=True, text=True)
            hwp_count = len([line for line in result_proc.stdout.strip().split('\n') if 'Hwp.exe' in line])
            if hwp_count > 1:
                result += f"\n⚠️ 주의: 한글 프로그램이 {hwp_count}개 실행 중입니다.\n"
                result += "   다른 인스턴스에 연결하려면 list_all_hwp_windows()를 호출하세요."
        except:
            pass
        
        logger.info(f"열린 문서 목록 조회 완료: {doc_count}개")
        return result
        
    except Exception as e:
        logger.error(f"문서 목록 조회 실패: {e}")
        return f"문서 목록 조회 실패: {e}"

@mcp.tool()
def list_all_hwp_windows() -> str:
    """실행 중인 모든 한글 창 목록을 조회합니다. (창 제목으로 파일명 확인)"""
    try:
        hwp_windows = []
        
        def enum_windows_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                class_name = win32gui.GetClassName(hwnd)
                
                if 'Hwp' in class_name or '한글' in window_text or window_text.endswith('.hwp'):
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    results.append({
                        'hwnd': hwnd,
                        'title': window_text,
                        'class': class_name,
                        'pid': pid
                    })
            return True
        
        win32gui.EnumWindows(enum_windows_callback, hwp_windows)
        
        if not hwp_windows:
            return "실행 중인 한글 창을 찾을 수 없습니다."
        
        result = f"실행 중인 한글 창 ({len(hwp_windows)}개):\n"
        for i, win in enumerate(hwp_windows):
            result += f"  [{i}] {win['title']}\n"
            result += f"      PID: {win['pid']}, HWND: {win['hwnd']}\n"
        
        result += "\n특정 창에 연결하려면 connect_to_hwp_window(파일명 일부)를 호출하세요."
        
        logger.info(f"한글 창 목록 조회: {len(hwp_windows)}개")
        return result
        
    except Exception as e:
        logger.error(f"한글 창 목록 조회 실패: {e}")
        return f"한글 창 목록 조회 실패: {e}"

@mcp.tool()
def connect_to_hwp_window(search_text: str) -> str:
    """특정 파일명이 포함된 한글 창을 활성화하고 연결합니다."""
    try:
        pythoncom.CoInitialize()
        
        hwp_windows = []
        
        def enum_windows_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                class_name = win32gui.GetClassName(hwnd)
                
                if 'Hwp' in class_name or '한글' in window_text or window_text.endswith('.hwp'):
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    results.append({
                        'hwnd': hwnd,
                        'title': window_text,
                        'class': class_name,
                        'pid': pid
                    })
            return True
        
        win32gui.EnumWindows(enum_windows_callback, hwp_windows)
        
        if not hwp_windows:
            return "실행 중인 한글 창을 찾을 수 없습니다."
        
        # 검색어가 포함된 창 찾기
        target_window = None
        for win in hwp_windows:
            if search_text.lower() in win['title'].lower():
                target_window = win
                break
        
        if target_window is None:
            titles = [w['title'] for w in hwp_windows]
            return f"'{search_text}'이(가) 포함된 창을 찾을 수 없습니다.\n실행 중인 창: {titles}"
        
        # 해당 창 활성화
        hwnd = target_window['hwnd']
        
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        
        try:
            win32gui.SetForegroundWindow(hwnd)
        except:
            try:
                win32gui.BringWindowToTop(hwnd)
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            except:
                pass
        
        import time
        time.sleep(0.3)
        
        try:
            hwp_controller.hwp = win32com.client.GetActiveObject("HWPFrame.HwpObject")
        except:
            hwp_controller.hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        
        hwp_controller.is_initialized = True
        
        doc_count = hwp_controller.hwp.XHwpDocuments.Count
        if doc_count > 0:
            current_doc = hwp_controller.hwp.XHwpDocuments.Item(0)
            hwp_controller.current_document = current_doc.Path if current_doc.Path else "새 문서"
        
        logger.info(f"한글 창 연결 완료: {target_window['title']}")
        return f"'{target_window['title']}' 창에 연결되었습니다. (열린 문서: {doc_count}개)"
        
    except Exception as e:
        logger.error(f"한글 창 연결 실패: {e}")
        return f"한글 창 연결 실패: {e}"

@mcp.tool()
def connect_to_running_hwp() -> str:
    """이미 실행 중인 한글 프로그램에 연결합니다."""
    try:
        pythoncom.CoInitialize()
        
        hwp = None
        
        try:
            hwp = win32com.client.GetActiveObject("HWPFrame.HwpObject")
        except:
            pass
        
        if hwp is None:
            try:
                hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                if hwp.XHwpDocuments.Count == 0:
                    hwp_controller.hwp = hwp
                    hwp_controller.is_initialized = True
                    return "실행 중인 한글이 없어 새로 시작했습니다. (열린 문서: 0개)"
            except Exception as e:
                return f"한글 연결 실패: {e}"
        
        hwp_controller.hwp = hwp
        hwp_controller.is_initialized = True
        
        doc_count = hwp_controller.hwp.XHwpDocuments.Count
        if doc_count > 0:
            current_doc = hwp_controller.hwp.XHwpDocuments.Item(0)
            hwp_controller.current_document = current_doc.Path if current_doc.Path else "새 문서"
        
        logger.info("실행 중인 한글에 연결 완료")
        return f"실행 중인 한글에 연결되었습니다. (열린 문서: {doc_count}개)"
        
    except Exception as e:
        logger.error(f"한글 연결 실패: {e}")
        return f"한글 연결 실패: {e}"

@mcp.tool()
def switch_to_document(file_name: str) -> str:
    """열린 문서 중 특정 파일로 전환합니다. 파일명 일부만 입력해도 됩니다."""
    try:
        pythoncom.CoInitialize()
        
        hwp = None
        
        if hwp_controller.is_initialized and hwp_controller.hwp:
            hwp = hwp_controller.hwp
        else:
            try:
                hwp = win32com.client.GetActiveObject("HWPFrame.HwpObject")
            except:
                pass
            
            if hwp is None:
                try:
                    hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                    if hwp.XHwpDocuments.Count == 0:
                        return "실행 중인 한글 프로그램이 없습니다."
                except:
                    return "실행 중인 한글 프로그램이 없습니다."
        
        doc_count = hwp.XHwpDocuments.Count
        if doc_count == 0:
            return "열린 문서가 없습니다."
        
        found_doc = None
        found_index = -1
        
        for i in range(doc_count):
            doc = hwp.XHwpDocuments.Item(i)
            doc_path = doc.Path if doc.Path else ""
            doc_name = doc_path.split("\\")[-1] if doc_path else f"새 문서 {i+1}"
            
            if file_name.lower() in doc_name.lower() or file_name.lower() in doc_path.lower():
                found_doc = doc
                found_index = i
                break
        
        if found_doc is None:
            return f"'{file_name}'이(가) 포함된 문서를 찾을 수 없습니다."
        
        hwp.XHwpDocuments.Item(found_index).SetActive()
        
        hwp_controller.hwp = hwp
        hwp_controller.is_initialized = True
        hwp_controller.current_document = found_doc.Path
        
        doc_name = found_doc.Path.split("\\")[-1] if found_doc.Path else f"새 문서"
        logger.info(f"문서 전환 완료: {doc_name}")
        return f"'{doc_name}' 문서로 전환했습니다."
        
    except Exception as e:
        logger.error(f"문서 전환 실패: {e}")
        return f"문서 전환 실패: {e}"

@mcp.tool()
def get_active_document_info() -> str:
    """현재 활성화된 문서의 정보를 조회합니다."""
    try:
        pythoncom.CoInitialize()
        
        hwp = None
        
        if hwp_controller.is_initialized and hwp_controller.hwp:
            hwp = hwp_controller.hwp
        else:
            try:
                hwp = win32com.client.GetActiveObject("HWPFrame.HwpObject")
            except:
                pass
            
            if hwp is None:
                try:
                    hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                    if hwp.XHwpDocuments.Count == 0:
                        return "실행 중인 한글 프로그램이 없습니다."
                except:
                    return "실행 중인 한글 프로그램이 없습니다."
        
        doc_count = hwp.XHwpDocuments.Count
        if doc_count == 0:
            return "열린 문서가 없습니다."
        
        active_doc = hwp.XHwpDocuments.Item(0)
        doc_path = active_doc.Path if active_doc.Path else "(새 문서 - 저장되지 않음)"
        doc_name = doc_path.split("\\")[-1] if active_doc.Path else "새 문서"
        
        page_count = hwp.PageCount
        
        result = f"""현재 활성 문서 정보:
- 파일명: {doc_name}
- 경로: {doc_path}
- 페이지 수: {page_count}
- 총 열린 문서 수: {doc_count}개"""
        
        logger.info(f"활성 문서 정보 조회: {doc_name}")
        return result
        
    except Exception as e:
        logger.error(f"문서 정보 조회 실패: {e}")
        return f"문서 정보 조회 실패: {e}"

@mcp.tool()
def create_document() -> str:
    """새 한글 문서를 생성합니다."""
    try:
        hwp_controller.check_initialization()
        
        hwp_controller.hwp.HAction.Run("FileNew")
        hwp_controller.current_document = "new_document"
        
        logger.info("새 문서 생성 완료")
        return "새 문서가 생성되었습니다."
        
    except Exception as e:
        logger.error(f"문서 생성 실패: {e}")
        return f"문서 생성 실패: {e}"

@mcp.tool()
def open_document(file_path: str) -> str:
    """지정된 경로의 한글 문서를 엽니다."""
    try:
        hwp_controller.check_initialization()
        
        if not os.path.exists(file_path):
            return f"파일을 찾을 수 없습니다: {file_path}"
        
        try:
            hwp_controller.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except:
            pass
        
        # 모든 대화상자 자동 처리: 버전 경고, 암호, 접근 권한 등
        result = hwp_controller.hwp.Open(file_path, "HWP", "forceopen:true;versionwarning:false;suspendpassword:true")

        if result:
            hwp_controller.current_document = file_path

            # 한글 창을 화면 제일 앞으로 가져오기
            try:
                hwnd = hwp_controller.hwp.XHwpWindows.Active_XHwpWindow.WindowHandle
                win32gui.SetForegroundWindow(hwnd)
                win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
            except:
                pass

            logger.info(f"문서 열기 완료: {file_path}")
            return f"문서를 열었습니다: {file_path}"
        else:
            hwp_controller.hwp.HAction.GetDefault("FileOpen", hwp_controller.hwp.HParameterSet.HFileOpenSave.HSet)
            hwp_controller.hwp.HParameterSet.HFileOpenSave.filename = file_path
            hwp_controller.hwp.HParameterSet.HFileOpenSave.Format = "HWP"
            hwp_controller.hwp.HAction.Execute("FileOpen", hwp_controller.hwp.HParameterSet.HFileOpenSave.HSet)
            
            hwp_controller.current_document = file_path
            logger.info(f"문서 열기 완료 (대체방법): {file_path}")
            return f"문서를 열었습니다: {file_path}"
        
    except Exception as e:
        logger.error(f"문서 열기 실패: {e}")
        return f"문서 열기 실패: {e}"

@mcp.tool()
def save_document(file_path: Optional[str] = None) -> str:
    """현재 문서를 저장합니다."""
    try:
        hwp_controller.check_initialization()
        
        if file_path:
            act = hwp_controller.hwp.CreateAction("FileSaveAs")
            pset = act.CreateSet()
            pset.SetItem("filename", file_path)
            pset.SetItem("format", "HWP")
            act.Execute(pset)
            
            hwp_controller.current_document = file_path
            logger.info(f"문서 저장 완료: {file_path}")
            return f"문서를 저장했습니다: {file_path}"
        else:
            hwp_controller.hwp.HAction.Run("FileSave")
            logger.info("문서 저장 완료")
            return "문서를 저장했습니다."
            
    except Exception as e:
        logger.error(f"문서 저장 실패: {e}")
        return f"문서 저장 실패: {e}"

@mcp.tool()
def close_document(save_changes: bool = False) -> str:
    """현재 문서를 닫습니다."""
    try:
        hwp_controller.check_initialization()
        
        if save_changes:
            hwp_controller.hwp.HAction.Run("FileSave")
        
        hwp_controller.hwp.HAction.Run("FileClose")
        hwp_controller.current_document = None
        
        logger.info("문서 닫기 완료")
        return "문서를 닫았습니다."
        
    except Exception as e:
        logger.error(f"문서 닫기 실패: {e}")
        return f"문서 닫기 실패: {e}"

@mcp.tool()
def close_all_documents(save_changes: bool = False) -> str:
    """모든 문서를 닫습니다."""
    try:
        hwp_controller.check_initialization()
        
        closed_count = 0
        max_attempts = 100
        
        for _ in range(max_attempts):
            try:
                if hwp_controller.hwp.XHwpDocuments.Count == 0:
                    break
                
                if save_changes:
                    hwp_controller.hwp.HAction.Run("FileSave")
                
                hwp_controller.hwp.XHwpDocuments.Item(0).SetModified(False)
                hwp_controller.hwp.HAction.Run("FileClose")
                closed_count += 1
            except Exception as e:
                logger.warning(f"문서 닫기 중 오류: {e}")
                break
        
        hwp_controller.current_document = None
        
        logger.info(f"모든 문서 닫기 완료: {closed_count}개")
        return f"모든 문서를 닫았습니다. ({closed_count}개)"
        
    except Exception as e:
        logger.error(f"모든 문서 닫기 실패: {e}")
        return f"모든 문서 닫기 실패: {e}"

@mcp.tool()
def quit_hwp() -> str:
    """한글 프로그램을 종료합니다."""
    try:
        hwp_controller.check_initialization()

        try:
            hwp_controller.hwp.Clear(1)      # 저장 안 함 모드로 정리 (COM 예외 무시)
        except Exception:
            pass
        hwp_controller.hwp.Quit()
        hwp_controller.hwp = None
        hwp_controller.is_initialized = False
        hwp_controller.current_document = None
        
        logger.info("한글 프로그램 종료 완료")
        return "한글 프로그램을 종료했습니다."
        
    except Exception as e:
        logger.error(f"한글 종료 실패: {e}")
        return f"한글 종료 실패: {e}"

@mcp.tool()
def insert_text(text: str, position: str = "current") -> str:
    """텍스트를 삽입합니다."""
    try:
        hwp_controller.check_initialization()
        
        if position == "current":
            act = hwp_controller.hwp.CreateAction("InsertText")
            pset = act.CreateSet()
            pset.SetItem("Text", text)
            act.Execute(pset)
        
        logger.info(f"텍스트 삽입 완료: {text[:50]}...")
        return f"텍스트를 삽입했습니다: {text[:50]}..."
        
    except Exception as e:
        logger.error(f"텍스트 삽입 실패: {e}")
        return f"텍스트 삽입 실패: {e}"

@mcp.tool()
def insert_text_at_position(text: str, x: int = 0, y: int = 0) -> str:
    """지정된 좌표에 텍스트를 삽입합니다."""
    try:
        hwp_controller.check_initialization()
        
        hwp_controller.hwp.SetPosBySet(x, y)
        
        act = hwp_controller.hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        pset.SetItem("Text", text)
        act.Execute(pset)
        
        logger.info(f"위치 ({x}, {y})에 텍스트 삽입 완료")
        return f"위치 ({x}, {y})에 텍스트 '{text}'를 삽입했습니다."
        
    except Exception as e:
        logger.error(f"위치별 텍스트 삽입 실패: {e}")
        return f"위치별 텍스트 삽입 실패: {e}"

@mcp.tool()
def apply_font_format(font_name: str = "맑은 고딕", 
                     font_size: int = 11, 
                     bold: bool = False, 
                     italic: bool = False, 
                     underline: bool = False,
                     color: str = "black") -> str:
    """선택된 텍스트에 글꼴 서식을 적용합니다."""
    try:
        hwp_controller.check_initialization()
        
        # 색상은 RGB로 받아 내부에서 COLORREF(BGR)로 변환한다.
        # (기존 버전은 RGB 값을 그대로 넣어 red가 파란색으로 나오는 버그가 있었음)
        _charshape_selection(hwp_controller.hwp, size_pt=font_size, bold=bold,
                             italic=italic, underline=underline,
                             color=color, face=font_name)
        
        logger.info(f"글꼴 서식 적용 완료: {font_name}, {font_size}pt")
        return f"글꼴 서식을 적용했습니다: {font_name}, {font_size}pt"
        
    except Exception as e:
        logger.error(f"글꼴 서식 적용 실패: {e}")
        return f"글꼴 서식 적용 실패: {e}"

@mcp.tool()
def select_text_range(start_pos: int, end_pos: int) -> str:
    """지정된 범위의 텍스트를 선택합니다."""
    try:
        hwp_controller.check_initialization()
        
        hwp_controller.hwp.SetPos(start_pos)
        hwp_controller.hwp.MovePos(2, end_pos - start_pos, 1)
        
        logger.info(f"텍스트 선택 완료: {start_pos} ~ {end_pos}")
        return f"텍스트를 선택했습니다: 위치 {start_pos} ~ {end_pos}"
        
    except Exception as e:
        logger.error(f"텍스트 선택 실패: {e}")
        return f"텍스트 선택 실패: {e}"

@mcp.tool()
def find_and_replace(find_text: str, replace_text: str, replace_all: bool = False) -> str:
    """텍스트를 찾아서 바꿉니다."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        if replace_all:
            # 사전 존재 확인 — 없으면 '0건 치환'을 성공으로 오보하는 것을 방지
            existed = _find_once(hwp, find_text)
            hwp.HAction.Run("Cancel")
            if not existed:
                return f"'{find_text}'를 찾을 수 없습니다."
            ok = _replace_all(hwp, find_text, replace_text)
            if ok:
                logger.info(f"찾기/바꾸기 완료: '{find_text}' -> '{replace_text}'")
                return f"'{find_text}'를 '{replace_text}'로 모두 바꾸었습니다. (잔존 없음 확인)"
            # 정확 매칭 실패 시 원인 안내 (특수문자·따옴표 차이 가능성)
            if _find_once(hwp, find_text):
                hwp.HAction.Run("Cancel")
                return (f"일부 치환에 실패했습니다: '{find_text}'가 아직 남아 있습니다. "
                        f"곡선따옴표(’ ‘) 등 특수문자 차이를 확인하거나 "
                        f"delete_paragraph_containing 도구를 사용하세요.")
            return f"'{find_text}'를 찾을 수 없습니다."
        else:
            # 단건 치환: 액션 ID('Replace'/'ExecReplace')는 버전별로 불확실 —
            # 찾기(선택 상태) + Delete + InsertText 조합이 결정적이고 안전하다
            if not _find_once(hwp, find_text):
                hwp.HAction.Run("Cancel")
                return f"'{find_text}'를 찾을 수 없습니다."
            hwp.HAction.Run("Delete")           # 선택된 찾은 텍스트 삭제
            ps = hwp.HParameterSet.HInsertText
            hwp.HAction.GetDefault("InsertText", ps.HSet)
            ps.Text = replace_text
            hwp.HAction.Execute("InsertText", ps.HSet)
            logger.info(f"찾기/바꾸기 완료: '{find_text}' -> '{replace_text}'")
            return f"첫 번째 '{find_text}'를 '{replace_text}'로 바꾸었습니다."

    except Exception as e:
        logger.error(f"찾기/바꾸기 실패: {e}")
        return f"찾기/바꾸기 실패: {e}"

@mcp.tool()
def create_table(rows: int, cols: int, border: bool = True,
                 width_mm: float = 150.0, row_height_mm: float = 5.2,
                 col_widths: Optional[str] = None,
                 data_json: Optional[str] = None,
                 header_bold: bool = True, font_size: float = 10.0) -> str:
    """커서 위치에 표를 생성합니다. (선택) 데이터까지 채우고 표 밖으로 안전하게 탈출합니다.

    - width_mm: 표 전체 폭(mm). 단 폭에 맞추되 이 값 기준으로 열 폭을 배분
    - col_widths: 열 폭 비율 JSON 배열 예: "[2,1,1]" (생략 시 균등)
    - data_json: 셀 내용 JSON 2차원 배열 예: '[["구분","값"],["A","1"]]'
      첫 행은 header_bold=True면 굵게·가운데 정렬
    주의: HWP 특성상 WidthType=0이어도 WidthValue를 반드시 지정해야 열 폭이 유지되며,
    표 입력 후 커서를 명시적으로 표 밖으로 빼지 않으면 이후 내용이 표 안에 중첩됩니다.
    이 도구는 두 문제를 모두 처리합니다."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        widths = None
        if col_widths:
            try:
                widths = json.loads(col_widths)
            except Exception:
                return f"col_widths JSON 파싱 실패: {col_widths}"
        data = None
        if data_json:
            try:
                data = json.loads(data_json)
            except Exception:
                return f"data_json JSON 파싱 실패: {data_json[:80]}"
            if len(data) > rows or any(len(r) > cols for r in data):
                return "data_json 크기가 rows/cols보다 큽니다."

        # 표 생성 전 위치 기록 (탈출용)
        L0, P0, _ = hwp.GetPos()

        act = hwp.CreateAction("TableCreate")
        pset = act.CreateSet()
        act.GetDefault(pset)
        pset.SetItem("Rows", rows)
        pset.SetItem("Cols", cols)
        pset.SetItem("WidthType", 0)                     # 단에 맞춤
        pset.SetItem("HeightType", 1)
        pset.SetItem("WidthValue", _mm(hwp, width_mm))   # WidthType=0이어도 필수!
        pset.SetItem("HeightValue", _mm(hwp, row_height_mm * rows))
        try:
            arr = pset.CreateItemArray("ColWidth", cols)
            ws = widths if widths and len(widths) == cols else [1.0] * cols
            tot = float(sum(ws))
            for i, v in enumerate(ws):
                arr.SetItem(i, _mm(hwp, width_mm * v / tot))
        except Exception:
            pass
        act.Execute(pset)

        # (선택) 데이터 채우기 — 생성 직후 커서는 첫 셀
        if data:
            for r in range(rows):
                for c in range(cols):
                    val = ""
                    if r < len(data) and c < len(data[r]):
                        val = str(data[r][c])
                    is_head = header_bold and r == 0
                    _charshape_selection(hwp, size_pt=font_size, bold=is_head)
                    try:
                        pp = hwp.HParameterSet.HParaShape
                        hwp.HAction.GetDefault("ParagraphShape", pp.HSet)
                        pp.AlignType = 3 if is_head else 1   # 헤더 가운데/본문 왼쪽
                        hwp.HAction.Execute("ParagraphShape", pp.HSet)
                    except Exception:
                        pass
                    if val:
                        tp = hwp.HParameterSet.HInsertText
                        hwp.HAction.GetDefault("InsertText", tp.HSet)
                        tp.Text = val
                        hwp.HAction.Execute("InsertText", tp.HSet)
                    if not (r == rows - 1 and c == cols - 1):
                        hwp.HAction.Run("TableRightCell")

        # 표 밖으로 확실히 탈출 (MoveDown은 커서가 표 안으로 되돌아가므로 금지)
        hwp.HAction.Run("Cancel")
        hwp.HAction.Run("CloseEx")
        cur = hwp.GetPos()
        if cur[0] != L0:
            hwp.SetPos(L0, P0 + 1, 0)
        else:
            hwp.SetPos(L0, cur[1], 0)
        hwp.HAction.Run("MoveLineEnd")

        filled = f", 데이터 {sum(len(r) for r in data)}셀 입력" if data else ""
        logger.info(f"표 생성 완료: {rows}x{cols}{filled}")
        return f"{rows}행 {cols}열 표를 생성했습니다 (폭 {width_mm}mm{filled}). 커서는 표 다음 문단에 있습니다."

    except Exception as e:
        logger.error(f"표 생성 실패: {e}")
        return f"표 생성 실패: {e}"

@mcp.tool()
def set_page_margins(top: int = 20, bottom: int = 20, left: int = 20, right: int = 20) -> str:
    """페이지 여백을 설정합니다. (단위: mm)"""
    try:
        hwp_controller.check_initialization()
        
        hwp = hwp_controller.hwp
        # 여백 항목은 파라미터셋 루트가 아니라 HSecDef.PageDef 하위에 있음
        # (루트 SetItem은 예외 없이 무시되는 silent no-op — 실기기 확인)
        sd = hwp.HParameterSet.HSecDef
        hwp.HAction.GetDefault("PageSetup", sd.HSet)
        sd.PageDef.TopMargin = _mm(hwp, top)
        sd.PageDef.BottomMargin = _mm(hwp, bottom)
        sd.PageDef.LeftMargin = _mm(hwp, left)
        sd.PageDef.RightMargin = _mm(hwp, right)
        sd.HSet.SetItem("ApplyClass", 24)
        sd.HSet.SetItem("ApplyTo", 3)
        hwp.HAction.Execute("PageSetup", sd.HSet)
        
        logger.info(f"페이지 여백 설정 완료: 상{top} 하{bottom} 좌{left} 우{right}mm")
        return f"페이지 여백을 설정했습니다: 상{top} 하{bottom} 좌{left} 우{right}mm"
        
    except Exception as e:
        logger.error(f"페이지 여백 설정 실패: {e}")
        return f"페이지 여백 설정 실패: {e}"

@mcp.tool()
def get_document_info() -> str:
    """현재 문서의 정보를 조회합니다."""
    try:
        hwp_controller.check_initialization()
        
        try:
            page_count = hwp_controller.hwp.PageCount
        except:
            page_count = "Unknown"
            
        try:
            current_pos = hwp_controller.hwp.GetPos()
        except:
            current_pos = "Unknown"
            
        try:
            list_count = getattr(hwp_controller.hwp, 'ListCount', 0)
        except:
            list_count = "Unknown"
        
        info = {
            "page_count": page_count,
            "current_pos": current_pos,
            "list_count": list_count,
            "document_name": hwp_controller.current_document or "새 문서"
        }
        
        result = f"""문서 정보:
- 문서명: {info['document_name']}
- 총 페이지 수: {info['page_count']}
- 현재 커서 위치: {info['current_pos']}
- 리스트 개수: {info['list_count']}"""
        
        logger.info("문서 정보 조회 완료")
        return result
        
    except Exception as e:
        logger.error(f"문서 정보 조회 실패: {e}")
        return f"문서 정보 조회 실패: {e}"

@mcp.tool()
def set_paragraph_format(align: str = "left", 
                        left_indent: int = 0, 
                        right_indent: int = 0, 
                        line_spacing: float = 1.0) -> str:
    """문단 서식을 설정합니다."""
    try:
        hwp_controller.check_initialization()
        
        align_map = {
            "left": 0,
            "center": 1,
            "right": 2,
            "justify": 3,
            "distribute": 4
        }
        
        align_value = align_map.get(align.lower(), 0)
        
        act = hwp_controller.hwp.CreateAction("ParagraphShape")
        pset = act.CreateSet()
        pset.SetItem("Align", align_value)
        pset.SetItem("IndentLeft", left_indent * 100)
        pset.SetItem("IndentRight", right_indent * 100)
        pset.SetItem("LineSpacing", int(line_spacing * 100))
        act.Execute(pset)
        
        logger.info(f"문단 서식 설정 완료: {align} 정렬, 줄간격 {line_spacing}")
        return f"문단 서식을 설정했습니다: {align} 정렬, 줄간격 {line_spacing}"
        
    except Exception as e:
        logger.error(f"문단 서식 설정 실패: {e}")
        return f"문단 서식 설정 실패: {e}"

@mcp.tool()
def set_page_size(width: int = 210, height: int = 297, orientation: str = "portrait") -> str:
    """용지 크기와 방향을 설정합니다. (단위: mm)"""
    try:
        hwp_controller.check_initialization()
        
        if orientation.lower() == "landscape":
            width, height = height, width
        
        hwp = hwp_controller.hwp
        # 용지 항목의 실명은 PageDef.PaperWidth/PaperHeight/Landscape (루트 Width/Height는 없음)
        sd = hwp.HParameterSet.HSecDef
        hwp.HAction.GetDefault("PageSetup", sd.HSet)
        sd.PageDef.PaperWidth = _mm(hwp, width)
        sd.PageDef.PaperHeight = _mm(hwp, height)
        sd.PageDef.Landscape = 1 if orientation.lower() == "landscape" else 0
        sd.HSet.SetItem("ApplyClass", 24)
        sd.HSet.SetItem("ApplyTo", 3)
        hwp.HAction.Execute("PageSetup", sd.HSet)
        
        logger.info(f"용지 설정 완료: {width}x{height}mm, {orientation}")
        return f"용지를 설정했습니다: {width}x{height}mm, {orientation}"
        
    except Exception as e:
        logger.error(f"용지 설정 실패: {e}")
        return f"용지 설정 실패: {e}"

@mcp.tool()
def insert_image(image_path: str, x: int = 0, y: int = 0,
                 width: float = 100, height: float = 0) -> str:
    """커서 위치에 이미지를 삽입합니다. width/height는 mm 단위.
    height=0이면 이미지 원본 비율로 자동 계산. 삽입 후 커서는 이미지 뒤로 이동합니다."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        image_path = os.path.abspath(image_path)
        if not os.path.exists(image_path):
            return f"이미지 파일을 찾을 수 없습니다: {image_path}"

        if not height:
            try:
                from PIL import Image as _PILImage
                with _PILImage.open(image_path) as im:
                    height = round(width * im.height / im.width, 1)
            except Exception:
                height = width * 0.6      # PIL 없으면 3:5 근사

        # InsertPicture의 sizeoption은 버전별 해석이 달라 신뢰할 수 없음 —
        # 삽입 후 반환된 컨트롤의 Properties로 크기를 강제 지정한다(검증된 경로).
        ctrl = hwp.InsertPicture(image_path, True, 2, False, False, 0,
                                 _mm(hwp, width), _mm(hwp, height))
        try:
            p = ctrl.Properties
            p.SetItem("Width", _mm(hwp, width))
            p.SetItem("Height", _mm(hwp, height))
            ctrl.Properties = p
        except Exception:
            pass
        hwp.HAction.Run("Cancel")        # 개체 선택 해제
        hwp.HAction.Run("MoveRight")     # 이미지 뒤로

        logger.info(f"이미지 삽입 완료: {image_path} ({width}x{height}mm)")
        return f"이미지를 삽입했습니다: {os.path.basename(image_path)} ({width}x{height}mm)"

    except Exception as e:
        logger.error(f"이미지 삽입 실패: {e}")
        return f"이미지 삽입 실패: {e}"

@mcp.tool()
def insert_shape(shape_type: str, x: int = 0, y: int = 0, width: int = 50, height: int = 50) -> str:
    """도형을 삽입합니다."""
    try:
        hwp_controller.check_initialization()
        
        shape_map = {
            "rectangle": 1,
            "ellipse": 2,
            "line": 3,
            "arrow": 4,
            "textbox": 5
        }
        
        shape_value = shape_map.get(shape_type.lower(), 1)
        
        act = hwp_controller.hwp.CreateAction("DrawObjDialog")
        pset = act.CreateSet()
        pset.SetItem("ShapeType", shape_value)
        pset.SetItem("TreatAsChar", False)
        act.Execute(pset)
        
        logger.info(f"도형 삽입 완료: {shape_type}")
        return f"도형을 삽입했습니다: {shape_type}"
        
    except Exception as e:
        logger.error(f"도형 삽입 실패: {e}")
        return f"도형 삽입 실패: {e}"

@mcp.tool()
def insert_header_footer(text: str, is_header: bool = True, position: str = "center") -> str:
    """머리글 또는 바닥글을 삽입합니다."""
    try:
        hwp_controller.check_initialization()
        
        if is_header:
            hwp_controller.hwp.HAction.Run("HeaderFooterEdit")
        else:
            hwp_controller.hwp.HAction.Run("HeaderFooterEdit")
        
        act = hwp_controller.hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        pset.SetItem("Text", text)
        act.Execute(pset)
        
        hwp_controller.hwp.HAction.Run("CloseEx")
        
        logger.info(f"{'머리글' if is_header else '바닥글'} 삽입 완료")
        return f"{'머리글' if is_header else '바닥글'}을 삽입했습니다: {text}"
        
    except Exception as e:
        logger.error(f"머리글/바닥글 삽입 실패: {e}")
        return f"머리글/바닥글 삽입 실패: {e}"

@mcp.tool()
def insert_page_break() -> str:
    """페이지 나누기를 삽입합니다."""
    try:
        hwp_controller.check_initialization()
        
        act = hwp_controller.hwp.CreateAction("BreakPage")
        act.Execute()
        
        logger.info("페이지 나누기 삽입 완료")
        return "페이지 나누기를 삽입했습니다."
        
    except Exception as e:
        logger.error(f"페이지 나누기 삽입 실패: {e}")
        return f"페이지 나누기 삽입 실패: {e}"

@mcp.tool()
def merge_table_cells(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    """표의 셀을 병합합니다."""
    try:
        hwp_controller.check_initialization()
        
        hwp_controller.hwp.TableCellBlock(start_row, start_col, end_row, end_col)
        
        act = hwp_controller.hwp.CreateAction("TableMergeCell")
        act.Execute()
        
        logger.info(f"셀 병합 완료: ({start_row},{start_col}) ~ ({end_row},{end_col})")
        return f"셀을 병합했습니다: ({start_row},{start_col}) ~ ({end_row},{end_col})"
        
    except Exception as e:
        logger.error(f"셀 병합 실패: {e}")
        return f"셀 병합 실패: {e}"

@mcp.tool()
def insert_hyperlink(text: str, url: str) -> str:
    """하이퍼링크를 삽입합니다."""
    try:
        hwp_controller.check_initialization()
        
        act = hwp_controller.hwp.CreateAction("InsertHyperlink")
        pset = act.CreateSet()
        pset.SetItem("Text", text)
        pset.SetItem("URL", url)
        act.Execute(pset)
        
        logger.info(f"하이퍼링크 삽입 완료: {text} -> {url}")
        return f"하이퍼링크를 삽입했습니다: {text} -> {url}"
        
    except Exception as e:
        logger.error(f"하이퍼링크 삽입 실패: {e}")
        return f"하이퍼링크 삽입 실패: {e}"

@mcp.tool()
def create_table_of_contents() -> str:
    """목차를 생성합니다."""
    try:
        hwp_controller.check_initialization()
        
        act = hwp_controller.hwp.CreateAction("InsertTableOfContents")
        pset = act.CreateSet()
        pset.SetItem("AutoUpdate", True)
        pset.SetItem("ShowPageNum", True)
        act.Execute(pset)
        
        logger.info("목차 생성 완료")
        return "목차를 생성했습니다."
        
    except Exception as e:
        logger.error(f"목차 생성 실패: {e}")
        return f"목차 생성 실패: {e}"

@mcp.tool()
def apply_heading_style(level: int, text: str) -> str:
    """제목 스타일을 적용합니다."""
    try:
        hwp_controller.check_initialization()
        
        act = hwp_controller.hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        pset.SetItem("Text", text)
        act.Execute(pset)
        
        style_name = f"제목 {level}"
        act = hwp_controller.hwp.CreateAction("StyleApply")
        pset = act.CreateSet()
        pset.SetItem("StyleName", style_name)
        act.Execute(pset)
        
        logger.info(f"제목 스타일 적용 완료: {style_name}")
        return f"제목 스타일을 적용했습니다: {style_name}"
        
    except Exception as e:
        logger.error(f"제목 스타일 적용 실패: {e}")
        return f"제목 스타일 적용 실패: {e}"

@mcp.tool()
def export_to_pdf(output_path: str) -> str:
    """현재 문서를 PDF로 내보냅니다."""
    try:
        hwp_controller.check_initialization()
        
        hwp = hwp_controller.hwp
        output_path = os.path.abspath(output_path)

        ok = False
        try:
            ok = bool(hwp.SaveAs(output_path, "PDF", ""))
        except Exception as e1:
            logger.warning(f"SaveAs PDF 실패, 액션 방식 재시도: {e1}")
            try:
                act = hwp.CreateAction("FileSaveAsPdf")
                pset = act.CreateSet()
                act.GetDefault(pset)
                pset.SetItem("filename", output_path)
                pset.SetItem("Format", "PDF")
                act.Execute(pset)
                ok = True
            except Exception as e2:
                return f"PDF 내보내기 실패: {e2}"

        if os.path.exists(output_path):
            logger.info(f"PDF 내보내기 완료: {output_path}")
            return f"PDF로 내보냈습니다: {output_path}"
        return f"PDF 내보내기가 실행됐지만 파일이 확인되지 않습니다: {output_path}"
        
    except Exception as e:
        logger.error(f"PDF 내보내기 실패: {e}")
        return f"PDF 내보내기 실패: {e}"

@mcp.tool()
def get_text_all() -> str:
    """문서 전체의 텍스트를 읽어옵니다."""
    try:
        hwp_controller.check_initialization()

        text = _extract_text(hwp_controller.hwp, "document")

        logger.info(f"전체 텍스트 읽기 완료: {len(text)} 글자")
        return text if text else "(문서가 비어 있습니다)"
        
    except Exception as e:
        logger.error(f"텍스트 읽기 실패: {e}")
        return f"텍스트 읽기 실패: {e}"

@mcp.tool()
def get_text_by_page(page_number: int) -> str:
    """특정 페이지의 텍스트를 읽어옵니다."""
    try:
        hwp_controller.check_initialization()

        hwp = hwp_controller.hwp
        total_pages = hwp.PageCount
        if page_number < 1 or page_number > total_pages:
            return f"페이지 번호가 범위를 벗어났습니다. (1~{total_pages})"
        # HGotoE에는 PageNumber 속성이 없음(기존 버그) — DialogResult+SetSelectionIndex 사용
        ps = hwp.HParameterSet.HGotoE
        hwp.HAction.GetDefault("Goto", ps.HSet)
        ps.HSet.SetItem("DialogResult", page_number)
        ps.SetSelectionIndex = 1          # 1 = 페이지 기준 이동
        hwp.HAction.Execute("Goto", ps.HSet)

        hwp.HAction.Run("MovePageBegin")
        hwp.HAction.Run("MoveSelPageDown")

        text = _extract_text(hwp, "selection")
        hwp.HAction.Run("Cancel")

        logger.info(f"{page_number}페이지 텍스트 읽기 완료")
        return text
        
    except Exception as e:
        logger.error(f"페이지 텍스트 읽기 실패: {e}")
        return f"페이지 텍스트 읽기 실패: {e}"

@mcp.tool()
def get_selected_text() -> str:
    """현재 선택된 텍스트를 읽어옵니다."""
    try:
        hwp_controller.check_initialization()

        text = _fix_mojibake(hwp_controller.hwp.GetTextFile("TEXT", "") or "")

        if text is None:
            text = ""

        logger.info(f"선택된 텍스트 읽기 완료: {len(text)} 글자")
        return text
        
    except Exception as e:
        logger.error(f"선택된 텍스트 읽기 실패: {e}")
        return f"선택된 텍스트 읽기 실패: {e}"

@mcp.tool()
def get_paragraph_text(paragraph_index: int = 0) -> str:
    """특정 문단의 텍스트를 읽어옵니다. (0부터 시작)"""
    try:
        hwp_controller.check_initialization()

        hwp_controller.hwp.HAction.Run("MoveDocBegin")

        for _ in range(paragraph_index):
            hwp_controller.hwp.HAction.Run("MoveParaDown")

        hwp_controller.hwp.HAction.Run("MoveSelParaDown")
        text = _fix_mojibake(hwp_controller.hwp.GetTextFile("TEXT", "") or "")
        hwp_controller.hwp.HAction.Run("Cancel")

        if text is None:
            text = ""

        logger.info(f"{paragraph_index}번째 문단 텍스트 읽기 완료")
        return text.strip()
        
    except Exception as e:
        logger.error(f"문단 텍스트 읽기 실패: {e}")
        return f"문단 텍스트 읽기 실패: {e}"

@mcp.tool()
def save_as_text(output_path: str) -> str:
    """문서 전체를 텍스트 파일로 저장합니다."""
    try:
        hwp_controller.check_initialization()
        
        hwp_controller.hwp.SaveAs(output_path, "TEXT")
        
        logger.info(f"텍스트 파일로 저장 완료: {output_path}")
        return f"텍스트 파일로 저장했습니다: {output_path}"
        
    except Exception as e:
        logger.error(f"텍스트 저장 실패: {e}")
        return f"텍스트 저장 실패: {e}"


# ============================================================
# 고급 분석 및 자동화 기능
# ============================================================

@mcp.tool()
def get_table_as_csv(table_index: int = 1, output_path: Optional[str] = None) -> str:
    """
    특정 표를 CSV 형식으로 추출합니다. (1부터 시작)
    output_path를 지정하면 파일로 저장, 아니면 텍스트로 반환합니다.
    """
    try:
        hwp_controller.check_initialization()
        
        hwp = hwp_controller.hwp
        current_table = 0
        
        hwp.HAction.Run("MoveDocBegin")
        
        ctrl = hwp.HeadCtrl
        target_ctrl = None
        
        while ctrl:
            try:
                if ctrl.CtrlID == 'tbl':
                    current_table += 1
                    if current_table == table_index:
                        target_ctrl = ctrl
                        break
                ctrl = ctrl.Next
            except:
                break
        
        if not target_ctrl:
            return f"{table_index}번째 표를 찾을 수 없습니다. (총 {current_table}개 표 존재)"
        
        rows = 0
        cols = 0
        try:
            tbl_set = target_ctrl.Properties
            rows = tbl_set.Item("RowCount")
            cols = tbl_set.Item("ColCount")
        except:
            pass
        
        cell_contents = []
        try:
            hwp.SetPosBySet(target_ctrl.GetAnchorPos(0))
            hwp.HAction.Run("ShapeObjTableSelCell")
            hwp.HAction.Run("TableColBegin")
            hwp.HAction.Run("TableRowBegin")
            hwp.HAction.Run("TableCellBlockExtendAll")
            
            table_text = _fix_mojibake(hwp.GetTextFile("TEXT", "") or "")

            if table_text is None:
                table_text = ""

            if table_text:
                cell_contents = [t.strip() for t in table_text.split('\r\n') if t.strip()]
            
            hwp.HAction.Run("Cancel")
            
        except Exception as e:
            logger.warning(f"표 내용 추출 중 오류: {e}")
            return f"표 내용 추출 실패: {e}"
        
        csv_lines = []
        if rows > 0 and cols > 0 and len(cell_contents) >= rows * cols:
            for r in range(rows):
                row_data = []
                for c in range(cols):
                    idx = r * cols + c
                    if idx < len(cell_contents):
                        cell = cell_contents[idx].replace('"', '""')
                        if ',' in cell or '"' in cell or '\n' in cell:
                            cell = f'"{cell}"'
                        row_data.append(cell)
                csv_lines.append(','.join(row_data))
        else:
            for cell in cell_contents:
                cell = cell.replace('"', '""')
                if ',' in cell or '"' in cell:
                    cell = f'"{cell}"'
                csv_lines.append(cell)
        
        csv_content = '\n'.join(csv_lines)
        
        if output_path:
            with open(output_path, 'w', encoding='utf-8-sig') as f:
                f.write(csv_content)
            logger.info(f"표 {table_index} CSV 저장 완료: {output_path}")
            return f"표 {table_index}을(를) CSV로 저장했습니다: {output_path}\n({rows}행 x {cols}열, {len(cell_contents)}개 셀)"
        else:
            logger.info(f"표 {table_index} CSV 추출 완료")
            return f"표 {table_index} CSV 내용 ({rows}행 x {cols}열):\n\n{csv_content}"
        
    except Exception as e:
        logger.error(f"표 CSV 추출 실패: {e}")
        return f"표 CSV 추출 실패: {e}"


@mcp.tool()
def batch_replace(replacements: str) -> str:
    """
    여러 텍스트를 한번에 바꿉니다.
    replacements: "찾을텍스트1->바꿀텍스트1, 찾을텍스트2->바꿀텍스트2" 형식
    예: "주식회사->㈜, 2023년->2024년, 홍길동->김철수"
    """
    try:
        hwp_controller.check_initialization()
        
        hwp = hwp_controller.hwp
        
        pairs = [p.strip() for p in replacements.split(',')]
        results = []
        total_replaced = 0
        
        for pair in pairs:
            if '->' not in pair:
                results.append(f"! '{pair}': 잘못된 형식 (->로 구분 필요)")
                continue
            
            parts = pair.split('->', 1)
            find_text = parts[0].strip()
            replace_text = parts[1].strip()
            
            if not find_text:
                results.append(f"! 빈 검색어는 건너뜁니다")
                continue
            
            hwp.HAction.Run("MoveDocBegin")

            hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
            hwp.HParameterSet.HFindReplace.FindString = find_text
            hwp.HParameterSet.HFindReplace.ReplaceString = replace_text
            hwp.HParameterSet.HFindReplace.ReplaceMode = 1

            execute_result = hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
            
            if execute_result:
                results.append(f"O '{find_text}' -> '{replace_text}': 완료")
                total_replaced += 1
            else:
                results.append(f"- '{find_text}': 찾을 수 없음")
        
        result = f"""일괄 바꾸기 결과
==============================
총 {len(pairs)}개 항목 중 {total_replaced}개 처리 완료

"""
        result += '\n'.join(results)
        
        logger.info(f"일괄 바꾸기 완료: {total_replaced}/{len(pairs)}개")
        return result
        
    except Exception as e:
        logger.error(f"일괄 바꾸기 실패: {e}")
        return f"일괄 바꾸기 실패: {e}"


@mcp.tool()
def find_text(search_text: str, show_context: bool = True) -> str:
    """
    문서에서 특정 텍스트를 찾아 위치와 주변 내용을 반환합니다.
    search_text: 찾을 텍스트
    show_context: True면 주변 텍스트도 함께 표시
    """
    try:
        hwp_controller.check_initialization()
        
        hwp = hwp_controller.hwp
        
        full_text = _extract_text(hwp, "document")

        if not full_text:
            return "문서에 내용이 없습니다."
        
        search_lower = search_text.lower()
        text_lower = full_text.lower()
        
        positions = []
        start = 0
        while True:
            pos = text_lower.find(search_lower, start)
            if pos == -1:
                break
            positions.append(pos)
            start = pos + 1
        
        if not positions:
            return f"'{search_text}'을(를) 찾을 수 없습니다."
        
        lines = full_text.split('\r\n')
        results = []
        
        for i, pos in enumerate(positions, 1):
            current_pos = 0
            line_num = 0
            
            for idx, line in enumerate(lines):
                line_end = current_pos + len(line)
                if current_pos <= pos < line_end + 2:
                    line_num = idx + 1
                    break
                current_pos = line_end + 2
            
            estimated_page = (line_num // 50) + 1
            
            if show_context:
                context_start = max(0, pos - 50)
                context_end = min(len(full_text), pos + len(search_text) + 50)
                context = full_text[context_start:context_end].replace('\r\n', ' ')
                
                highlight_pos = pos - context_start
                context_display = (
                    context[:highlight_pos] + 
                    f"[{search_text}]" + 
                    context[highlight_pos + len(search_text):]
                )
                
                results.append(f"  [{i}] {line_num}번째 줄 (약 {estimated_page}페이지)\n      ...{context_display}...")
            else:
                results.append(f"  [{i}] {line_num}번째 줄 (약 {estimated_page}페이지)")
        
        result = f"""검색 결과: '{search_text}'
==============================
총 {len(positions)}개 발견

"""
        result += '\n\n'.join(results)
        
        logger.info(f"텍스트 검색 완료: '{search_text}' - {len(positions)}개 발견")
        return result
        
    except Exception as e:
        logger.error(f"텍스트 검색 실패: {e}")
        return f"텍스트 검색 실패: {e}"


@mcp.tool()
def fill_template(field_values: str) -> str:
    """
    문서의 필드(플레이스홀더)를 값으로 채웁니다.
    필드는 {{필드명}} 또는 {필드명} 형식으로 문서에 있어야 합니다.
    
    field_values: "필드명1=값1, 필드명2=값2" 형식
    예: "이름=홍길동, 날짜=2024-01-01, 금액=1,000,000원"
    """
    try:
        hwp_controller.check_initialization()
        
        hwp = hwp_controller.hwp
        
        pairs = [p.strip() for p in field_values.split(',')]
        results = []
        total_filled = 0
        
        for pair in pairs:
            if '=' not in pair:
                results.append(f"! '{pair}': 잘못된 형식 (=로 구분 필요)")
                continue
            
            parts = pair.split('=', 1)
            field_name = parts[0].strip()
            field_value = parts[1].strip()
            
            if not field_name:
                results.append(f"! 빈 필드명은 건너뜁니다")
                continue
            
            placeholders = [
                "{{" + field_name + "}}",
                "{" + field_name + "}",
                "[" + field_name + "]",
                "<" + field_name + ">",
                "$" + field_name + "$",
                field_name
            ]
            
            replaced = False
            for placeholder in placeholders:
                if not _find_once(hwp, placeholder):
                    hwp.HAction.Run("Cancel")
                    continue
                hwp.HAction.Run("Cancel")
                if _replace_all(hwp, placeholder, field_value):
                    results.append(f"O {placeholder} -> '{field_value}': 완료")
                    total_filled += 1
                    replaced = True
                    break
            
            if not replaced:
                results.append(f"- '{field_name}': 해당 필드를 찾을 수 없음")
        
        result = f"""템플릿 채우기 결과
==============================
총 {len(pairs)}개 필드 중 {total_filled}개 채움 완료

"""
        result += '\n'.join(results)
        
        if total_filled < len(pairs):
            result += "\n\n팁: 문서에서 필드는 {{이름}}, {이름}, [이름] 등의 형식으로 작성해주세요."
        
        logger.info(f"템플릿 채우기 완료: {total_filled}/{len(pairs)}개")
        return result
        
    except Exception as e:
        logger.error(f"템플릿 채우기 실패: {e}")
        return f"템플릿 채우기 실패: {e}"


@mcp.tool()
def get_document_structure() -> str:
    """
    문서의 전체 구조를 분석합니다. (페이지 수, 문단 수, 표 개수, 이미지 개수, 제목/개요 구조)
    """
    try:
        hwp_controller.check_initialization()
        
        hwp = hwp_controller.hwp
        
        page_count = hwp.PageCount
        
        doc_path = "(새 문서)"
        try:
            if hwp.XHwpDocuments.Count > 0:
                active_doc = hwp.XHwpDocuments.Item(0)
                if active_doc.Path:
                    doc_path = active_doc.Path
        except:
            pass
        
        doc_name = doc_path.split("\\")[-1] if doc_path != "(새 문서)" else "새 문서"
        
        full_text = _extract_text(hwp, "document")

        char_count = len(full_text.replace('\r\n', '').replace(' ', '')) if full_text else 0
        paragraph_count = full_text.count('\r\n') + 1 if full_text else 0
        
        table_count = 0
        image_count = 0
        shape_count = 0
        
        ctrl = hwp.HeadCtrl
        while ctrl:
            try:
                ctrl_code = ctrl.CtrlID
                if ctrl_code == 'tbl':
                    table_count += 1
                elif ctrl_code in ['ole', 'pic']:
                    image_count += 1
                elif ctrl_code == 'gso':
                    shape_count += 1
                ctrl = ctrl.Next
            except:
                break
        
        headings = []
        lines = full_text.split('\r\n') if full_text else []
        
        heading_patterns = [
            (r'^\s*([IVX]+)\.\s*(.+)$', 1, "로마자"),
            (r'^\s*(\d+)\.\s*(.+)$', 2, "숫자"),
            (r'^\s*([가-힣])\.\s*(.+)$', 3, "가나다"),
            (r'^\s*\((\d+)\)\s*(.+)$', 3, "괄호숫자"),
            (r'^\s*(제\s*\d+\s*[장절조항])\s*(.*)$', 1, "장절"),
            (r'^\s*(붙임|별첨|부록)\s*[\d]*\.?\s*(.*)$', 1, "붙임"),
        ]
        
        for i, line in enumerate(lines[:200]):
            line_stripped = line.strip()
            if not line_stripped or len(line_stripped) > 100:
                continue
            
            for pattern, level, ptype in heading_patterns:
                if re.match(pattern, line_stripped):
                    headings.append({
                        "line": i + 1,
                        "level": level,
                        "text": line_stripped[:60] + ("..." if len(line_stripped) > 60 else ""),
                        "type": ptype
                    })
                    break
        
        result = f"""문서 구조 분석 결과
{'='*50}

기본 정보:
  - 파일명: {doc_name}
  - 총 페이지: {page_count}페이지
  - 총 문단: {paragraph_count}개 (추정)
  - 총 글자: {char_count:,}자 (공백 제외)

포함된 요소:
  - 표(테이블): {table_count}개
  - 이미지/그림: {image_count}개
  - 도형: {shape_count}개
"""
        
        if headings:
            result += f"\n문서 개요 구조 ({len(headings)}개 항목):\n"
            for h in headings[:20]:
                indent = "  " * h['level']
                result += f"{indent}- {h['text']}\n"
            if len(headings) > 20:
                result += f"  ... 외 {len(headings) - 20}개 항목\n"
        else:
            result += "\n문서 개요: (번호 체계 없음)\n"
        
        result += f"\n{'='*50}"
        
        logger.info(f"문서 구조 분석 완료: {doc_name}")
        return result
        
    except Exception as e:
        logger.error(f"문서 구조 분석 실패: {e}")
        return f"문서 구조 분석 실패: {e}"


# ============================================================
# 개선된 위치 제어 및 편집 기능
# ============================================================

@mcp.tool()
def move_to_page(page_number: int) -> str:
    """
    특정 페이지로 커서를 이동합니다.
    page_number: 이동할 페이지 번호 (1부터 시작)
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        total_pages = hwp.PageCount
        if page_number < 1 or page_number > total_pages:
            return f"페이지 번호가 범위를 벗어났습니다. (1~{total_pages})"

        ps = hwp.HParameterSet.HGotoE
        hwp.HAction.GetDefault("Goto", ps.HSet)
        ps.HSet.SetItem("DialogResult", page_number)
        ps.SetSelectionIndex = 1          # 1 = 페이지 기준 이동
        hwp.HAction.Execute("Goto", ps.HSet)

        logger.info(f"{page_number}페이지로 이동 완료")
        return f"{page_number}페이지로 이동했습니다."

    except Exception as e:
        logger.error(f"페이지 이동 실패: {e}")
        return f"페이지 이동 실패: {e}"


@mcp.tool()
def move_to_paragraph_number(paragraph_number: int) -> str:
    """
    특정 문단으로 커서를 이동합니다. (0부터 시작)
    paragraph_number: 이동할 문단 번호
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        hwp.HAction.Run("MoveDocBegin")

        for i in range(paragraph_number):
            hwp.HAction.Run("MoveParaDown")

        logger.info(f"{paragraph_number}번째 문단으로 이동 완료")
        return f"{paragraph_number}번째 문단으로 이동했습니다."

    except Exception as e:
        logger.error(f"문단 이동 실패: {e}")
        return f"문단 이동 실패: {e}"


@mcp.tool()
def move_to_document_end() -> str:
    """문서의 끝으로 커서를 이동합니다."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        hwp.HAction.Run("MoveDocEnd")

        logger.info("문서 끝으로 이동 완료")
        return "문서 끝으로 이동했습니다."

    except Exception as e:
        logger.error(f"문서 끝 이동 실패: {e}")
        return f"문서 끝 이동 실패: {e}"


@mcp.tool()
def move_to_document_start() -> str:
    """문서의 시작으로 커서를 이동합니다."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        hwp.HAction.Run("MoveDocBegin")

        logger.info("문서 시작으로 이동 완료")
        return "문서 시작으로 이동했습니다."

    except Exception as e:
        logger.error(f"문서 시작 이동 실패: {e}")
        return f"문서 시작 이동 실패: {e}"


# ============================================================
# 텍스트 삭제 기능
# ============================================================

@mcp.tool()
def delete_selected_text() -> str:
    """현재 선택된 텍스트를 삭제합니다."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        # 선택된 텍스트 확인
        selected_text = _fix_mojibake(hwp.GetTextFile("TEXT", "") or "")
        if selected_text is None:
            selected_text = ""

        if not selected_text:
            return "선택된 텍스트가 없습니다."

        # Delete 키 실행
        hwp.HAction.Run("Delete")

        logger.info(f"선택된 텍스트 삭제 완료: {len(selected_text)}자")
        return f"선택된 텍스트를 삭제했습니다. ({len(selected_text)}자)"

    except Exception as e:
        logger.error(f"선택 텍스트 삭제 실패: {e}")
        return f"선택 텍스트 삭제 실패: {e}"


@mcp.tool()
def delete_all_occurrences(text: str) -> str:
    """
    문서에서 특정 텍스트를 모두 찾아서 삭제합니다.
    text: 삭제할 텍스트
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        if not text:
            return "삭제할 텍스트를 지정해주세요."

        existed = _find_once(hwp, text)
        hwp.HAction.Run("Cancel")
        if not existed:
            return f"'{text}'를 찾을 수 없습니다."

        if _replace_all(hwp, text, ""):
            logger.info(f"'{text}' 모두 삭제 완료")
            return f"'{text}'를 모두 삭제했습니다. (잔존 없음 확인)"
        return (f"일부 삭제에 실패했습니다: '{text}'가 아직 남아 있습니다. "
                f"특수문자 차이가 있으면 delete_paragraph_containing을 사용하세요.")

    except Exception as e:
        logger.error(f"텍스트 삭제 실패: {e}")
        return f"텍스트 삭제 실패: {e}"


@mcp.tool()
def delete_current_line() -> str:
    """현재 커서가 있는 줄 전체를 삭제합니다."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        # 줄 시작으로 이동
        hwp.HAction.Run("MoveLineBegin")
        # 줄 끝까지 선택
        hwp.HAction.Run("MoveSelLineEnd")
        # 삭제
        hwp.HAction.Run("Delete")
        # 줄바꿈도 삭제 (다음 줄과 합쳐짐)
        hwp.HAction.Run("Delete")

        logger.info("현재 줄 삭제 완료")
        return f"현재 줄을 삭제했습니다."

    except Exception as e:
        logger.error(f"줄 삭제 실패: {e}")
        return f"줄 삭제 실패: {e}"


@mcp.tool()
def delete_current_paragraph() -> str:
    """현재 커서가 있는 문단 전체를 삭제합니다."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        # 문단 선택
        hwp.HAction.Run("MoveSelParaDown")
        # 삭제
        hwp.HAction.Run("Delete")

        logger.info("현재 문단 삭제 완료")
        return "현재 문단을 삭제했습니다."

    except Exception as e:
        logger.error(f"문단 삭제 실패: {e}")
        return f"문단 삭제 실패: {e}"


@mcp.tool()
def delete_page_content(page_number: int) -> str:
    """
    특정 페이지의 모든 내용을 삭제합니다.
    page_number: 삭제할 페이지 번호 (1부터 시작)
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        total_pages = hwp.PageCount
        if page_number < 1 or page_number > total_pages:
            return f"페이지 번호가 범위를 벗어났습니다. (1~{total_pages})"

        # 페이지로 이동
        hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
        hwp.HParameterSet.HGotoE.PageNumber = page_number
        hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)

        # 페이지 시작으로 이동
        hwp.HAction.Run("MovePageBegin")
        # 페이지 전체 선택
        hwp.HAction.Run("MoveSelPageDown")
        # 삭제
        hwp.HAction.Run("Delete")

        logger.info(f"{page_number}페이지 내용 삭제 완료")
        return f"{page_number}페이지의 내용을 삭제했습니다."

    except Exception as e:
        logger.error(f"페이지 삭제 실패: {e}")
        return f"페이지 삭제 실패: {e}"


# ============================================================
# 서식 유지 및 가져오기 기능
# ============================================================

@mcp.tool()
def get_current_char_shape() -> str:
    """
    현재 커서 위치의 글자 서식 정보를 가져옵니다.
    (글꼴, 크기, 굵기, 기울임, 밑줄, 색상 등)
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        # CharShape 정보 가져오기
        pset = hwp.HParameterSet.HCharShape
        hwp.HAction.GetDefault("CharShape", pset.HSet)

        # 서식 정보 추출 (직접 속성 접근)
        try:
            font_name = pset.FaceNameHangul if hasattr(pset, 'FaceNameHangul') and pset.FaceNameHangul else "알 수 없음"
        except:
            font_name = "알 수 없음"

        try:
            font_size = pset.Height // 100 if hasattr(pset, 'Height') and pset.Height else 0
        except:
            font_size = 0

        try:
            is_bold = bool(pset.Bold) if hasattr(pset, 'Bold') else False
        except:
            is_bold = False

        try:
            is_italic = bool(pset.Italic) if hasattr(pset, 'Italic') else False
        except:
            is_italic = False

        try:
            is_underline = bool(pset.Underline) if hasattr(pset, 'Underline') else False
        except:
            is_underline = False

        try:
            text_color = pset.TextColor if hasattr(pset, 'TextColor') else 0x000000
        except:
            text_color = 0x000000

        # 색상을 RGB로 변환
        r = text_color & 0xFF
        g = (text_color >> 8) & 0xFF
        b = (text_color >> 16) & 0xFF

        result = f"""현재 위치의 글자 서식:
- 글꼴: {font_name}
- 크기: {font_size}pt
- 굵게: {'예' if is_bold else '아니오'}
- 기울임: {'예' if is_italic else '아니오'}
- 밑줄: {'예' if is_underline else '아니오'}
- 색상: RGB({r}, {g}, {b})"""

        logger.info("글자 서식 정보 조회 완료")
        return result

    except Exception as e:
        logger.error(f"서식 정보 조회 실패: {e}")
        return f"서식 정보 조회 실패: {e}"


@mcp.tool()
def insert_text_preserving_format(text: str) -> str:
    """
    현재 위치의 서식을 유지하면서 텍스트를 삽입합니다.
    한글의 기본 동작이 서식을 유지하므로, 단순 삽입으로 동작합니다.
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        # 한글은 기본적으로 현재 위치의 서식을 유지하면서 텍스트를 삽입함
        act = hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        pset.SetItem("Text", text)
        act.Execute(pset)

        logger.info(f"서식 유지 텍스트 삽입 완료: {text[:50]}...")
        return f"서식을 유지하면서 텍스트를 삽입했습니다: {text[:50]}..."

    except Exception as e:
        logger.error(f"서식 유지 삽입 실패: {e}")
        return f"서식 유지 삽입 실패: {e}"


# ============================================================
# 고급 삽입 기능 (특정 위치에 삽입)
# ============================================================

@mcp.tool()
def insert_after_text(search_text: str, new_text: str, nth_occurrence: int = 1) -> str:
    """
    특정 텍스트를 찾아서 그 뒤에 새 텍스트를 삽입합니다.
    search_text: 찾을 텍스트
    new_text: 삽입할 텍스트
    nth_occurrence: 몇 번째 발견된 텍스트 뒤에 삽입할지 (1부터 시작, 기본값 1)
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        if not search_text:
            return "찾을 텍스트를 지정해주세요."

        # 문서 시작으로 이동
        hwp.HAction.Run("MoveDocBegin")

        # n번째 발견까지 반복 검색
        found_count = 0
        for i in range(nth_occurrence):
            hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
            hwp.HParameterSet.HFindReplace.FindString = search_text

            result = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)

            if not result:
                if i == 0:
                    return f"'{search_text}'를 찾을 수 없습니다."
                else:
                    return f"'{search_text}'를 {i}번만 찾았습니다. ({nth_occurrence}번째를 요청했으나 존재하지 않음)"

            found_count = i + 1

        # 찾은 텍스트의 끝으로 이동 (선택 영역의 오른쪽)
        hwp.HAction.Run("MoveSelRight")

        # 새 텍스트 삽입
        act = hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        pset.SetItem("Text", new_text)
        act.Execute(pset)

        logger.info(f"'{search_text}' ({nth_occurrence}번째) 뒤에 '{new_text}' 삽입 완료")
        return f"'{search_text}' ({nth_occurrence}번째) 뒤에 '{new_text}'를 삽입했습니다."

    except Exception as e:
        logger.error(f"텍스트 뒤 삽입 실패: {e}")
        return f"텍스트 뒤 삽입 실패: {e}"


@mcp.tool()
def insert_before_text(search_text: str, new_text: str, nth_occurrence: int = 1) -> str:
    """
    특정 텍스트를 찾아서 그 앞에 새 텍스트를 삽입합니다.
    search_text: 찾을 텍스트
    new_text: 삽입할 텍스트
    nth_occurrence: 몇 번째 발견된 텍스트 앞에 삽입할지 (1부터 시작, 기본값 1)
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        if not search_text:
            return "찾을 텍스트를 지정해주세요."

        # 문서 시작으로 이동
        hwp.HAction.Run("MoveDocBegin")

        # n번째 발견까지 반복 검색
        found_count = 0
        for i in range(nth_occurrence):
            hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
            hwp.HParameterSet.HFindReplace.FindString = search_text

            result = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)

            if not result:
                if i == 0:
                    return f"'{search_text}'를 찾을 수 없습니다."
                else:
                    return f"'{search_text}'를 {i}번만 찾았습니다. ({nth_occurrence}번째를 요청했으나 존재하지 않음)"

            found_count = i + 1

        # 찾은 텍스트의 시작으로 이동 (선택 영역의 왼쪽)
        hwp.HAction.Run("MoveSelLeft")

        # 새 텍스트 삽입
        act = hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        pset.SetItem("Text", new_text)
        act.Execute(pset)

        logger.info(f"'{search_text}' ({nth_occurrence}번째) 앞에 '{new_text}' 삽입 완료")
        return f"'{search_text}' ({nth_occurrence}번째) 앞에 '{new_text}'를 삽입했습니다."

    except Exception as e:
        logger.error(f"텍스트 앞 삽입 실패: {e}")
        return f"텍스트 앞 삽입 실패: {e}"


@mcp.tool()
def append_to_paragraph(paragraph_number: int, text: str) -> str:
    """
    특정 문단의 끝에 텍스트를 추가합니다.
    paragraph_number: 문단 번호 (0부터 시작)
    text: 추가할 텍스트
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        # 문단으로 이동
        hwp.HAction.Run("MoveDocBegin")
        for i in range(paragraph_number):
            hwp.HAction.Run("MoveParaDown")

        # 문단 끝으로 이동
        hwp.HAction.Run("MoveParaEnd")

        # 텍스트 삽입
        act = hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        pset.SetItem("Text", text)
        act.Execute(pset)

        logger.info(f"{paragraph_number}번째 문단 끝에 텍스트 추가 완료")
        return f"{paragraph_number}번째 문단 끝에 '{text[:30]}...'를 추가했습니다."

    except Exception as e:
        logger.error(f"문단 끝 추가 실패: {e}")
        return f"문단 끝 추가 실패: {e}"


@mcp.tool()
def prepend_to_paragraph(paragraph_number: int, text: str) -> str:
    """
    특정 문단의 앞에 텍스트를 추가합니다.
    paragraph_number: 문단 번호 (0부터 시작)
    text: 추가할 텍스트
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        # 문단으로 이동
        hwp.HAction.Run("MoveDocBegin")
        for i in range(paragraph_number):
            hwp.HAction.Run("MoveParaDown")

        # 문단 시작으로 이동 (이미 시작 위치)
        # 텍스트 삽입
        act = hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        pset.SetItem("Text", text)
        act.Execute(pset)

        logger.info(f"{paragraph_number}번째 문단 앞에 텍스트 추가 완료")
        return f"{paragraph_number}번째 문단 앞에 '{text[:30]}...'를 추가했습니다."

    except Exception as e:
        logger.error(f"문단 앞 추가 실패: {e}")
        return f"문단 앞 추가 실패: {e}"


@mcp.tool()
def insert_at_page_start(page_number: int, text: str) -> str:
    """
    특정 페이지의 시작 부분에 텍스트를 삽입합니다.
    page_number: 페이지 번호 (1부터 시작)
    text: 삽입할 텍스트
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        total_pages = hwp.PageCount
        if page_number < 1 or page_number > total_pages:
            return f"페이지 번호가 범위를 벗어났습니다. (1~{total_pages})"

        # 페이지로 이동
        hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
        hwp.HParameterSet.HGotoE.PageNumber = page_number
        hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)

        # 페이지 시작으로 이동
        hwp.HAction.Run("MovePageBegin")

        # 텍스트 삽입
        act = hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        pset.SetItem("Text", text)
        act.Execute(pset)

        logger.info(f"{page_number}페이지 시작에 텍스트 삽입 완료")
        return f"{page_number}페이지 시작에 '{text[:30]}...'를 삽입했습니다."

    except Exception as e:
        logger.error(f"페이지 시작 삽입 실패: {e}")
        return f"페이지 시작 삽입 실패: {e}"


@mcp.tool()
def insert_at_page_end(page_number: int, text: str) -> str:
    """
    특정 페이지의 끝 부분에 텍스트를 삽입합니다.
    page_number: 페이지 번호 (1부터 시작)
    text: 삽입할 텍스트
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        total_pages = hwp.PageCount
        if page_number < 1 or page_number > total_pages:
            return f"페이지 번호가 범위를 벗어났습니다. (1~{total_pages})"

        # 페이지로 이동
        hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
        hwp.HParameterSet.HGotoE.PageNumber = page_number
        hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)

        # 페이지 끝으로 이동
        hwp.HAction.Run("MovePageEnd")

        # 텍스트 삽입
        act = hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        pset.SetItem("Text", text)
        act.Execute(pset)

        logger.info(f"{page_number}페이지 끝에 텍스트 삽입 완료")
        return f"{page_number}페이지 끝에 '{text[:30]}...'를 삽입했습니다."

    except Exception as e:
        logger.error(f"페이지 끝 삽입 실패: {e}")
        return f"페이지 끝 삽입 실패: {e}"


# ============================================================
# 선택 기능
# ============================================================

@mcp.tool()
def select_paragraph_by_number(paragraph_number: int) -> str:
    """
    특정 문단을 선택합니다.
    paragraph_number: 선택할 문단 번호 (0부터 시작)
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        # 문단으로 이동
        hwp.HAction.Run("MoveDocBegin")
        for i in range(paragraph_number):
            hwp.HAction.Run("MoveParaDown")

        # 문단 선택
        hwp.HAction.Run("MoveSelParaDown")

        # 선택된 내용 확인
        selected_text = _fix_mojibake(hwp.GetTextFile("TEXT", "") or "")
        if selected_text is None:
            selected_text = ""

        logger.info(f"{paragraph_number}번째 문단 선택 완료")
        return f"{paragraph_number}번째 문단을 선택했습니다. ({len(selected_text)}자)"

    except Exception as e:
        logger.error(f"문단 선택 실패: {e}")
        return f"문단 선택 실패: {e}"


@mcp.tool()
def select_page_content(page_number: int) -> str:
    """
    특정 페이지의 모든 내용을 선택합니다.
    page_number: 선택할 페이지 번호 (1부터 시작)
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        total_pages = hwp.PageCount
        if page_number < 1 or page_number > total_pages:
            return f"페이지 번호가 범위를 벗어났습니다. (1~{total_pages})"

        # 페이지로 이동
        hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
        hwp.HParameterSet.HGotoE.PageNumber = page_number
        hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)

        # 페이지 시작으로 이동
        hwp.HAction.Run("MovePageBegin")
        # 페이지 전체 선택
        hwp.HAction.Run("MoveSelPageDown")

        # 선택된 내용 확인
        selected_text = _fix_mojibake(hwp.GetTextFile("TEXT", "") or "")
        if selected_text is None:
            selected_text = ""

        logger.info(f"{page_number}페이지 선택 완료")
        return f"{page_number}페이지를 선택했습니다. ({len(selected_text)}자)"

    except Exception as e:
        logger.error(f"페이지 선택 실패: {e}")
        return f"페이지 선택 실패: {e}"


# ============================================================
# 성능 최적화 및 자동화 제어
# ============================================================

@mcp.tool()
def set_screen_updating(enabled: bool = True) -> str:
    """
    화면 업데이트를 켜거나 끕니다.
    대량 작업 시 False로 설정하면 성능이 크게 향상됩니다.
    작업 완료 후 반드시 True로 되돌려야 합니다.

    enabled: True=화면 업데이트 켜기, False=화면 업데이트 끄기
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        if enabled:
            hwp.SetScreenUpdate(1)  # 화면 업데이트 켜기
            logger.info("화면 업데이트 활성화")
            return "화면 업데이트를 활성화했습니다. (정상 속도)"
        else:
            hwp.SetScreenUpdate(0)  # 화면 업데이트 끄기
            logger.info("화면 업데이트 비활성화")
            return "화면 업데이트를 비활성화했습니다. (고속 모드)"

    except Exception as e:
        logger.error(f"화면 업데이트 설정 실패: {e}")
        return f"화면 업데이트 설정 실패: {e}"


@mcp.tool()
def set_automation_mode(enabled: bool = True) -> str:
    """
    자동화 모드를 켜거나 끕니다.
    enabled=True이면 모든 확인 대화상자가 자동으로 승인됩니다.

    enabled: True=자동화 모드 (대화상자 없음), False=일반 모드 (대화상자 표시)
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        if enabled:
            # 자동화 모드 활성화
            try:
                hwp.MessageBoxMode = 0  # 0=자동응답
            except:
                pass
            try:
                hwp.SetMessageBoxMode(0)
            except:
                pass
            logger.info("자동화 모드 활성화")
            return "자동화 모드를 활성화했습니다. (모든 확인창 자동 승인)"
        else:
            # 일반 모드로 전환
            try:
                hwp.MessageBoxMode = 1  # 1=대화상자표시
            except:
                pass
            try:
                hwp.SetMessageBoxMode(1)
            except:
                pass
            logger.info("일반 모드 활성화")
            return "일반 모드로 전환했습니다. (확인창 표시)"

    except Exception as e:
        logger.error(f"자동화 모드 설정 실패: {e}")
        return f"자동화 모드 설정 실패: {e}"


@mcp.tool()
def optimize_for_bulk_operations() -> str:
    """
    대량 작업을 위한 최적화 설정을 적용합니다.
    - 화면 업데이트 비활성화
    - 자동화 모드 활성화
    - 성능 최대화

    작업 완료 후 restore_normal_mode()를 호출하세요.
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        # 화면 업데이트 끄기
        try:
            hwp.SetScreenUpdate(0)
        except:
            pass

        # 자동화 모드 켜기
        try:
            hwp.MessageBoxMode = 0
        except:
            pass

        # 자동 저장 비활성화 (성능 향상)
        try:
            hwp.SetAutoSave(0)
        except:
            pass

        logger.info("대량 작업 최적화 모드 활성화")
        return "대량 작업 최적화 모드를 활성화했습니다. (최고 성능)\n작업 완료 후 restore_normal_mode()를 호출하세요."

    except Exception as e:
        logger.error(f"최적화 모드 설정 실패: {e}")
        return f"최적화 모드 설정 실패: {e}"


@mcp.tool()
def restore_normal_mode() -> str:
    """
    최적화 설정을 해제하고 일반 모드로 복원합니다.
    - 화면 업데이트 활성화
    - 자동 저장 활성화
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        # 화면 업데이트 켜기
        try:
            hwp.SetScreenUpdate(1)
        except:
            pass

        # 자동 저장 켜기
        try:
            hwp.SetAutoSave(1)
        except:
            pass

        # 화면 갱신
        try:
            hwp.Run("Repaginate")
        except:
            pass

        logger.info("일반 모드 복원 완료")
        return "일반 모드로 복원했습니다. (화면 업데이트 활성화)"

    except Exception as e:
        logger.error(f"일반 모드 복원 실패: {e}")
        return f"일반 모드 복원 실패: {e}"


@mcp.tool()
def replace_paragraph(paragraph_number: int, new_text: str) -> str:
    """
    특정 문단의 내용을 완전히 새 텍스트로 교체합니다.
    paragraph_number: 교체할 문단 번호 (0부터 시작)
    new_text: 새로운 텍스트
    """
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        # 문단으로 이동
        hwp.HAction.Run("MoveDocBegin")
        for i in range(paragraph_number):
            hwp.HAction.Run("MoveParaDown")

        # 문단 선택
        hwp.HAction.Run("MoveSelParaDown")

        # 기존 내용 확인
        old_text = _fix_mojibake(hwp.GetTextFile("TEXT", "") or "")
        if old_text is None:
            old_text = ""

        # 삭제
        hwp.HAction.Run("Delete")

        # 새 텍스트 삽입
        act = hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        pset.SetItem("Text", new_text)
        act.Execute(pset)

        logger.info(f"{paragraph_number}번째 문단 교체 완료")
        return f"{paragraph_number}번째 문단을 교체했습니다.\n이전: {old_text[:50] if old_text else '(빈 문단)'}...\n새 내용: {new_text[:50]}..."

    except Exception as e:
        logger.error(f"문단 교체 실패: {e}")
        return f"문단 교체 실패: {e}"


# ============================================================
# v2 신규 도구 — 실전 문서 자동화에서 검증된 패턴들
# ============================================================

@mcp.tool()
def find_text_and_move(search_text: str, occurrence: int = 1,
                       select: bool = False) -> str:
    """텍스트를 찾아 커서를 실제로 그 위치로 이동합니다. (기존 find_text는 위치 보고만 함)
    occurrence: 몇 번째 일치로 이동할지 (1부터). select=True면 찾은 텍스트를 선택 상태로 둠."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        found = _find_once(hwp, search_text, from_start=True)
        n = 1
        while found and n < occurrence:
            found = _find_once(hwp, search_text, from_start=False)
            n += 1
        if not found:
            hwp.HAction.Run("Cancel")
            return f"'{search_text}' {occurrence}번째 일치를 찾을 수 없습니다. ({n-1}개까지 확인)"
        if not select:
            hwp.HAction.Run("Cancel")
        pos = hwp.GetPos()
        logger.info(f"찾기 이동: '{search_text}' #{occurrence} -> {pos}")
        return (f"'{search_text}' {occurrence}번째 위치로 이동했습니다. "
                f"(list={pos[0]}, para={pos[1]}, pos={pos[2]}"
                f"{', 선택됨' if select else ''})")
    except Exception as e:
        logger.error(f"찾기 이동 실패: {e}")
        return f"찾기 이동 실패: {e}"


@mcp.tool()
def fill_cell_after_label(label: str, value: str, hops: int = 1,
                          clear_existing: bool = True) -> str:
    """표에서 라벨 셀을 찾아 오른쪽 셀에 값을 기입합니다. 신청서·양식 빈칸 채우기의 표준 방법.
    hops: 라벨에서 오른쪽으로 이동할 셀 수."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        if not _find_once(hwp, label):
            hwp.HAction.Run("Cancel")
            return f"라벨을 찾을 수 없습니다: {label}"
        hwp.HAction.Run("Cancel")
        for _ in range(hops):
            hwp.HAction.Run("TableRightCell")
        # 표 밖이면 TableRightCell이 no-op → 본문을 훼손하기 전에 중단 (필수 가드)
        if hwp.GetPos()[0] == 0:
            return (f"'{label}'이(가) 표 안에 없습니다 (첫 일치가 본문 텍스트). "
                    f"문서를 변경하지 않았습니다. 라벨을 더 구체적으로 지정하세요.")
        if clear_existing:
            hwp.HAction.Run("MoveParaBegin")
            hwp.HAction.Run("MoveSelParaEnd")
            hwp.HAction.Run("Delete")
        ps = hwp.HParameterSet.HInsertText
        hwp.HAction.GetDefault("InsertText", ps.HSet)
        ps.Text = value
        hwp.HAction.Execute("InsertText", ps.HSet)
        logger.info(f"셀 기입: {label} -> {value}")
        return f"'{label}' 오른쪽 {hops}번째 셀에 기입했습니다: {value}"
    except Exception as e:
        logger.error(f"셀 기입 실패: {e}")
        return f"셀 기입 실패: {e}"


@mcp.tool()
def clear_cell_after_label(label: str, hops: int = 1) -> str:
    """라벨 셀 기준 오른쪽 셀의 내용을 완전히 비웁니다 (여러 문단·자동번호 포함).
    안내문 삭제('※ 삭제 후 제출' 지시 이행) 등에 사용."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        if not _find_once(hwp, label):
            hwp.HAction.Run("Cancel")
            return f"라벨을 찾을 수 없습니다: {label}"
        hwp.HAction.Run("Cancel")
        for _ in range(hops):
            hwp.HAction.Run("TableRightCell")
        lst = hwp.GetPos()[0]
        if lst == 0:
            return "커서가 표 안이 아닙니다. 라벨이 표 밖에 있는 것 같습니다."
        cleared = _clear_list_content(hwp, lst)
        logger.info(f"셀 비움: {label} (list={lst}, 문단 {cleared}개)")
        return f"'{label}' 오른쪽 셀을 비웠습니다. (문단 {cleared}개, 자동번호 해제 포함)"
    except Exception as e:
        logger.error(f"셀 비움 실패: {e}")
        return f"셀 비움 실패: {e}"


@mcp.tool()
def delete_paragraph_containing(search_text: str) -> str:
    """지정 텍스트를 '포함한' 문단의 텍스트 전체를 삭제합니다.
    곡선따옴표 등 특수문자 때문에 정확 매칭 치환이 실패할 때의 확실한 대안.
    주의: 빈 문단 자체는 남습니다(문단 병합은 커서 상태를 꼬이게 할 수 있어 하지 않음)."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        if not _find_once(hwp, search_text):
            hwp.HAction.Run("Cancel")
            return f"'{search_text}'를 찾을 수 없습니다."
        hwp.HAction.Run("Cancel")
        hwp.HAction.Run("MoveParaBegin")
        hwp.HAction.Run("MoveSelParaEnd")
        hwp.HAction.Run("Delete")
        logger.info(f"문단 삭제: '{search_text}' 포함 문단")
        return f"'{search_text}'를 포함한 문단의 텍스트를 삭제했습니다."
    except Exception as e:
        logger.error(f"문단 삭제 실패: {e}")
        return f"문단 삭제 실패: {e}"


@mcp.tool()
def format_text_occurrence(search_text: str, color: Optional[str] = None,
                           bold: Optional[bool] = None,
                           italic: Optional[bool] = None,
                           font_size: Optional[float] = None,
                           occurrence: int = 1) -> str:
    """특정 텍스트를 찾아 그 부분만 서식을 바꿉니다. color는 RGB(예: '#000000', 'red').
    양식의 파란 이탤릭 안내 서식을 검정 정자체로 바꾸는 등에 사용."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        found = _find_once(hwp, search_text, from_start=True)
        n = 1
        while found and n < occurrence:
            found = _find_once(hwp, search_text, from_start=False)
            n += 1
        if not found:
            hwp.HAction.Run("Cancel")
            return f"'{search_text}'를 찾을 수 없습니다."
        # 찾은 텍스트가 선택된 상태 — 그대로 서식 적용
        _charshape_selection(hwp, size_pt=font_size, bold=bold,
                             italic=italic, color=color)
        hwp.HAction.Run("Cancel")
        changes = []
        if color is not None:
            changes.append(f"색={color}")
        if bold is not None:
            changes.append(f"굵게={bold}")
        if italic is not None:
            changes.append(f"기울임={italic}")
        if font_size is not None:
            changes.append(f"{font_size}pt")
        logger.info(f"서식 변경: '{search_text}' ({', '.join(changes)})")
        return f"'{search_text}'의 서식을 변경했습니다: {', '.join(changes) or '(변경 없음)'}"
    except Exception as e:
        logger.error(f"서식 변경 실패: {e}")
        return f"서식 변경 실패: {e}"


@mcp.tool()
def list_document_images() -> str:
    """문서 내 모든 그림 개체의 번호와 크기(mm)를 나열합니다. resize_document_image와 함께 사용."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        pics = _iter_pictures(hwp)
        if not pics:
            return "문서에 그림·도형 개체가 없습니다."
        lines = [f"그림·도형 개체 {len(pics)}개 (gso — 그림 외 글상자·도형 포함 가능):"]
        for i, c in enumerate(pics):
            try:
                p = c.Properties
                wmm = p.Item("Width") / MM_TO_HWPUNIT
                hmm = p.Item("Height") / MM_TO_HWPUNIT
                lines.append(f"  #{i}: {wmm:.1f} x {hmm:.1f} mm")
            except Exception:
                lines.append(f"  #{i}: (크기 조회 실패)")
        return "\n".join(lines)
    except Exception as e:
        logger.error(f"그림 나열 실패: {e}")
        return f"그림 나열 실패: {e}"


@mcp.tool()
def resize_document_image(index: int, width_mm: float,
                          height_mm: float = 0) -> str:
    """문서 내 index번째 그림(list_document_images 기준)의 크기를 조정합니다.
    height_mm=0이면 원본 비율 유지. 분량 조절 시 텍스트를 지우는 대신 이미지를 줄일 때 사용."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        pics = _iter_pictures(hwp)
        if not (0 <= index < len(pics)):
            return f"그림 번호 범위 초과: {index} (0~{len(pics)-1})"
        c = pics[index]
        p = c.Properties
        wm = p.Item("Width") / MM_TO_HWPUNIT
        hm = p.Item("Height") / MM_TO_HWPUNIT
        th = height_mm if height_mm else round(hm * width_mm / wm, 1)
        p.SetItem("Width", int(width_mm * MM_TO_HWPUNIT))
        p.SetItem("Height", int(th * MM_TO_HWPUNIT))
        c.Properties = p
        logger.info(f"그림 리사이즈 #{index}: {wm:.1f}x{hm:.1f} -> {width_mm}x{th}")
        return f"그림 #{index}: {wm:.1f}x{hm:.1f}mm -> {width_mm}x{th}mm"
    except Exception as e:
        logger.error(f"그림 리사이즈 실패: {e}")
        return f"그림 리사이즈 실패: {e}"


@mcp.tool()
def measure_layout(marker_texts: Optional[str] = None) -> str:
    """문서를 임시 PDF로 내보내 실제 레이아웃을 실측합니다. (PageCount보다 정확)
    - 총 페이지 수
    - marker_texts(JSON 배열, 예: '["1-1. 창업 동기","서식 1"]')가 처음 나타나는 페이지
    - 각 마커 페이지의 이미지 개수 (텍스트만 보면 이미지 넘침을 놓치므로)
    PyMuPDF(fitz)가 필요합니다: pip install pymupdf"""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        try:
            import fitz
        except ImportError:
            return ("PyMuPDF가 설치되어 있지 않습니다: pip install pymupdf\n"
                    f"(참고: 한글 PageCount = {hwp.PageCount})")

        import tempfile
        tmp_pdf = os.path.join(tempfile.gettempdir(),
                               f"hwp_measure_{os.getpid()}.pdf")
        try:
            hwp.SaveAs(tmp_pdf, "PDF", "")
        except Exception as e:
            return f"실측용 PDF 내보내기 실패: {e}"
        if not os.path.exists(tmp_pdf):
            return "실측용 PDF가 생성되지 않았습니다."

        markers = []
        if marker_texts:
            try:
                parsed = json.loads(marker_texts)
                if isinstance(parsed, list):
                    markers = [str(m) for m in parsed]
                else:
                    markers = [str(parsed)]     # 스칼라 JSON도 마커 1개로 취급
            except Exception:
                markers = [m.strip() for m in marker_texts.split(",") if m.strip()]

        doc = fitz.open(tmp_pdf)
        lines = [f"총 {len(doc)}페이지 (한글 PageCount={hwp.PageCount})"]
        pos = {}
        img_per_page = {}
        for i, page in enumerate(doc):
            img_per_page[i + 1] = len(page.get_images(full=True))
            if markers:
                t = page.get_text()
                for m in markers:
                    if m not in pos and m in t:
                        pos[m] = i + 1
        for m in markers:
            pg = pos.get(m)
            if pg:
                lines.append(f"  '{m}': p{pg} (해당 페이지 이미지 {img_per_page.get(pg, 0)}개)")
            else:
                lines.append(f"  '{m}': 못 찾음")
        if len(pos) >= 2:
            vals = [pos[m] for m in markers if m in pos]
            lines.append(f"  구간 길이(마지막-첫): {max(vals) - min(vals)}p")
        doc.close()
        try:
            os.remove(tmp_pdf)
        except Exception:
            pass
        return "\n".join(lines)
    except Exception as e:
        logger.error(f"레이아웃 실측 실패: {e}")
        return f"레이아웃 실측 실패: {e}"


@mcp.tool()
def render_content_blocks(blocks_json: str, base_font_size: float = 10.5) -> str:
    """구조화된 블록(JSON 배열)을 커서 위치에 조판 규칙대로 렌더링합니다.
    저수준 커서 조작 없이 문서 본문을 만들 때 사용. 블록 타입:
      {"type":"h","text":...}        제목 (ㅇ 접두, 굵게)
      {"type":"b","text":...}        불릿 (- 접두)
      {"type":"b2","text":...}       하위 불릿 (· 접두)
      {"type":"p","text":...}        문단
      {"type":"callout","text":...}  강조 (▶ 접두, 진갈색)
      {"type":"note","text":...}     주석 (※ 접두, 회색)
      {"type":"t","headers":[...],"rows":[[...]],"widths":[...],"title":...}  표
      {"type":"img","path":...,"width_mm":...,"caption":...}  이미지
    표는 폭 유지·탈출 처리가 내장되어 있습니다."""
    try:
        hwp_controller.check_initialization()
        hwp = hwp_controller.hwp

        try:
            blocks = json.loads(blocks_json)
        except Exception as e:
            return f"blocks_json 파싱 실패: {e}"
        if not isinstance(blocks, list):
            return "blocks_json은 JSON 배열이어야 합니다."

        def para_shape(align=0, left=0, indent=0, sb=0.0, sa=0.0, ls=150):
            ps = hwp.HParameterSet.HParaShape
            hwp.HAction.GetDefault("ParagraphShape", ps.HSet)
            ps.AlignType = align
            ps.LeftMargin = _mm(hwp, left)
            ps.Indentation = _mm(hwp, indent)
            ps.PrevSpacing = int(sb * 100)
            ps.NextSpacing = int(sa * 100)
            ps.LineSpacing = ls
            hwp.HAction.Execute("ParagraphShape", ps.HSet)

        def put(text):
            ps = hwp.HParameterSet.HInsertText
            hwp.HAction.GetDefault("InsertText", ps.HSet)
            ps.Text = text
            hwp.HAction.Execute("InsertText", ps.HSet)

        def para(text="", size=None, bold=False, color=None, align=0,
                 left=0, indent=0, sb=0.0, sa=0.0, ls=150):
            para_shape(align, left, indent, sb, sa, ls)
            _charshape_selection(hwp, size_pt=size or base_font_size,
                                 bold=bold, color=color or 0x000000)
            if text:
                put(text)
            hwp.HAction.Run("BreakPara")

        bs = base_font_size
        rendered = 0
        skipped = 0
        for b in blocks:
            if not isinstance(b, dict):
                skipped += 1
                continue
            t = b.get("type")
            if t == "h":
                para("ㅇ " + b.get("text", ""), size=bs + 0.5, bold=True,
                     sb=5, sa=1.5, ls=160)
            elif t == "b":
                para("- " + b.get("text", ""), left=3.5, indent=-3.5,
                     sa=0.5, ls=150)
            elif t == "b2":
                para("· " + b.get("text", ""), size=bs - 0.5, left=7.0,
                     indent=-3.0, sa=0.4, ls=145, color=0x333333)
            elif t == "p":
                para(b.get("text", ""), left=3.5, sa=1.0, ls=155)
            elif t == "callout":
                para("▶ " + b.get("text", ""), bold=True, left=3.5,
                     indent=-3.5, sb=2.0, sa=2.0, ls=150, color=0x8B2B00)
            elif t == "note":
                para("※ " + b.get("text", ""), size=bs - 0.5, left=3.5,
                     sa=1.0, ls=140, color=0x666666)
            elif t == "t":
                if b.get("title"):
                    para("· " + b["title"], size=bs - 0.5, bold=True,
                         left=3.5, sb=2.0, sa=1.0, ls=145)
                hd = b.get("headers") or []
                rw = b.get("rows") or []
                data = ([hd] + rw) if hd else rw
                if data:
                    nr, nc = len(data), max(len(r) for r in data)
                    res = create_table(nr, nc, width_mm=b.get("width_mm", 150),
                                       col_widths=json.dumps(b.get("widths"))
                                       if b.get("widths") else None,
                                       data_json=json.dumps(data, ensure_ascii=False),
                                       header_bold=bool(hd),
                                       font_size=bs - 0.5)
                    if "실패" in res:
                        return f"표 렌더 중단: {res}"
                    para("", size=5, ls=100)
            elif t == "img":
                path = b.get("path", "")
                if path and os.path.exists(path):
                    res = insert_image(path, width=b.get("width_mm", 100),
                                       height=b.get("height_mm", 0))
                    hwp.HAction.Run("BreakPara")
                    if b.get("caption"):
                        para("※ " + b["caption"], size=bs - 0.5, align=3,
                             color=0x888888, sa=1.5, ls=130)
                else:
                    para(f"[이미지 없음: {path}]", color="red")
            rendered += 1

        logger.info(f"블록 렌더 완료: {rendered}개 (건너뜀 {skipped}개)")
        note = f" (dict가 아닌 항목 {skipped}개 건너뜀)" if skipped else ""
        return f"{rendered}개 블록을 렌더링했습니다.{note}"
    except Exception as e:
        logger.error(f"블록 렌더 실패: {e}")
        return f"블록 렌더 실패: {e}"


@mcp.tool()
def kill_hwp_processes() -> str:
    """모든 한글(Hwp) 프로세스를 강제 종료합니다. 파일 잠금(PermissionError)으로
    문서를 열 수 없을 때 사용하세요. 저장하지 않은 변경사항은 사라지니 주의."""
    try:
        import subprocess
        killed = []
        for name in ("Hwp.exe", "HwpApi.exe", "Hwp"):
            try:
                r = subprocess.run(["taskkill", "/F", "/IM",
                                    name if name.endswith(".exe") else name + ".exe"],
                                   capture_output=True, text=True, timeout=15)
                if r.returncode == 0:
                    killed.append(name)
            except Exception:
                pass
        # 컨트롤러 상태 리셋
        hwp_controller.hwp = None
        hwp_controller.is_initialized = False
        hwp_controller.current_document = None
        if killed:
            return f"한글 프로세스를 종료했습니다: {', '.join(killed)}. 다음 도구 호출 시 자동으로 재초기화됩니다."
        return "종료할 한글 프로세스가 없습니다."
    except Exception as e:
        return f"프로세스 종료 실패: {e}"


def main():
    """메인 함수"""
    try:
        logger.info("Advanced HWP MCP Server 시작")
        
        mcp.run()
        
    except Exception as e:
        logger.error(f"서버 실행 중 오류: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
