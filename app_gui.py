#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import time
from datetime import datetime
import queue
import shutil
import tempfile
import re
import subprocess
import psutil
import logging
import random
from openpyxl.utils import get_column_letter

# 기존 분석기 및 비디오 프로세서 임포트
from video_processor import VideoProcessor
from image_analyzer import GoogleVisionAnalyzer
from main import is_valid_format, check_api_keys, validate_frame_times, is_valid_video_file, clear_directory

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# 파일 유효성 검사 함수
def is_valid_file(file_path, valid_extensions):
    """파일 유효성 검사 (확장자 기준)"""
    if not file_path or not os.path.isfile(file_path):
        return False
    _, ext = os.path.splitext(file_path)
    return ext.lower() in valid_extensions

def is_valid_image_file(file_path):
    """이미지 파일 유효성 검사"""
    valid_exts = ['.jpg', '.jpeg', '.png', '.gif', '.bmp']
    return is_valid_file(file_path, valid_exts)
    
def is_valid_video_file(file_path):
    """비디오 파일 유효성 검사"""
    valid_exts = ['.mp4', '.avi', '.mov', '.wmv', '.mkv', '.flv']
    return is_valid_file(file_path, valid_exts)

class RedirectText:
    """콘솔 출력을 GUI로 리다이렉트하는 클래스"""
    def __init__(self, text_widget, max_messages=1000):
        self.text_widget = text_widget
        self.queue = queue.Queue()
        self.update_timer = None
        self.last_message = ""  # 마지막으로 출력된 메시지 저장
        self.max_messages = max_messages  # 최대 메시지 수 (메모리 관리)
        self.message_count = 0
        
        # 텍스트 태그 설정
        self.text_widget.tag_configure("success", foreground="green")
        self.text_widget.tag_configure("error", foreground="red")
        self.text_widget.tag_configure("warning", foreground="orange")
        self.text_widget.tag_configure("info", foreground="blue")
        self.text_widget.tag_configure("header", foreground="black", font=("Malgun Gothic", 10, "bold"))
        
        # 키워드 강조 태그 설정
        self.text_widget.tag_configure("success_keyword", foreground="green", font=("Malgun Gothic", 10, "bold"))
        self.text_widget.tag_configure("error_keyword", foreground="red", font=("Malgun Gothic", 10, "bold"))
        self.text_widget.tag_configure("warning_keyword", foreground="orange", font=("Malgun Gothic", 10, "bold"))

    def write(self, string):
        # 빈 문자열이면 무시
        if not string:
            return
            
        # 로그 필터링 - 불필요한 디버그 메시지 필터링
        if self._should_filter_message(string):
            return
            
        # 중복 메시지 필터링 (동일한 메시지가 연속으로 출력되는 것 방지)
        if string.strip() and string.strip() == self.last_message:
            return
            
        self.last_message = string.strip()
        self.queue.put(string)
        
        # 처음 호출되는 경우에만 타이머 시작
        if self.update_timer is None:
            # 로그 업데이트 간격을 100ms에서 20ms로 줄여 더 빠른 응답성 제공
            self.update_timer = self.text_widget.after(20, self.update_text)
        
        # 중요 로그 패턴 감지 시 즉시 업데이트 시도
        if string.startswith("✓") or string.startswith("❌") or string.startswith("⚠️") or "[" in string:
            # 이미 타이머가 설정된 경우 취소하고 즉시 업데이트
            if self.update_timer:
                self.text_widget.after_cancel(self.update_timer)
                self.update_timer = self.text_widget.after(1, self.update_text)
    
    def _should_filter_message(self, message):
        """필터링할 메시지인지 확인"""
        # 빈 줄 여러개 필터링
        if message.strip() == "" and self.last_message.strip() == "":
            return True
            
        # 디버깅 목적의 상세 로그 필터링
        filter_patterns = [
            "행 분석:", "열 위치 계산:", "셀 RGB 값:", "현재 셀 배경색", 
            "디버깅", "배경색 확인", "추출 성공:", "cell_width", "cell_height",
            "    ->", "    행", "    열", "객체 정리", "메모리 정리"
        ]
        
        for pattern in filter_patterns:
            if pattern in message:
                return True
                
        return False
            
    def _get_tag_for_message(self, message):
        """메시지 유형에 따른 태그 결정"""
        message = message.strip()
        
        if message.startswith("✓"):
            return "success"
        elif message.startswith("❌"):
            return "error"
        elif message.startswith("⚠️"):
            return "warning"
        elif message.startswith("[") and "]" in message:
            return "header"
        elif message.startswith("•") or message.startswith("==="):
            return "info"
        else:
            return None

    def update_text(self):
        """큐에 있는 메시지를 텍스트 위젯에 업데이트"""
        self.update_timer = None
        
        try:
            # 한 번의 업데이트에서 최대 처리할 메시지 수 제한
            messages_processed = 0
            max_batch_size = 10  # 한 번에 최대 10개 메시지 처리
            
            while not self.queue.empty() and messages_processed < max_batch_size:
                string = self.queue.get_nowait()
                self.text_widget.configure(state='normal')
                
                # 메시지 유형에 따른 태그 결정
                tag = self._get_tag_for_message(string)
                
                # 텍스트 위젯에 메시지 추가 (태그 적용)
                if tag:
                    # 메시지 전체 삽입
                    self.text_widget.insert(tk.END, string, tag)
                    
                    # 성공, 실패, 오류 등의 키워드에 대한 강조 처리
                    self._highlight_keywords(string)
                else:
                    # 일반 메시지 삽입
                    self.text_widget.insert(tk.END, string)
                    
                    # 키워드 강조 처리
                    self._highlight_keywords(string)
                
                self.message_count += 1
                messages_processed += 1
                
                # 최대 메시지 수를 초과하면 오래된 메시지 제거 (메모리 관리)
                if self.message_count > self.max_messages:
                    self.text_widget.delete(1.0, 2.0)
                    self.message_count -= 1
                
                # 스크롤을 최신 메시지로 이동
                self.text_widget.see(tk.END)
                self.text_widget.configure(state='disabled')
                self.queue.task_done()
                
            # 메시지 처리 후 강제 업데이트
            self.text_widget.update_idletasks()
            
            # 큐에 메시지가 더 있거나 배치 처리가 완료된 경우 다시 타이머 설정
            if not self.queue.empty() or messages_processed >= max_batch_size:
                # 즉시 다음 배치 처리 예약
                self.update_timer = self.text_widget.after(10, self.update_text)
            else:
                # 큐가 비어있으면 약간 더 길게 대기 후 다시 확인
                self.update_timer = self.text_widget.after(50, self.update_text)
                
        except queue.Empty:
            # 큐가 비어있으면 일정 시간 후 다시 확인
            self.update_timer = self.text_widget.after(50, self.update_text)
        except Exception as e:
            # 예외 발생 시 복구 시도
            print(f"로그 업데이트 중 오류: {str(e)}")
            self.update_timer = self.text_widget.after(100, self.update_text)
    
    def _highlight_keywords(self, message):
        """메시지 내의 성공, 실패, 오류 등의 키워드 강조 처리"""
        # 현재 위치 (마지막에 삽입된 텍스트)
        current_line = self.text_widget.index(tk.END + "-1c linestart")
        
        # 성공 관련 키워드 강조
        success_keywords = ["성공", "[성공]", "완료", "처리 완료"]
        for keyword in success_keywords:
            start_idx = 0
            while True:
                start_pos = self.text_widget.search(keyword, current_line + f"+{start_idx}c", tk.END)
                if not start_pos:
                    break
                end_pos = f"{start_pos}+{len(keyword)}c"
                self.text_widget.tag_add("success_keyword", start_pos, end_pos)
                # 다음 검색을 위해 인덱스 업데이트
                start_idx_parts = start_pos.split('.')
                if len(start_idx_parts) > 1:
                    start_idx = int(start_idx_parts[1]) + len(keyword)
        
        # 오류 관련 키워드 강조
        error_keywords = ["실패", "[실패]", "오류", "에러", "Error", "error"]
        for keyword in error_keywords:
            start_idx = 0
            while True:
                start_pos = self.text_widget.search(keyword, current_line + f"+{start_idx}c", tk.END)
                if not start_pos:
                    break
                end_pos = f"{start_pos}+{len(keyword)}c"
                self.text_widget.tag_add("error_keyword", start_pos, end_pos)
                # 다음 검색을 위해 인덱스 업데이트
                start_idx_parts = start_pos.split('.')
                if len(start_idx_parts) > 1:
                    start_idx = int(start_idx_parts[1]) + len(keyword)
        
        # 경고 관련 키워드 강조
        warning_keywords = ["건너뜀", "[건너뜀]", "경고", "주의"]
        for keyword in warning_keywords:
            start_idx = 0
            while True:
                start_pos = self.text_widget.search(keyword, current_line + f"+{start_idx}c", tk.END)
                if not start_pos:
                    break
                end_pos = f"{start_pos}+{len(keyword)}c"
                self.text_widget.tag_add("warning_keyword", start_pos, end_pos)
                # 다음 검색을 위해 인덱스 업데이트
                start_idx_parts = start_pos.split('.')
                if len(start_idx_parts) > 1:
                    start_idx = int(start_idx_parts[1]) + len(keyword)
    
    def flush(self):
        """파이썬 출력 스트림 호환을 위한 메서드"""
        # 큐에 있는 모든 메시지 즉시 처리 시도
        if self.update_timer:
            self.text_widget.after_cancel(self.update_timer)
            self.update_timer = self.text_widget.after(1, self.update_text)
        
    def clear(self):
        """텍스트 위젯의 내용을 지움"""
        self.text_widget.configure(state='normal')
        self.text_widget.delete(1.0, tk.END)
        self.text_widget.configure(state='disabled')
        self.message_count = 0

class FileRenamerXApp:
    """FileRenamerX 애플리케이션 메인 클래스"""
    def __init__(self, root):
        self.root = root
        self.root.title("FileRenamerX 1.0.0")
        self.root.geometry("700x900")  # 창 크기 설정
        self.root.minsize(700, 600)    # 최소 창 크기 제한
        
        # 상태 변수 초기화
        self.process_thread = None     # 작업 스레드
        self.running = False           # 실행 상태
        self.is_running = False        # 중복 방지용 상태 플래그
        self.temp_dir = None           # 임시 디렉토리 경로
        self.selected_files = []       # 선택된 파일 목록

        # UI 변수 초기화
        self.folder_path = tk.StringVar(value="./작업폴더")
        
        # 작업 설정 초기화
        self.frame_times_value = [2, 3, 5]  # 기본 프레임 시간 (초)

        # UI 구성
        self.create_widgets()
        
        # 기본 작업 폴더 생성
        self.ensure_work_folder()
        
        # 표준 출력 리다이렉션 (이전 참조 저장)
        self.old_stdout = sys.stdout
        sys.stdout = self.redirect

    def ensure_work_folder(self):
        """기본 작업 폴더가 존재하는지 확인하고 없으면 생성"""
        work_folder = "./작업폴더"
        if not os.path.exists(work_folder):
            try:
                os.makedirs(work_folder)
                logger.info(f"기본 작업 폴더 생성: {work_folder}")
            except Exception as e:
                logger.error(f"작업 폴더 생성 실패: {e}")

    def create_widgets(self):
        """UI 위젯 생성 및 배치"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 설정 프레임
        self.create_settings_frame(main_frame)
        
        # 3. 도움말 프레임
        self.create_help_frame(main_frame)
        
        # 4. 버튼 프레임
        self.create_button_frame(main_frame)
        
        # 5. 로그 프레임
        self.create_log_frame(main_frame)
        
        # 6. 상태바
        self.status_bar = ttk.Label(self.root, text="준비", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def create_settings_frame(self, parent):
        """설정 프레임 생성"""
        self.settings_frame = ttk.LabelFrame(parent, text="설정", padding=10)
        self.settings_frame.pack(fill=tk.X, pady=5)

        # 파일 경로 설정 프레임
        self.path_frame = ttk.Frame(self.settings_frame)
        self.path_frame.pack(fill=tk.X, pady=5)
        
        # 작업 폴더 선택 UI
        self.folder_button_frame = ttk.Frame(self.path_frame)
        self.folder_button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(
            self.folder_button_frame, 
            text="작업 폴더 선택", 
            command=self.browse_folder
        ).pack(side=tk.LEFT, padx=5)
        
        self.path_label = ttk.Label(self.folder_button_frame, text="./작업폴더")
        self.path_label.pack(side=tk.LEFT, padx=5)

    def create_help_frame(self, parent):
        """도움말 프레임 생성"""
        self.help_frame = ttk.LabelFrame(parent, text="프로그램 사용 안내", padding=10)
        self.help_frame.pack(fill=tk.X, pady=5)
        
        # 도움말 텍스트 위젯
        self.help_text = tk.Text(
            self.help_frame, 
            wrap=tk.WORD, 
            height=5, 
            font=("Malgun Gothic", 9)
        )
        self.help_text.pack(fill=tk.X)
        
        # 굵은 글씨 스타일 설정
        self.help_text.tag_configure("bold", font=("Malgun Gothic", 9, "bold"))
        
        # 초기 도움말 텍스트 설정
        self.help_text.insert(tk.END, "파일명 변경 모드\n", "bold")
        self.help_text.insert(tk.END, "작업 폴더 : 비디오 파일과 이미지 파일이 있는 폴더\n")
        self.help_text.insert(tk.END, "프레임 시간 : 비디오에서 추출할 프레임 시간(초)\n")
        self.help_text.insert(tk.END, "바로에코 제품으로 인코딩된 동영상만 처리 가능합니다.\n")
        self.help_text.configure(state='disabled')

    def create_button_frame(self, parent):
        """버튼 프레임 생성"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=10)

        # 시작 버튼
        self.start_button = ttk.Button(
            button_frame, 
            text="시작", 
            command=self.start_process
        )
        self.start_button.pack(side=tk.LEFT, padx=5)

        # 중지 버튼 (초기에는 비활성화)
        self.stop_button = ttk.Button(
            button_frame, 
            text="중지", 
            command=self.stop_process, 
            state=tk.DISABLED
        )
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # 종료 버튼
        self.exit_button = ttk.Button(
            button_frame, 
            text="종료", 
            command=self.on_exit
        )
        self.exit_button.pack(side=tk.RIGHT, padx=5)

    def create_log_frame(self, parent):
        """로그 프레임 생성"""
        log_frame = ttk.LabelFrame(parent, text="로그", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # 스크롤 가능한 텍스트 위젯
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            wrap=tk.WORD, 
            state='disabled'
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 출력 리다이렉션 설정
        self.redirect = RedirectText(self.log_text)

    def on_exit(self):
        """프로그램 종료 처리"""
        # 작업 중이면 중지 확인
        if self.is_running:
            if not messagebox.askyesno("확인", "작업이 진행 중입니다. 정말 종료하시겠습니까?"):
                return
            self.stop_process()
        
        # 표준 출력 복원
        if hasattr(self, 'old_stdout') and self.old_stdout:
            sys.stdout = self.old_stdout
            
        # 임시 폴더 정리
        self.cleanup_temp_dir()
        
        # 프로그램 종료
        self.root.destroy()


    def browse_folder(self):
        """작업 폴더 선택"""
        folder = filedialog.askdirectory(title="작업 폴더를 선택하세요")
        if not folder:
            return
            
        self.folder_path.set(folder)
        self.selected_files = []  # 폴더를 선택하면 개별 파일 선택 초기화
        self.path_label.config(text=f"{folder}")
        
        # 파일명 변경 모드에서는 비디오/이미지 파일 표시
        self.display_target_files(folder)

    def display_target_files(self, folder):
        """폴더 내 비디오/이미지 파일 목록 표시 (파일명 변경 모드)"""
        if not os.path.exists(folder):
            print(f"❌ 폴더를 찾을 수 없습니다: {folder}")
            return
            
        # 로그 초기화
        if hasattr(self, 'redirect'):
            self.redirect.clear()
        
        video_files = []
        image_files = []
        
        # 폴더 내 파일 검색
        try:
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                if os.path.isfile(file_path):
                    if is_valid_video_file(file_path):
                        video_files.append(file_path)
                    elif is_valid_image_file(file_path):
                        image_files.append(file_path)
        except PermissionError:
            print(f"❌ 폴더 접근 권한이 없습니다: {folder}")
            return
        except Exception as e:
            print(f"❌ 폴더 읽기 오류: {str(e)}")
            return
        
        total_files = len(video_files) + len(image_files)
        if total_files > 0:
            print(f"대상 파일 목록 (총 {total_files}개):")
            print(f"- 비디오 파일: {len(video_files)}개")
            print(f"- 이미지 파일: {len(image_files)}개")
            print("")
            
            # 최근 수정된 순으로 정렬하여 표시
            all_files = sorted(video_files + image_files, key=lambda x: os.path.getmtime(x), reverse=True)
            for i, file_path in enumerate(all_files, 1):
                filename = os.path.basename(file_path)
                file_type = "비디오" if is_valid_video_file(file_path) else "이미지"
                print(f"{i}. [{file_type}] {filename}")
        else:
            print("처리할 파일이 없습니다.")
    
    def display_image_files(self, folder):
        """폴더 내 이미지 파일 목록 표시 (이상 배관 업데이트 모드)"""
        if not os.path.exists(folder):
            print(f"❌ 폴더를 찾을 수 없습니다: {folder}")
            return
            
        # 로그 초기화
        if hasattr(self, 'redirect'):
            self.redirect.clear()
        
        image_files = []
        
        # 폴더 내 이미지 파일 검색
        try:
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                if os.path.isfile(file_path) and is_valid_image_file(file_path):
                    image_files.append(file_path)
        except PermissionError:
            print(f"❌ 폴더 접근 권한이 없습니다: {folder}")
            return
        except Exception as e:
            print(f"❌ 폴더 읽기 오류: {str(e)}")
            return
        
        if image_files:
            # 이미지 파일 개수 출력
            print(f"이미지 파일 목록 (총 {len(image_files)}개):")
            
            # 최근 수정된 순으로 정렬하여 표시
            sorted_files = sorted(image_files, key=lambda x: os.path.getmtime(x), reverse=True)
            for i, file_path in enumerate(sorted_files, 1):
                filename = os.path.basename(file_path)
                print(f"{i}. {filename}")
        else:
            print("처리할 이미지 파일이 없습니다.")

    def display_video_files(self, folder):
        """폴더 내 동영상 파일 목록 표시 (작업 현황 업데이트 모드)"""
        if not os.path.exists(folder):
            print(f"❌ 폴더를 찾을 수 없습니다: {folder}")
            return
            
        # 로그 초기화
        if hasattr(self, 'redirect'):
            self.redirect.clear()
        
        video_files = []
        
        # 폴더 내 동영상 파일 검색
        try:
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                if os.path.isfile(file_path) and is_valid_video_file(file_path):
                    video_files.append(file_path)
        except PermissionError:
            print(f"❌ 폴더 접근 권한이 없습니다: {folder}")
            return
        except Exception as e:
            print(f"❌ 폴더 읽기 오류: {str(e)}")
            return
        
        if video_files:
            # 동영상 파일 개수 출력
            print(f"동영상 파일 목록 (총 {len(video_files)}개):")
            
            # 최근 수정된 순으로 정렬하여 표시
            sorted_files = sorted(video_files, key=lambda x: os.path.getmtime(x), reverse=True)
            for i, file_path in enumerate(sorted_files, 1):
                filename = os.path.basename(file_path)
                print(f"{i}. {filename}")
        else:
            print("처리할 동영상 파일이 없습니다.")

    def start_process(self):
        """작업 시작"""
        try:
            # 이미 실행 중인지 확인
            if self.is_running:
                return
                
            # 작업 폴더 확인
            work_dir = self.folder_path.get()
            if not work_dir or not os.path.exists(work_dir):
                messagebox.showerror("오류", "작업 폴더를 선택해주세요.")
                return
                
            
            # UI 업데이트
            self.update_ui_for_processing(True)
            
            # 출력 폴더 생성 - 삭제
            output_dir = None  # 기본값으로 None 설정
            
            # 프레임 시간 설정
            frame_times = [2, 3, 5]  # 기본 프레임 시간
            
            # 작업 시작 시간 기록
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"\n[{current_time}] 작업 시작")
            
            # 파일명 변경 모드
            if True:  # 항상 파일명 변경 모드
                # 작업 폴더 내 파일 확인
                if not self.selected_files:
                    # 폴더 내 모든 비디오/이미지 파일 확인
                    all_files = []
                    for f in os.listdir(work_dir):
                        file_path = os.path.join(work_dir, f)
                        if os.path.isfile(file_path) and (is_valid_video_file(file_path) or is_valid_image_file(file_path)):
                            all_files.append(file_path)
                            
                    if not all_files:
                        messagebox.showerror("오류", "작업 폴더에 처리할 비디오/이미지 파일이 없습니다.")
                        return
                
                # API 키 유효성 확인
                missing_keys = check_api_keys()
                if missing_keys:  # 비어있지 않은 경우에만 오류
                    messagebox.showerror("오류", "API 키가 설정되지 않았거나 유효하지 않습니다.")
                    return
                    
                # 임시 디렉토리 생성
                try:
                    self.cleanup_temp_dir()  # 기존 임시 폴더 정리
                    self.temp_dir = tempfile.mkdtemp(prefix="filerenamerx_temp_")
                    output_dir = self.temp_dir
                    logger.info(f"임시 폴더 생성: {self.temp_dir}")
                except Exception as e:
                    messagebox.showerror("오류", f"임시 폴더 생성 실패: {str(e)}")
                    return
                    
                # API 요청 간 제한 시간 설정 (처리량 조절)
                min_api_delay = 1  # API 요청 사이의 최소 시간 (초)
                print(f"⚠️ 처리량 조절: API 요청 사이에 최소 {min_api_delay}초 간격을 유지합니다.")
            
            # 작업 실행 플래그 설정
            self.running = True
            self.is_running = True
            
            # 멀티스레딩으로 작업 실행
            thread = threading.Thread(
                target=self.process_in_thread,
                args=("rename", work_dir, output_dir, frame_times)  # 항상 rename 모드
            )
            thread.daemon = True
            thread.start()
            
        except Exception as e:
            messagebox.showerror("오류", f"작업 시작 중 오류가 발생했습니다: {str(e)}")
            self.is_running = False
            self.running = False
            self.update_ui_for_processing(False)

    def update_ui_for_processing(self, is_processing):
        """처리 중 UI 상태 업데이트"""
        if is_processing:
            # 처리 시작 시 UI 상태
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.exit_button.config(state=tk.DISABLED)
            self.status_bar.config(text="처리 중...")
            self.is_running = True
            self.running = True
        else:
            # 처리 종료 시 UI 상태
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.exit_button.config(state=tk.NORMAL)
            self.status_bar.config(text="완료")
            self.is_running = False
            self.running = False

    def stop_process(self):
        """작업 중지"""
        if not self.is_running:
            return
            
        # 실행 상태 변경
        self.running = False
        print("\n⚠️ 사용자에 의해 작업이 중지되었습니다.")
        self.status_bar.config(text="중지됨")
        
        # UI 업데이트
        self.update_ui_for_processing(False)

    def cleanup_temp_dir(self):
        """임시 디렉토리 정리"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                shutil.rmtree(self.temp_dir)
                self.temp_dir = None
                logger.info("임시 폴더 삭제 완료")
            except Exception as e:
                logger.error(f"임시 폴더 삭제 중 오류: {e}")

    def process_videos(self, work_dir, output_dir, frame_times):
        """비디오/이미지 처리 작업 수행"""
        try:
            # 비디오 프로세서 및 이미지 분석기 초기화
            print("✓ 작업 초기화 중...")
            video_processor = VideoProcessor(work_dir, output_dir)
            
            # 동영상 파일 메타데이터 검증
            print("✓ 동영상 파일 메타데이터 검증 중...")
            video_files = [f for f in os.listdir(work_dir) 
                          if os.path.isfile(os.path.join(work_dir, f)) and 
                          is_valid_video_file(os.path.join(work_dir, f))]
            
            if video_files:
                invalid_files = []
                for video_file in video_files:
                    video_path = os.path.join(work_dir, video_file)
                    is_valid, error_msg = video_processor.validate_video_metadata(video_path)
                    if not is_valid:
                        invalid_files.append((video_file, error_msg))
                        print(f"  ❌ {video_file}: 바로에코 제품을 사용해 주세요")
                
                if invalid_files:
                    # 검증 실패 시 팝업 표시 및 작업 중단
                    error_message = "바로에코 제품을 사용해 주세요"
                    
                    # 메인 스레드에서 팝업 표시를 위해 root.after 사용
                    self.root.after(0, lambda: messagebox.showerror("동영상 파일 검증 실패", error_message))
                    self.finish_process()
                    return
                else:
                    print(f"  ✓ 모든 동영상 파일 검증 완료 ({len(video_files)}개)")
            else:
                print("  ⚠️ 동영상 파일이 없습니다. 이미지 파일만 처리합니다.")
            
            # 분석기 초기화
            try:
                analyzer = GoogleVisionAnalyzer()
                print("✓ Vision API + ChatGPT 분석기가 준비되었습니다.")
            except Exception as e:
                print(f"❌ 분석기 초기화 중 오류: {str(e)}")
                self.finish_process()
                return
            
            # 작업 결과 저장 (변경 전/후 파일명 기록)
            file_changes = []
            
            # 파일 목록 가져오기 (비디오와, 이미지 포함)
            print("✓ 작업 파일 검색 중...")
            all_files = self.get_target_files(work_dir)
            
            if not all_files:
                print("❌ 처리할 파일이 없습니다.")
                self.finish_process()
                return
            
            # 비디오/이미지 파일 개수 카운트
            video_count = sum(1 for _, path in all_files if is_valid_video_file(path))
            image_count = sum(1 for _, path in all_files if is_valid_image_file(path))
            
            print(f"✓ 총 {len(all_files)}개 파일을 처리합니다:")
            print(f"  • 비디오: {video_count}개")
            print(f"  • 이미지: {image_count}개")
            print("✓ 파일명 변경 작업을 시작합니다...")
            
            # 진행 상황 표시용 변수
            total_processed = 0
            video_processed = 0
            image_processed = 0
            
            # 마지막으로 변경된 비디오 파일명 저장
            last_video_name = None
            last_video_base_name = None
            image_counter = 0
            
            # 각 파일 처리
            for idx, (file_rel_path, file_full_path) in enumerate(all_files, 1):
                if not self.running:
                    break
                
                # 진행률 계산 및 표시
                progress_pct = int(idx / len(all_files) * 100)
                self.status_bar.config(text=f"처리 중... {progress_pct}% ({idx}/{len(all_files)})")
                
                # 원본 파일명 저장
                original_filename = os.path.basename(file_full_path)
                is_video = is_valid_video_file(file_full_path)
                is_image = is_valid_image_file(file_full_path)
                
                # 파일 유형에 따라 처리
                if is_video:
                    # 비디오 파일 처리
                    result = self.process_video_file(
                        idx, len(all_files), file_rel_path, file_full_path, original_filename,
                        video_processor, analyzer, frame_times, 
                        file_changes, total_processed, video_processed,
                        last_video_name, last_video_base_name, image_counter
                    )
                    
                    if result:
                        # 결과에서 필요한 정보만 가져옴 (file_changes는 이미 process_video_file에서 추가됨)
                        total_processed += 1
                        video_processed += 1
                        image_counter = result.get('image_counter')
                        last_video_name = result.get('last_video_name')
                        last_video_base_name = result.get('last_video_base_name')
                        
                elif is_image and last_video_name:
                    # 이미지 파일 처리 (직전 동영상 파일명 기준으로 변경)
                    result = self.process_image_file(
                        idx, len(all_files), file_full_path, original_filename,
                        last_video_base_name, image_counter
                    )
                    
                    if result.get('success'):
                        file_changes.append((original_filename, result.get('new_filename')))
                        total_processed += 1
                        image_processed += 1
                        image_counter = result.get('image_counter')
                        
                elif is_image:
                    print(f"\n[{idx}/{len(all_files)}] [이미지] {original_filename}")
                    print(f"⚠️ 이전 비디오 파일이 없어 이름을 변경하지 않습니다.")
            
            if self.running:  # 정상 종료인 경우에만 완료 메시지 출력
                # 작업 결과 요약
                print("\n[ 작업 결과 요약 ]")
                print(f"• 총 파일: {len(all_files)}개")
                print(f"• 처리 완료: {total_processed}개")
                print(f"• 비디오: {video_processed}개")
                print(f"• 이미지: {image_processed}개")
                
                # 변경된 파일 목록 출력
                if file_changes:
                    print("\n✓ 변경된 파일:")
                    for i, (orig, new) in enumerate(file_changes, 1):
                        print(f"  {i}. {orig} → {new}")
                
        except Exception as e:
            print(f"\n❌ 오류 발생: {str(e)}")
            import traceback
            print(traceback.format_exc())  # 디버깅을 위한 스택 트레이스 출력
        finally:
            self.finish_process()
    
    def get_target_files(self, work_dir):
        """처리 대상 파일 목록 가져오기"""
        all_files = []
        
        if self.selected_files:
            # 선택한 파일 목록 사용
            for file_path in self.selected_files:
                if is_valid_video_file(file_path) or is_valid_image_file(file_path):
                    # 파일 경로가 작업 폴더 내에 있는지 확인
                    if os.path.dirname(file_path) == work_dir:
                        # 파일명만 사용
                        rel_path = os.path.basename(file_path)
                        all_files.append((rel_path, file_path))
                    else:
                        # 작업 폴더 외부의 파일은 상대 경로 계산
                        try:
                            rel_path = os.path.relpath(file_path, work_dir)
                            all_files.append((rel_path, file_path))
                        except ValueError:
                            # 다른 드라이브 등의 문제가 있으면 파일명만 사용
                            rel_path = os.path.basename(file_path)
                            all_files.append((rel_path, file_path))
        else:
            # 폴더 내 모든 파일 사용 (하위 폴더 제외)
            try:
                for f in os.listdir(work_dir):
                    file_path = os.path.join(work_dir, f)
                    if os.path.isfile(file_path) and (is_valid_video_file(file_path) or is_valid_image_file(file_path)):
                        all_files.append((f, file_path))
            except Exception as e:
                print(f"❌ 폴더 읽기 오류: {str(e)}")
        
        # 파일 수정 날짜 기준 내림차순 정렬
        all_files.sort(key=lambda x: os.path.getmtime(x[1]), reverse=True)
        return all_files
    
    def process_video_file(self, idx, total, file_rel_path, file_full_path, original_filename,
                          video_processor, analyzer, frame_times, 
                          file_changes, total_processed, video_processed,
                          last_video_name, last_video_base_name, image_counter):
        """비디오 파일 처리"""
        print(f"\n[{idx}/{total}] [비디오] {original_filename}")
        sys.stdout.flush()  # 명시적 flush 추가
        
        try:
            # 비디오에서 프레임 추출
            print(f"  - 프레임 추출 중... ({', '.join(map(str, frame_times))}초)")
            sys.stdout.flush()  # 명시적 flush 추가
            frame_paths = video_processor.extract_frames(file_rel_path, frame_times)
            
            if not frame_paths:
                print(f"❌ 프레임 추출 실패 - 비디오 파일이 손상되었거나 접근할 수 없습니다.")
                sys.stdout.flush()  # 명시적 flush 추가
                return None
            
            # 각 프레임 분석
            video_results = self.analyze_video_frames(frame_paths, analyzer)
            
            # 현재 비디오에서 가장 좋은 결과 선택
            if video_results:
                # 가장 많이 나온 결과 사용
                most_common = max(set(video_results), key=video_results.count)
                print(f"  ✓ 최종 결과: {most_common}")
                sys.stdout.flush()  # 명시적 flush 추가
                
                # 동영상 파일 이름 변경
                try:
                    # 파일 확장자 유지
                    filename, ext = os.path.splitext(file_full_path)
                    dir_path = os.path.dirname(file_full_path)
                    
                    # 새 파일 이름 생성 (형식: 101동 101호 급수 급수.mp4)
                    new_name = most_common.replace('[', '').replace(']', '')
                    new_path = os.path.join(dir_path, f"{new_name}{ext}")
                    
                    # 이미 같은 이름의 파일이 있는지 확인
                    if os.path.exists(new_path) and os.path.abspath(file_full_path) != os.path.abspath(new_path):
                        # 파일 이름에 번호 추가
                        base_name = new_name
                        counter = 1
                        while os.path.exists(os.path.join(dir_path, f"{base_name} {counter:02d}{ext}")):
                            counter += 1
                        new_name = f"{base_name} {counter:02d}"
                        new_path = os.path.join(dir_path, f"{new_name}{ext}")
                    
                    # 파일 이름 변경
                    os.rename(file_full_path, new_path)
                    new_filename = os.path.basename(new_path)
                    print(f"  ✓ 파일명 변경: {original_filename} > {new_filename}")
                    sys.stdout.flush()  # 명시적 flush 추가
                    
                    # 마지막 비디오 이름 저장 (확장자 제외)
                    last_video_name = new_name
                    last_video_base_name = new_name
                    image_counter = 0  # 이미지 카운터 초기화
                    
                    # 변경 결과 기록
                    file_changes.append((original_filename, new_filename))
                    
                    return {
                        'success': True,
                        'new_filename': new_filename,
                        'image_counter': image_counter,
                        'last_video_name': last_video_name,
                        'last_video_base_name': last_video_base_name
                    }
                except PermissionError as e:
                    print(f"❌ 파일 이름 변경 권한이 없습니다: {str(e)}")
                    sys.stdout.flush()  # 명시적 flush 추가
                except FileNotFoundError as e:
                    print(f"❌ 원본 파일을 찾을 수 없습니다: {str(e)}")
                    sys.stdout.flush()  # 명시적 flush 추가
                except Exception as e:
                    print(f"❌ 파일 이름 변경 실패: {str(e)}")
                    sys.stdout.flush()  # 명시적 flush 추가
            else:
                print(f"❌ 유효한 결과 없음 - 모든 프레임 분석에 실패했습니다. 다음 파일로 진행합니다.")
                sys.stdout.flush()  # 명시적 flush 추가
        
        except Exception as e:
            print(f"❌ 비디오 처리 중 오류: {str(e)}")
            import traceback
            print(f"❌ 오류 상세정보: {traceback.format_exc()}")
            sys.stdout.flush()  # 명시적 flush 추가
        
        return None
    
    def analyze_video_frames(self, frame_paths, analyzer):
        """비디오 프레임 분석"""
        video_results = []
        
        # 각 프레임 분석
        print(f"  - 프레임 분석 중... ({len(frame_paths)}개)")
        sys.stdout.flush()  # 명시적 flush 추가
        
        # 처리량 조절 - 프레임 제한 (최대 3개 프레임만 처리)
        if len(frame_paths) > 3:
            print(f"  ⚠️ 처리량 조절: {len(frame_paths)}개 프레임 중 3개만 분석합니다.")
            sys.stdout.flush()  # 명시적 flush 추가
            frame_paths = frame_paths[:3]
        
        for i, frame_path in enumerate(frame_paths, 1):
            if not self.running:
                break
            
            # 파일 존재 확인
            if not os.path.exists(frame_path):
                print(f"  ❌ 프레임 파일이 존재하지 않습니다: {frame_path}")
                sys.stdout.flush()  # 명시적 flush 추가
                continue
            
            # 분석 시도 (최대 5회, 이전 3회에서 증가)
            retry_count = 0
            max_retries = 5
            extracted_info = None
            base_delay = 2  # 기본 대기 시간 (초)
            
            print(f"    프레임 {i}/{len(frame_paths)} 분석 중...")
            sys.stdout.flush()  # 명시적 flush 추가
            
            while retry_count < max_retries and extracted_info is None:
                try:
                    # 지수 백오프 적용 (재시도마다 대기 시간 증가)
                    if retry_count > 0:
                        # 2^n 공식 적용 (2, 4, 8, 16, 32초)
                        current_delay = base_delay * (2 ** (retry_count - 1))
                        # 약간의 랜덤성 추가 (지터)
                        jitter = random.uniform(0.8, 1.2)
                        delay_with_jitter = current_delay * jitter
                        
                        print(f"    ⚠️ API 요청 재시도 {retry_count}/{max_retries} (대기: {delay_with_jitter:.2f}초)")
                        sys.stdout.flush()  # 명시적 flush 추가
                        time.sleep(delay_with_jitter)
                    
                    # API 호출
                    extracted_info = analyzer.analyze_image(frame_path)
                    
                    # 분석 결과 출력
                    if extracted_info:
                        print(f"    ✓ 분석 성공: {extracted_info}")
                        sys.stdout.flush()  # 명시적 flush 추가
                    else:
                        print(f"    ❌ 분석 결과가 없습니다.")
                        sys.stdout.flush()  # 명시적 flush 추가
                        retry_count += 1
                except Exception as e:
                    retry_count += 1
                    error_message = str(e)
                    
                    # 특정 오류 유형 감지 및 로깅
                    if "429" in error_message or "rate limit" in error_message.lower() or "too many requests" in error_message.lower():
                        print(f"    ❌ API 요청 한도 초과 (Rate Limit): {error_message}")
                    elif "timeout" in error_message.lower() or "connection" in error_message.lower():
                        print(f"    ❌ 네트워크 연결 오류: {error_message}")
                    elif "authentication" in error_message.lower() or "auth" in error_message.lower() or "key" in error_message.lower():
                        print(f"    ❌ 인증 오류: {error_message}")
                    else:
                        print(f"    ❌ 분석 오류: {error_message}")
                    sys.stdout.flush()  # 명시적 flush 추가
            
            # 최종 결과 처리
            if extracted_info:
                video_results.append(extracted_info)
            else:
                print(f"    ❌ 프레임 {i} 분석에 모든 시도가 실패했습니다.")
                sys.stdout.flush()  # 명시적 flush 추가
            
            # 처리량 조절 - 프레임 간 대기 시간 추가
            if i < len(frame_paths) and extracted_info:
                delay_between_frames = 1  # 프레임 간 2초 대기
                print(f"    ✓ 다음 프레임 분석 전 {delay_between_frames}초 대기 중...")
                sys.stdout.flush()  # 명시적 flush 추가
                time.sleep(delay_between_frames)
        
        # 결과 요약
        if video_results:
            print(f"  ✓ 프레임 분석 완료: {len(video_results)}/{len(frame_paths)} 성공")
        else:
            print(f"  ❌ 모든 프레임 분석에 실패했습니다.")
        sys.stdout.flush()  # 명시적 flush 추가
        
        return video_results
        
    def process_image_file(self, idx, total, file_full_path, original_filename, 
                          last_video_base_name, image_counter):
        """이미지 파일 처리"""
        print(f"\n[{idx}/{total}] [이미지] {original_filename}")
        
        try:
            # 파일 확장자 유지
            _, ext = os.path.splitext(file_full_path)
            dir_path = os.path.dirname(file_full_path)
            
            # 이미지 카운터 증가
            image_counter += 1
            
            # 새 파일 이름 생성 (형식: 101동 101호 급수 급수_1.jpg)
            new_image_name = f"{last_video_base_name}_{image_counter}"
            new_image_path = os.path.join(dir_path, f"{new_image_name}{ext}")
            
            # 이미 같은 이름의 파일이 있는지 확인
            if os.path.exists(new_image_path) and os.path.abspath(file_full_path) != os.path.abspath(new_image_path):
                counter = 1
                while os.path.exists(os.path.join(dir_path, f"{new_image_name}_{counter}{ext}")):
                    counter += 1
                new_image_path = os.path.join(dir_path, f"{new_image_name}_{counter}{ext}")
            
            # 파일 이름 변경
            os.rename(file_full_path, new_image_path)
            new_filename = os.path.basename(new_image_path)
            print(f"  ✓ 파일명 변경: {original_filename} > {new_filename}")
            
            return {
                'success': True,
                'new_filename': new_filename,
                'image_counter': image_counter
            }
        except PermissionError:
            print(f"❌ 파일 이름 변경 권한이 없습니다.")
        except FileNotFoundError:
            print(f"❌ 원본 파일을 찾을 수 없습니다.")
        except Exception as e:
            print(f"❌ 파일 이름 변경 실패: {str(e)}")
            
        return {'success': False}
    
    def finish_process(self):
        """작업 종료 처리"""
        # 상태 플래그 업데이트
        self.running = False
        self.is_running = False
        
        # 임시 폴더 삭제
        self.cleanup_temp_dir()
        
        # 작업폴더/output 폴더가 있다면 삭제
        try:
            work_dir = self.folder_path.get()
            output_dir = os.path.join(work_dir, "output")
            if os.path.exists(output_dir) and os.path.isdir(output_dir):
                shutil.rmtree(output_dir)
                print(f"✓ 작업폴더 내 output 폴더 삭제 완료")
        except Exception as e:
            print(f"⚠️ output 폴더 삭제 중 오류: {str(e)}")
        
        # 메모리 정리 및 COM 객체 정리
        try:
            import gc
            gc.collect()
            
            # COM 객체 정리 (Windows 환경)
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except:
                pass
        except:
            pass
        
        # UI 상태 업데이트 (메인 스레드에서 실행)
        self.root.after(0, lambda: self.update_ui_for_processing(False))
        
        # 작업 종료 로그
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"\n[{current_time}] 작업 종료")

    def cleanup_memory(self):
        """메모리 정리"""
        try:
            import gc
            gc.collect()
        except Exception:
            pass
            
    def cleanup_com_objects(self):
        """COM 객체 정리 (Windows 환경)"""
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass

    def process_in_thread(self, current_mode, work_dir, output_dir, frame_times):
        """별도 스레드에서 작업 실행"""
        try:
            # 파일명 변경 모드 작업 실행
            print(f"• 모드: 파일명 변경")
            print(f"• 분석 방법: Vision API + ChatGPT")
            print(f"• 프레임 시간: {', '.join(map(str, frame_times))}초")
            print(f"• 처리량 조절 적용됨")
            
            # 파일명 변경 모드 작업 실행
            self.process_videos(work_dir, output_dir, frame_times)
            
        except Exception as e:
            print(f"\n❌ 예상치 못한 오류 발생: {str(e)}")
            import traceback
            print(traceback.format_exc())  # 디버깅을 위한 스택 트레이스 출력

def main():
    """메인 함수: 애플리케이션 시작"""
    # 프로그램 시작 로그
    logger.info("FileRenamerX 애플리케이션 시작")
    
    # Tkinter 루트 윈도우 생성
    root = tk.Tk()
    
    try:
        # 애플리케이션 인스턴스 생성
        app = FileRenamerXApp(root)
        
        # 시스템 종료 시 표준 출력 복원을 위해 참조 저장
        old_stdout = sys.stdout
        
        # 메인 루프 시작
        root.mainloop()
    except Exception as e:
        # 예상치 못한 오류 처리
        logger.error(f"애플리케이션 실행 중 오류: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        
        # 메시지박스로 오류 표시
        messagebox.showerror("오류", f"애플리케이션 실행 중 오류가 발생했습니다:\n{str(e)}")
    finally:
        # 표준 출력 복원
        sys.stdout = old_stdout

if __name__ == "__main__":
    main()