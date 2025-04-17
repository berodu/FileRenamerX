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

# 기존 분석기 및 비디오 프로세서 임포트
from video_processor import VideoProcessor
from image_analyzer import GoogleVisionAnalyzer
from excel_processor import ExcelProcessor  # 새로 추가한 모듈
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
            self.update_timer = self.text_widget.after(100, self.update_text)
    
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
            while True:
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
                
                # 최대 메시지 수를 초과하면 오래된 메시지 제거 (메모리 관리)
                if self.message_count > self.max_messages:
                    self.text_widget.delete(1.0, 2.0)
                    self.message_count -= 1
                
                # 스크롤을 최신 메시지로 이동
                self.text_widget.see(tk.END)
                self.text_widget.configure(state='disabled')
                self.queue.task_done()
        except queue.Empty:
            # 큐가 비어있으면 일정 시간 후 다시 확인
            self.update_timer = self.text_widget.after(100, self.update_text)
    
    def _highlight_keywords(self, message):
        """메시지 내의 성공, 실패, 오류 등의 키워드 강조 처리"""
        # 현재 위치 (마지막에 삽입된 텍스트)
        current_line = self.text_widget.index(tk.END + "-1c linestart")
        
        # 성공 관련 키워드 강조
        success_keywords = ["성공", "[성공]", "완료", "처리 완료"]
        for keyword in success_keywords:
            start_pos = self.text_widget.search(keyword, current_line, tk.END)
            if start_pos:
                end_pos = f"{start_pos}+{len(keyword)}c"
                self.text_widget.tag_add("success_keyword", start_pos, end_pos)
        
        # 오류 관련 키워드 강조
        error_keywords = ["실패", "[실패]", "오류", "에러", "Error", "error"]
        for keyword in error_keywords:
            start_pos = self.text_widget.search(keyword, current_line, tk.END)
            if start_pos:
                end_pos = f"{start_pos}+{len(keyword)}c"
                self.text_widget.tag_add("error_keyword", start_pos, end_pos)
        
        # 경고 관련 키워드 강조
        warning_keywords = ["건너뜀", "[건너뜀]", "경고", "주의"]
        for keyword in warning_keywords:
            start_pos = self.text_widget.search(keyword, current_line, tk.END)
            if start_pos:
                end_pos = f"{start_pos}+{len(keyword)}c"
                self.text_widget.tag_add("warning_keyword", start_pos, end_pos)
    
    def flush(self):
        """파이썬 출력 스트림 호환을 위한 메서드"""
        pass
        
    def clear(self):
        """텍스트 위젯의 내용을 지움"""
        self.text_widget.configure(state='normal')
        self.text_widget.delete(1.0, tk.END)
        self.text_widget.configure(state='disabled')
        self.message_count = 0

class TaskieXApp:
    """TaskieX 애플리케이션 메인 클래스"""
    def __init__(self, root):
        self.root = root
        self.root.title("TaskieX")
        self.root.geometry("700x900")  # 창 크기 설정
        self.root.minsize(700, 600)    # 최소 창 크기 제한
        
        # 상태 변수 초기화
        self.process_thread = None     # 작업 스레드
        self.running = False           # 실행 상태
        self.is_running = False        # 중복 방지용 상태 플래그
        self.temp_dir = None           # 임시 디렉토리 경로
        self.selected_files = []       # 선택된 파일 목록

        # UI 변수 초기화
        self.excel_path = tk.StringVar(value="")
        self.work_mode = tk.StringVar(value="rename")  # 기본값: 파일명 변경 모드
        self.folder_path = tk.StringVar(value="./작업폴더")
        
        # 작업 설정 초기화
        self.frame_times_value = [2, 3, 5]  # 기본 프레임 시간 (초)
        
        # 엑셀 프로세서 초기화
        self.excel_processor = ExcelProcessor()

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

        # 1. 작업 모드 선택 프레임
        self.create_mode_frame(main_frame)
        
        # 2. 설정 프레임
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

    def create_mode_frame(self, parent):
        """작업 모드 선택 프레임 생성"""
        mode_frame = ttk.LabelFrame(parent, text="작업 모드", padding=10)
        mode_frame.pack(fill=tk.X, pady=5)
        
        # 모드 선택 라디오 버튼
        ttk.Radiobutton(
            mode_frame, 
            text="파일명 변경", 
            variable=self.work_mode, 
            value="rename", 
            command=self.toggle_mode
        ).pack(side=tk.LEFT, padx=10)
        
        ttk.Radiobutton(
            mode_frame, 
            text="이상 배관 업데이트", 
            variable=self.work_mode, 
            value="update_pipe", 
            command=self.toggle_mode
        ).pack(side=tk.LEFT, padx=10)
        
        ttk.Radiobutton(
            mode_frame, 
            text="작업 현황 업데이트", 
            variable=self.work_mode, 
            value="update_status", 
            command=self.toggle_mode
        ).pack(side=tk.LEFT, padx=10)

    def create_settings_frame(self, parent):
        """설정 프레임 생성"""
        self.settings_frame = ttk.LabelFrame(parent, text="설정", padding=10)
        self.settings_frame.pack(fill=tk.X, pady=5)

        # 파일 경로 설정 프레임
        self.path_frame = ttk.Frame(self.settings_frame)
        self.path_frame.pack(fill=tk.X, pady=5)
        
        # 작업 폴더 선택 UI
        self.folder_button_frame = ttk.Frame(self.path_frame)
        self.folder_button_frame.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        ttk.Button(
            self.folder_button_frame, 
            text="작업 폴더 선택", 
            command=self.browse_folder
        ).pack(side=tk.LEFT, padx=5)
        
        self.path_label = ttk.Label(self.folder_button_frame, text="./작업폴더")
        self.path_label.pack(side=tk.LEFT, padx=5)
        
        # 엑셀 파일 선택 UI (초기에는 숨김)
        self.excel_frame = ttk.Frame(self.path_frame)
        self.excel_frame.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        
        ttk.Button(
            self.excel_frame, 
            text="엑셀 파일 선택", 
            command=self.browse_excel
        ).pack(side=tk.LEFT, padx=5)
        
        self.excel_label = ttk.Label(self.excel_frame, text="선택되지 않음")
        self.excel_label.pack(side=tk.LEFT, padx=5)
        
        # 초기 모드에 따라 엑셀 프레임 표시/숨김
        if self.work_mode.get() in ["update_pipe", "update_status"]:
            self.excel_frame.grid()
        else:
            self.excel_frame.grid_remove()

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
        self.update_help_text()

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

    def update_help_text(self):
        """작업 모드에 따라 도움말 텍스트 업데이트"""
        self.help_text.config(state=tk.NORMAL)
        self.help_text.delete(1.0, tk.END)
        
        mode = self.work_mode.get()
        
        if mode == "rename":
            # 파일명 변경 모드 도움말
            self.help_text.insert(tk.END, "- 파일명 변경 모드\n", "bold")
            self.help_text.insert(tk.END, "작업폴더 : 파일명을 변경할 동영상, 이미지 파일이 저장되어 있는 폴더\n")
            self.help_text.insert(tk.END, "동영상 파일은 Vision API와 ChatGPT로 분석하여 파일명이 변경됩니다.\n")
            self.help_text.insert(tk.END, "이미지 파일은 직전에 처리된 동영상 파일명을 기준으로 이름이 변경됩니다.")
        elif mode == "update_pipe":
            # 이상 배관 업데이트 모드 도움말
            self.help_text.insert(tk.END, "- 이상 배관 업데이트 모드\n", "bold")
            self.help_text.insert(tk.END, "작업폴더 : 엑셀 파일에 삽입할 이미지 파일이 저장되어 있는 폴더\n")
            self.help_text.insert(tk.END, "엑셀파일 : 보고서 엑셀 파일\n")
            self.help_text.insert(tk.END, "이미지 파일명 : '[동] [호] [배관종류] [배관명]_[이상소견]_[이상위치]' 형식\n")
            self.help_text.insert(tk.END, "결과 : '2.이상배관위치'와 '3.이상배관LIST' 시트에 작성")
        else:  # update_status 모드
            # 작업 현황 업데이트 모드 도움말
            self.help_text.insert(tk.END, "- 작업 현황 업데이트 모드\n", "bold")
            self.help_text.insert(tk.END, "작업폴더 : 분석할 동영상 파일이 저장되어 있는 폴더\n")
            self.help_text.insert(tk.END, "엑셀파일 : 작업 현황을 업데이트할 엑셀 파일\n")
            self.help_text.insert(tk.END, "동영상 파일명 : '[동] [호] [배관종류] [배관명]' 형식\n")
            self.help_text.insert(tk.END, "결과 : '1.작업현황_[배관종류]' 시트에 작업 완료 표시")
            
        self.help_text.config(state=tk.DISABLED)

    def toggle_mode(self):
        """작업 모드 전환 시 UI 업데이트"""
        mode = self.work_mode.get()
        
        # 로그 초기화
        if hasattr(self, 'redirect'):
            self.redirect.clear()
        
        # 도움말 텍스트 업데이트
        self.update_help_text()
        
        # 현재 작업 폴더 경로
        current_folder = self.folder_path.get()
        
        # 이상 배관 업데이트 모드 또는 작업 현황 업데이트 모드
        if mode in ["update_pipe", "update_status"]:
            # 엑셀 파일 선택 프레임 표시
            self.excel_frame.grid()
            # 선택된 파일 목록 초기화
            self.selected_files = []
            
            if mode == "update_pipe":
                # 상태 메시지 업데이트
                print("이상 배관 업데이트 모드로 전환되었습니다.")
                print("작업 폴더의 이미지 파일들을 엑셀에 삽입합니다.")
                print("작업 폴더와 엑셀 파일을 선택해주세요.")
                
                # 현재 폴더가 설정되어 있으면 이미지 파일 목록 표시
                if os.path.exists(current_folder):
                    self.display_image_files(current_folder)
            else:  # update_status 모드
                # 상태 메시지 업데이트
                print("작업 현황 업데이트 모드로 전환되었습니다.")
                print("작업 폴더의 동영상 파일을 분석하여 엑셀의 작업 현황을 업데이트합니다.")
                print("작업 폴더와 엑셀 파일을 선택해주세요.")
                
                # 현재 폴더가 설정되어 있으면 동영상 파일 목록 표시
                if os.path.exists(current_folder):
                    self.display_video_files(current_folder)
        
        # 파일명 변경 모드
        else:
            # 엑셀 파일 선택 프레임 숨기기
            self.excel_frame.grid_remove()
            # 상태 메시지 업데이트
            print("파일명 변경 모드로 전환되었습니다.")
            print("작업 폴더의 비디오/이미지 파일 이름을 Vision API+ChatGPT로 분석하여 변경합니다.")
            print("작업 폴더를 선택해주세요.")
            
            # 현재 폴더가 설정되어 있으면 대상 파일 목록 표시
            if os.path.exists(current_folder):
                self.display_target_files(current_folder)

    def browse_folder(self):
        """작업 폴더 선택"""
        folder = filedialog.askdirectory(title="작업 폴더를 선택하세요")
        if not folder:
            return
            
        self.folder_path.set(folder)
        self.selected_files = []  # 폴더를 선택하면 개별 파일 선택 초기화
        self.path_label.config(text=f"{folder}")
        
        # 현재 모드 확인
        current_mode = self.work_mode.get()
        
        # 폴더 내 파일 목록 표시 (모드에 따라 다르게)
        if current_mode == "rename":
            # 파일명 변경 모드에서는 비디오/이미지 파일 표시
            self.display_target_files(folder)
        elif current_mode == "update_pipe":
            # 이상 배관 업데이트 모드에서는 이미지 파일만 표시
            self.display_image_files(folder)
        elif current_mode == "update_status":
            # 작업 현황 업데이트 모드에서는 동영상 파일만 표시
            self.display_video_files(folder)
    
    def browse_excel(self):
        """엑셀 파일 선택"""
        if self.work_mode.get() not in ["update_pipe", "update_status"]:
            return
        
        file_path = filedialog.askopenfilename(
            title="엑셀 파일을 선택하세요",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        self.excel_path.set(file_path)
        self.excel_label.config(text=f"{os.path.basename(file_path)}")
        
        # 모드에 따라 다른 메시지 출력
        mode = self.work_mode.get()
        if mode == "update_pipe":
            print(f"✓ 선택된 엑셀 파일: {os.path.basename(file_path)}")
        elif mode == "update_status":
            print(f"✓ 선택된 엑셀 파일: {os.path.basename(file_path)}")
            print(f"📋 작업 현황 업데이트 모드 안내:")
            print(f"  • 동영상 파일명은 '[동] [호] [배관종류] [배관명]' 형식이어야 합니다.")
            print(f"  • 예: '103동 1호 입상관 공용오수.mp4', '102동 1903호 세대매립관 세탁.mp4'")
            print(f"  • 지원되는 배관종류: 입상관, 세대매립관, 세대PD, 세대층상배관, 횡주관")
            print(f"  • 엑셀 파일의 '1.작업현황_[배관종류]' 시트에 작업 현황이 표시됩니다.")

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
        # 이미 실행 중인지 확인
        if self.is_running:
            return
        
        # 현재 작업 모드
        current_mode = self.work_mode.get()
        
        # 작업 폴더 확인
        work_dir = self.folder_path.get().strip()
        if not work_dir or not os.path.exists(work_dir):
            messagebox.showerror("오류", "작업 폴더가 존재하지 않습니다.")
            return
        
        # 프레임 시간 설정 (파일명 변경 모드에서만 사용)
        frame_times = self.frame_times_value
        
        # 파일명 변경 모드
        if current_mode == "rename":
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
            if not check_api_keys():
                messagebox.showerror("오류", "API 키가 설정되지 않았거나 유효하지 않습니다.")
                return
                
            # 프레임 시간 검증
            if not validate_frame_times(frame_times):
                messagebox.showerror("오류", "프레임 시간이 올바르지 않습니다.")
                return
                
            # 임시 디렉토리 생성
            try:
                self.cleanup_temp_dir()  # 기존 임시 폴더 정리
                self.temp_dir = tempfile.mkdtemp(prefix="taskiex_temp_")
                output_dir = self.temp_dir
                logger.info(f"임시 폴더 생성: {self.temp_dir}")
            except Exception as e:
                messagebox.showerror("오류", f"임시 폴더 생성 실패: {str(e)}")
                return
            
        else:  # update_pipe 또는 update_status 모드
            # 엑셀 파일 선택 확인
            excel_path = self.excel_path.get()
            if not excel_path:
                messagebox.showerror("오류", "엑셀 파일을 선택해주세요.")
                return
            
            if not os.path.exists(excel_path):
                messagebox.showerror("오류", f"선택한 엑셀 파일이 존재하지 않습니다: {excel_path}")
                return
                
            # 엑셀 파일이 열려있는지 확인하고 닫기
            print(f"✓ 엑셀 파일 확인 중: {os.path.basename(excel_path)}")
            self.close_excel_file(excel_path)
        
        # UI 상태 업데이트
        self.update_ui_for_processing(True)
        
        # 로그 초기화
        if hasattr(self, 'redirect'):
            self.redirect.clear()
        
        # 시작 로그 출력
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{current_time}] 작업 시작")
        
        # 모드별 작업 스레드 시작
        if current_mode == "rename":
            print(f"• 모드: 파일명 변경")
            print(f"• 분석 방법: Vision API + ChatGPT")
            print(f"• 프레임 시간: {','.join(map(str, frame_times))}초")
            print("")
            
            # 파일명 변경 모드 작업 스레드 시작
            self.process_thread = threading.Thread(
                target=self.process_videos, 
                args=(work_dir, output_dir, frame_times)
            )
        elif current_mode == "update_pipe":  # 이상 배관 업데이트 모드
            print(f"• 모드: 이상 배관 업데이트")
            print(f"• 엑셀 파일: {os.path.basename(self.excel_path.get())}")
            print("")
            
            # 이미지 삽입 모드 작업 스레드 시작
            self.process_thread = threading.Thread(
                target=self.process_excel, 
                args=(work_dir, self.excel_path.get())
            )
        else:  # update_status 모드 (작업 현황 업데이트)
            print(f"• 모드: 작업 현황 업데이트")
            print(f"• 엑셀 파일: {os.path.basename(self.excel_path.get())}")
            print("")
            
            # 작업 현황 업데이트 모드 작업 스레드 시작
            self.process_thread = threading.Thread(
                target=self.update_status_excel, 
                args=(work_dir, self.excel_path.get())
            )
        
        # 스레드 데몬 설정 및 시작
        self.process_thread.daemon = True
        self.process_thread.start()

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
                        file_changes.append((original_filename, result.get('new_filename')))
                        total_processed += 1
                        video_processed += 1
                        image_counter = result.get('image_counter')
                        
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
        
        try:
            # 비디오에서 프레임 추출
            print(f"  - 프레임 추출 중... ({', '.join(map(str, frame_times))}초)")
            frame_paths = video_processor.extract_frames(file_rel_path, frame_times)
            
            if not frame_paths:
                print(f"❌ 프레임 추출 실패")
                return None
            
            # 각 프레임 분석
            video_results = self.analyze_video_frames(frame_paths, analyzer)
            
            # 현재 비디오에서 가장 좋은 결과 선택
            if video_results:
                # 가장 많이 나온 결과 사용
                most_common = max(set(video_results), key=video_results.count)
                print(f"  ✓ 최종 결과: {most_common}")
                
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
                    
                    # 마지막 비디오 이름 저장 (확장자 제외)
                    last_video_name = new_name
                    last_video_base_name = new_name
                    image_counter = 0  # 이미지 카운터 초기화
                    
                    # 변경 결과 기록
                    file_changes.append((original_filename, new_filename))
                    total_processed += 1
                    video_processed += 1
                    
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
            else:
                print(f"❌ 유효한 결과 없음")
        
        except Exception as e:
            print(f"❌ 비디오 처리 중 오류: {str(e)}")
            
        return None
    
    def analyze_video_frames(self, frame_paths, analyzer):
        """비디오 프레임 분석"""
        video_results = []
        
        # 각 프레임 분석
        print(f"  - 프레임 분석 중... ({len(frame_paths)}개)")
        for i, frame_path in enumerate(frame_paths, 1):
            if not self.running:
                break
                
            # 파일 존재 확인
            if not os.path.exists(frame_path):
                continue
            
            # 분석 시도 (최대 3회)
            retry_count = 0
            max_retries = 3
            extracted_info = None
            
            while retry_count < max_retries and extracted_info is None:
                if retry_count > 0:
                    print(f"    재시도 중... ({retry_count}/{max_retries})")
                    time.sleep(2)  # API 호출 간 딜레이
                
                try:
                    print(f"    프레임 {i}/{len(frame_paths)} 분석 중...")
                    extracted_info = analyzer.analyze_image(frame_path)
                except Exception as e:
                    print(f"❌ 분석 오류: {str(e)}")
                    retry_count += 1
                    continue
                
                # 분석 결과 출력
                if extracted_info:
                    print(f"    결과: {extracted_info}")
                
                retry_count += 1
            
            if extracted_info:
                video_results.append(extracted_info)
            else:
                print(f"❌ 분석 실패")
                
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

    def process_excel(self, image_folder, excel_path):
        """이미지 삽입 모드: 이미지 폴더의 이미지를 엑셀에 삽입"""
        try:
            # 이미지 파일 수 확인
            image_files = [f for f in os.listdir(image_folder) 
                          if os.path.isfile(os.path.join(image_folder, f)) and 
                          is_valid_image_file(os.path.join(image_folder, f))]
            
            total_images = len(image_files)
            if total_images == 0:
                print("❌ 처리할 이미지 파일이 없습니다.")
                self.finish_process()
                return
            
            print(f"✓ 총 {total_images}개 이미지 파일을 처리합니다.")
            print(f"✓ 엑셀 파일에 이미지 삽입 작업을 시작합니다...")
            
            # 로그 수준 조정을 위한 간략화된 로그 함수
            def log_func(message, level="info"):
                # 중요 메시지만 출력 (에러, 경고, 주요 단계)
                if level in ["error", "warning"] or message.startswith("✓") or message.startswith("❌"):
                    print(message)
            
            # 메모리 정리
            self.cleanup_memory()
            
            # 엑셀 처리기 호출 전에 엑셀 프로세스 다시 한번 확인
            self.close_excel_file(excel_path)
            
            # 엑셀 처리기 호출 (로그 간략화 함수 전달)
            result = self.excel_processor.process_images(excel_path, image_folder, log_func)
            
            if not result["success"]:
                print(f"\n❌ 엑셀 처리 중 오류 발생: {result.get('error', '알 수 없는 오류')}")
            else:
                self.display_excel_result(result)
            
            # COM 객체 정리 및 메모리 정리
            self.cleanup_com_objects()
            self.cleanup_memory()
            
            # 작업 완료 후 한번 더 엑셀 프로세스 종료 확인
            time.sleep(0.5)  # 잠시 대기 후 프로세스 확인
            self.close_excel_file(excel_path)
            
        except Exception as e:
            print(f"\n❌ 예상치 못한 오류 발생: {str(e)}")
            import traceback
            print(traceback.format_exc())  # 디버깅을 위한 스택 트레이스 출력
        finally:
            self.finish_process()
            
    def display_excel_result(self, result):
        """엑셀 처리 결과 표시"""
        # 이미지 처리 결과 요약 표시
        processed_count = len(result.get('processed', []))
        skipped_count = len(result.get('skipped', []))
        
        print(f"\n✓ 엑셀 작업 완료!")
        print(f"  • 총 이미지: {result.get('total', 0)}개")
        print(f"  • 처리 성공: {processed_count}개")
        print(f"  • 처리 실패: {skipped_count}개")
        
        # 성공한 이미지 모두 표시
        if processed_count > 0:
            print("\n✓ 처리된 이미지:")
            for i, img_info in enumerate(result.get('processed', []), 1):
                # img_info가 문자열인 경우 바로 출력, 딕셔너리인 경우 경로 정보 추출
                if isinstance(img_info, str):
                    img_name = os.path.basename(img_info)
                else:
                    img_name = os.path.basename(img_info.get('image_path', ''))
                print(f"  {i}. {img_name}")
        
        # 실패한 이미지가 있으면 모두 표시
        if skipped_count > 0:
            print("\n⚠️ 처리 실패한 이미지:")
            for i, img_info in enumerate(result.get('skipped', []), 1):
                # img_info가 문자열인 경우 바로 출력, 딕셔너리인 경우 경로와 이유 정보 추출
                if isinstance(img_info, str):
                    print(f"  {i}. {os.path.basename(img_info)}")
                else:
                    img_name = os.path.basename(img_info.get('image_path', ''))
                    reason = img_info.get('reason', '알 수 없는 이유')
                    print(f"  {i}. {img_name} - {reason}")
            
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

    def close_excel_file(self, file_path):
        """
        지정된 엑셀 파일이 열려있다면 해당 프로세스만 종료하고, 
        모든 엑셀 관련 리소스를 정리합니다.
        
        Args:
            file_path (str): 확인할 엑셀 파일의 경로
        
        Returns:
            bool: 프로세스가 종료되었으면 True, 아니면 False
        """
        if not file_path or not os.path.exists(file_path):
            return False
        
        # 엑셀 프로세스 종료 (ExcelProcessor의 메서드 사용)
        try:
            # 로그 출력 없이 프로세스만 종료하기 위해 콜백 함수를 None으로 설정
            return self.excel_processor.terminate_excel_processes(file_path, callback=None)
        except Exception as e:
            print(f"❌ 엑셀 파일 처리 중 오류: {str(e)}")
            return False

    def update_status_excel(self, video_folder, excel_path):
        """작업 현황 업데이트 모드: 동영상 파일 정보로 엑셀의 작업 현황 업데이트"""
        try:
            from openpyxl import load_workbook
            from openpyxl.styles import PatternFill, Font
            import re
            
            # 각 파일별 처리 결과 추적용 딕셔너리 초기화
            file_results = {}  # {파일명: {"status": "성공"|"실패", "pipe_type": 배관종류, "reason": 실패 이유}}
            pipe_type_results = {}  # {배관종류: {"files": [성공한 파일들], "failed": [실패한 파일들 + 이유]}}
            
            # 처리된 파일 추적
            processed_files = set()
            
            # 동영상 파일 수 확인
            video_file_paths = []
            video_file_names = []
            for f in os.listdir(video_folder):
                file_path = os.path.join(video_folder, f)
                if os.path.isfile(file_path) and is_valid_video_file(file_path):
                    video_file_paths.append(file_path)
                    video_file_names.append(f)
            
            total_videos = len(video_file_paths)
            if total_videos == 0:
                print("❌ 처리할 동영상 파일이 없습니다.")
                self.finish_process()
                return
            
            print(f"✓ 총 {total_videos}개 동영상 파일을 처리합니다.")
            print("✓ 발견된 동영상 파일:")
            for i, fname in enumerate(video_file_names[:5], 1):  # 처음 5개만 출력
                print(f"  {i}. {fname}")
            if total_videos > 5:
                print(f"  ... 외 {total_videos - 5}개")
            
            print(f"✓ 엑셀 파일 작업 현황 업데이트를 시작합니다...")
            
            # 메모리 정리
            self.cleanup_memory()
            
            # 엑셀 처리 전에 엑셀 프로세스 확인
            self.close_excel_file(excel_path)
            
            # 엑셀 파일 로드
            print(f"✓ 엑셀 파일 로드 중: {os.path.basename(excel_path)}")
            wb = load_workbook(excel_path)
            
            # 스타일 정의
            blue_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')  # 하늘색
            yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')  # 노란색
            black_font = Font(color='000000', size=10)  # 검정색, 크기 10
            
            # 배관종류 리스트 하드코딩
            PIPE_TYPE_LIST = ["입상관", "세대매립관", "세대PD", "세대층상배관", "횡주관"]
            # 처리 제외할 배관종류
            EXCLUDED_PIPE_TYPES = []  # 모든 배관종류 처리
            
            # 입력창 시트 확인
            if '입력창' not in wb.sheetnames:
                print("❌ 엑셀 파일에 '입력창' 시트가 없습니다.")
                self.finish_process()
                return
                
            ws_input = wb['입력창']
            
            # 1. 모든 배관종류와 범례 추출 (배관종류 -> {배관명 -> 번호})
            pipe_types = {}  # 키: 배관종류, 값: {배관명 -> 번호} 사전
            
            print("✓ 배관종류 및 범례 정보 추출 중...")
            
            # 입력창 7행에서 하드코딩된 배관종류 찾기
            col = 1
            while col <= ws_input.max_column:
                cell = ws_input.cell(row=7, column=col)
                cell_value = str(cell.value).strip() if cell.value else ""
                
                # 하드코딩된 배관종류 리스트에 있고 제외 리스트에 없는 경우만 처리
                if cell_value in PIPE_TYPE_LIST and cell_value not in EXCLUDED_PIPE_TYPES:
                    pipe_type = cell_value
                    pipe_start_col = col
                    pipe_end_col = None
                    
                    # 병합셀 범위 확인
                    for merged_cell in ws_input.merged_cells.ranges:
                        if cell.coordinate in merged_cell:
                            pipe_end_col = merged_cell.max_col
                            break
                    
                    # 병합셀 범위가 확인되지 않았다면 다음 열을 확인하여 범위 추정
                    if not pipe_end_col:
                        next_col = pipe_start_col + 1
                        while next_col <= ws_input.max_column:
                            next_cell_value = ws_input.cell(row=7, column=next_col).value
                            if next_cell_value and next_cell_value != cell.value:
                                pipe_end_col = next_col - 1
                                break
                            next_col += 1
                        
                        # 마지막 열까지 모두 동일한 값이면
                        if not pipe_end_col:
                            pipe_end_col = ws_input.max_column
                    
                    # 해당 배관종류의 범례 추출
                    legend = {}
                    number = 1
                    for legend_col in range(pipe_start_col, pipe_end_col + 1):
                        val = ws_input.cell(row=8, column=legend_col).value
                        if val is not None:
                            legend[val] = str(number)
                            number += 1
                    
                    # 배관종류와 범례 저장
                    pipe_types[pipe_type] = legend
                    
                    # 다음 검색을 위해 열 위치 업데이트
                    col = pipe_end_col + 1
                else:
                    col += 1
            
            # 배관종류 및 범례 정보 출력
            print("✓ 추출된 배관종류 및 범례:")
            for pipe_type, legend in pipe_types.items():
                print(f"  • {pipe_type}: {legend}")
            
            # 결과 저장용 변수
            updated_count = 0
            total_updated_cells = 0
            updated_pipe_types = set()
            
            # 3. 각 배관종류별로 처리
            for pipe_type, pipe_legend in pipe_types.items():
                # 해당 배관종류의 작업현황 시트 찾기
                sheet_name = f"1.작업현황_{pipe_type.strip()}"
                if sheet_name not in wb.sheetnames:
                    print(f"❌ 시트를 찾을 수 없음: {sheet_name}")
                    continue
                
                ws_target = wb[sheet_name]
                print(f"\n✓ 처리 중인 시트: {sheet_name}")
                
                # 입상관인 경우 특별 처리
                if pipe_type.strip() == "입상관":
                    print("  • 입상관 처리 시작")
                    
                    # 동 및 라인 열 인덱스 설정
                    building_col = 2  # B열
                    line_col = 3      # C열
                    
                    # 3행에서 배관명 추출
                    pipe_names = []
                    column_to_pipe = {}  # 열 인덱스 -> 배관명 매핑
                    
                    # 입상관 범례에서 배관명 목록 가져오기
                    valid_pipe_names = list(pipe_legend.keys())
                    
                    # 3행에서 실제 배관명만 추출
                    for col_idx in range(1, ws_target.max_column + 1):
                        cell_value = ws_target.cell(row=3, column=col_idx).value
                        if cell_value and str(cell_value).strip() in valid_pipe_names:
                            pipe_names.append(cell_value)
                            column_to_pipe[col_idx] = cell_value
                    
                    # 동영상 파일에서 입상관 정보 추출
                    riser_inspections = {}  # 키: (동, 라인), 값: set(배관명)
                    
                    # 파일별 입상관 처리 정보 저장용 변수
                    processed_riser_files = {}  # 파일명: {"status": 성공여부, "building": 동, "line": 라인, "pipe": 배관명}
                    
                    # 입상관 파일 정보 추출
                    for fname in video_file_names:
                        # 파일명에서 정보 추출 (예: "102동 1903호 입상관 온수.mp4")
                        m = re.match(r"(\d+)동\s*(\d+)호\s*(.+?)\s*(\S+)\.mp4", fname)
                        if not m:
                            print(f"  ⚠️ 파일명 형식이 맞지 않음: {fname}")
                            # 파일 처리 실패 기록
                            file_results[fname] = {
                                "status": "실패",
                                "pipe_type": "입상관",
                                "reason": "파일명 형식이 맞지 않음"
                            }
                            processed_files.add(fname)
                            continue
                        
                        building = int(m.group(1))
                        unit_str = m.group(2)  # 예: "1903"
                        video_pipe_type = m.group(3)  # 배관종류 (예: "입상관")
                        pipe_name = m.group(4)  # 배관명 (예: "온수")
                        
                        # 배관명에서 괄호 이후 부분 제거
                        if '(' in pipe_name:
                            pipe_name = pipe_name.split('(')[0].strip()
                            print(f"  • 괄호 처리: 배관명을 '{pipe_name}'으로 정리")
                        
                        print(f"  • 파일 분석: 동={building}, 호={unit_str}, 배관종류={video_pipe_type}, 배관명={pipe_name}")
                        
                        # 현재 처리 중인 배관종류와 일치하는지 확인
                        if pipe_type.strip() not in video_pipe_type:
                            # 배관종류 불일치 로그를 출력하지 않음
                            # 처리된 파일 표시
                            processed_files.add(fname)
                            continue
                            
                        # 호수에서 라인 추출 (마지막 두 자리)
                        line = None
                        try:
                            if len(unit_str) >= 3:
                                line = int(unit_str[-2:])  # 마지막 두 자리
                            elif len(unit_str) == 2:
                                line = int(unit_str)       # 두 자리 전체
                            elif len(unit_str) == 1:
                                line = int(unit_str)       # 한 자리 전체
                            else:
                                # 파일 처리 실패 기록
                                file_results[fname] = {
                                    "status": "실패",
                                    "pipe_type": pipe_type,
                                    "reason": f"호수 형식 오류: {unit_str}"
                                }
                                # 처리된 파일 표시
                                processed_files.add(fname)
                                continue
                        except ValueError:
                            print(f"  ⚠️ 호수에서 라인 추출 실패: {unit_str}")
                            # 파일 처리 실패 기록
                            file_results[fname] = {
                                "status": "실패",
                                "pipe_type": pipe_type,
                                "reason": f"호수에서 라인 추출 실패: {unit_str}"
                            }
                            # 처리된 파일 표시
                            processed_files.add(fname)
                            continue
                        
                        # 파일 정보 임시 저장 (아직 성공/실패 결정 안됨)
                        processed_riser_files[fname] = {
                            "building": building,
                            "line": line,
                            "pipe": pipe_name,
                            "status": False  # 아직 처리 안됨
                        }
                        
                        key = (building, line)
                        if key not in riser_inspections:
                            riser_inspections[key] = set()
                        
                        riser_inspections[key].add(pipe_name)
                        print(f"  ✓ 추출 성공: 동={building}, 라인={line}, 배관={pipe_name}")
                    
                    # 배관종류별 결과 추가
                    if pipe_type not in pipe_type_results:
                        pipe_type_results[pipe_type] = {"files": [], "failed": []}
                    
                    # 시트에서 행-동-라인 매핑
                    all_rows_info = []  # 모든 행의 정보를 저장할 리스트
                    
                    # 병합 셀 처리를 위한 이전 동/라인 값 저장
                    prev_building_val = None
                    prev_line_val = None
                    
                    # 시트의 모든 동-라인 값 로깅 (디버깅용)
                    print(f"  • 입상관 시트 동-라인 정보 분석 시작 (행 4 ~ {ws_target.max_row}):")
                    
                    for row_idx in range(4, ws_target.max_row + 1):
                        building_val = ws_target.cell(row=row_idx, column=building_col).value
                        line_val = ws_target.cell(row=row_idx, column=line_col).value
                        
                        # 디버깅용 원본 값 출력
                        if building_val is not None or line_val is not None:
                            print(f"    행 {row_idx}: 원본 동={building_val}, 라인={line_val}")
                        
                        # 병합 셀 처리: 값이 None이면 이전 값 사용
                        if building_val is None:
                            building_val = prev_building_val
                        else:
                            prev_building_val = building_val
                        
                        if line_val is None:
                            line_val = prev_line_val
                        else:
                            prev_line_val = line_val
                        
                        if not building_val or not line_val:
                            continue
                        
                        # 문자열로 변환하고 숫자만 추출
                        building_str = str(building_val).strip()
                        line_str = str(line_val).strip()
                        
                        # 디버깅용 변환 후 값 출력
                        print(f"    행 {row_idx}: 처리 후 동={building_str}, 라인={line_str}")
                        
                        # 숫자 추출 (여러 숫자 패턴 시도)
                        # 특별히 101동 처리 - 1동이 101동을 의미할 수 있음
                        if building_str == "1동" or building_str == "1" or "1동" in building_str:
                            building = 101
                            print(f"    ✓ 특별 처리: 1동을 101동으로 인식")
                        else:
                            building_match = re.search(r'\d+', building_str)
                            if not building_match:
                                print(f"    ⚠️ 숫자 추출 실패: 동={building_str}, 라인={line_str}")
                                continue
                            building = int(building_match.group())
                        
                        line_match = re.search(r'\d+', line_str)
                        if not line_match:
                            print(f"    ⚠️ 숫자 추출 실패: 동={building_str}, 라인={line_str}")
                            continue
                        
                        line = int(line_match.group())
                        
                        print(f"    ✓ 추출 성공: 동={building}, 라인={line}")
                        all_rows_info.append((row_idx, building, line))
                    
                    # 디버깅용 - 모든 찾은 동-라인 정보 상세 출력
                    print(f"  • 시트에서 추출한 행-동-라인 정보 ({len(all_rows_info)}개):")
                    for row_idx, building, line in all_rows_info:
                        print(f"    행 {row_idx}: 동={building}, 라인={line}")
                    
                    # 찾을 동-라인 정보 로깅
                    print(f"  • 파일에서 추출한 동-라인 정보 ({len(riser_inspections)}개):")
                    for (building, line), pipes in riser_inspections.items():
                        pipe_names_str = ", ".join(pipes)
                        print(f"    동={building}, 라인={line}, 배관={pipe_names_str}")
                    
                    # 시트와 파일 정보 간 매칭 시도
                    matched_count = 0
                    
                    # 각 동-라인에 대한 검사 정보 처리
                    for key, inspected_pipes in riser_inspections.items():
                        building, line = key
                        
                        # 해당 동-라인이 있는 행 찾기 (더 유연한 매칭 시도)
                        matching_rows = []
                        
                        # 정확한 매칭
                        exact_matches = [row_idx for row_idx, bldg, ln in all_rows_info if bldg == building and ln == line]
                        if exact_matches:
                            matching_rows = exact_matches
                            print(f"  ✓ 정확한 매치 발견: 동={building}, 라인={line}")
                        
                        # 매칭된 행이 없으면 조금 더 유연한 매칭 시도
                        # 라인 번호만 일치하는 경우 (임시 조치)
                        if not matching_rows:
                            line_matches = [row_idx for row_idx, bldg, ln in all_rows_info if ln == line]
                            if line_matches:
                                print(f"  ⚠️ 라인 번호만 일치: 동={building}, 라인={line}")
                                for row_idx in line_matches:
                                    # 같은 동 건물의 다른 라인이라면 추가
                                    matching_bldg = [bldg for r_idx, bldg, ln in all_rows_info if r_idx == row_idx][0]
                                    print(f"    - 행 {row_idx}: 동={matching_bldg}, 라인={line}")
                        
                        # 최종 매칭 결과 처리
                        if matching_rows:
                            for row_idx in matching_rows:
                                # 각 배관명에 대해 처리
                                processed = False
                                for col_idx, pipe_name in column_to_pipe.items():
                                    if pipe_name in inspected_pipes:
                                        cell = ws_target.cell(row=row_idx, column=col_idx)
                                        cell.value = "완료"
                                        cell.fill = blue_fill
                                        cell.font = black_font
                                        print(f"  • [입상관] 동: {building}, 라인: {line}, 배관: {pipe_name} -> 완료 표시 (행 {row_idx})")
                                        updated_count += 1
                                        total_updated_cells += 1
                                        updated_pipe_types.add(pipe_type)
                                        processed = True
                                        
                                        # 해당 동-라인-배관과 일치하는 파일들을 성공으로 표시
                                        for f_name, f_info in processed_riser_files.items():
                                            if (f_info["building"] == building and 
                                                f_info["line"] == line and 
                                                (f_info["pipe"] == pipe_name or pipe_name in f_info["pipe"])):
                                                # 파일 처리 성공 기록
                                                file_results[f_name] = {
                                                    "status": "성공",
                                                    "pipe_type": pipe_type
                                                }
                                                processed_files.add(f_name)
                                                f_info["status"] = True  # 처리 완료 표시
                                                # 배관종류별 성공 결과 추가
                                                if pipe_type not in pipe_type_results:
                                                    pipe_type_results[pipe_type] = {"files": [], "failed": []}
                                                if f_name not in pipe_type_results[pipe_type]["files"]:
                                                    pipe_type_results[pipe_type]["files"].append(f_name)
                                
                                if processed:
                                    matched_count += 1
                        else:
                            # 매칭된 행이 없으면 수동 확인 필요
                            pipe_names_str = ", ".join(inspected_pipes)
                            print(f"  ⚠️ 동-라인 정보를 시트에서 찾을 수 없음: 동={building}, 라인={line}, 배관={pipe_names_str}")
                            print(f"     ├── 작업 현황 시트를 수동 확인하세요.")
                            print(f"     └── 가능한 원인: 1) 시트에 해당 동-라인이 없음, 2) 시트 구조가 예상과 다름, 3) 동-라인 표기 형식이 다름")
                    
                    # 매칭 결과 요약
                    print(f"  • 입상관 처리 결과: {matched_count}/{len(riser_inspections)} 매칭됨")
                    
                    # 배관명 불일치로 매칭되지 않은 파일들에 대해 오류 상태 표시
                    for f_name, f_info in processed_riser_files.items():
                        if not f_info["status"]:  # 아직 처리되지 않은 파일
                            # 동과 라인은 일치하지만 배관명이 불일치인 경우 확인
                            for (check_building, check_line), _ in riser_inspections.items():
                                if f_info["building"] == check_building and f_info["line"] == check_line:
                                    # 배관명 불일치 오류로 표시
                                    file_results[f_name] = {
                                        "status": "실패",
                                        "pipe_type": pipe_type,
                                        "reason": f"배관명 불일치: {f_info['pipe']}"
                                    }
                                    processed_files.add(f_name)
                                    # 배관종류별 실패 결과 추가
                                    if pipe_type not in pipe_type_results:
                                        pipe_type_results[pipe_type] = {"files": [], "failed": []}
                                    pipe_type_results[pipe_type]["failed"].append({"file": f_name, "reason": f"배관명 불일치: {f_info['pipe']}"})
                                    print(f"  ⚠️ {f_name}: 동/라인은 일치하지만 배관명({f_info['pipe']})이 시트와 일치하지 않음")
                                    break
                else:
                    # 기존 처리 (입상관 아닌 경우)
                    # 세대별 검사된 배관 번호 모음
                    inspected_by_unit = {}  # 키: (동, 호문자열), 값: set(배관번호 문자열들)
                    
                    # 파일별 처리 정보 저장용 변수
                    processed_unit_files = {}  # 파일명: {"building": 동, "unit": 호수, "pipe_num": 배관번호, "processed": 처리여부}
                    
                    print(f"  • {pipe_type} 처리 시작 (비입상관 처리)")
                    
                    for fname in video_file_names:
                        # 파일명에서 정보 추출 (예: "102동 1903호 세대매립관 세탁.mp4")
                        m = re.match(r"(\d+)동\s*(\d+)호\s*(.+?)\s*(\S+)\.mp4", fname)
                        if not m:
                            print(f"  ⚠️ 파일명 형식이 맞지 않음: {fname}")
                            # 파일 처리 실패 기록
                            file_results[fname] = {
                                "status": "실패",
                                "pipe_type": pipe_type,
                                "reason": "파일명 형식이 맞지 않음"
                            }
                            processed_files.add(fname)
                            continue
                        
                        building = int(m.group(1))
                        unit_str = m.group(2)  # 예: "1903"
                        video_pipe_type = m.group(3)  # 배관종류 (예: "세대매립관")
                        pipe_name = m.group(4)  # 배관명 (예: "세탁")
                        
                        # 배관명에서 괄호 이후 부분 제거
                        if '(' in pipe_name:
                            pipe_name = pipe_name.split('(')[0].strip()
                            print(f"  • 괄호 처리: 배관명을 '{pipe_name}'으로 정리")
                        
                        # 현재 처리 중인 배관종류와 일치하는지 확인
                        if pipe_type.strip() not in video_pipe_type:
                            # 배관종류 불일치 로그를 출력하지 않음
                            # 처리된 파일 표시
                            processed_files.add(fname)
                            continue
                            
                        # 배관종류가 일치하는 경우에만 파일 정보 출력
                        print(f"  • 파일 분석: 동={building}, 호={unit_str}, 배관종류={video_pipe_type}, 배관명={pipe_name}")
                        
                        # (building, unit_str) 키가 없으면 초기화
                        if (building, unit_str) not in inspected_by_unit:
                            inspected_by_unit[(building, unit_str)] = set()
                        
                        # 범례 사전으로 배관명을 번호로 변환하여 저장
                        if pipe_name in pipe_legend:
                            pipe_num = pipe_legend[pipe_name]
                            inspected_by_unit[(building, unit_str)].add(pipe_num)
                            updated_pipe_types.add(pipe_type)
                            print(f"  ✓ 추출 성공: 동={building}, 호={unit_str}, 배관={pipe_name}, 번호={pipe_legend[pipe_name]}")
                            
                            # 파일 정보 임시 저장
                            processed_unit_files[fname] = {
                                "building": building,
                                "unit": unit_str,
                                "pipe_num": pipe_num,
                                "processed": False  # 아직 처리 안됨
                            }
                        else:
                            print(f"  ⚠️ 범례에 없는 배관명: {pipe_name}")
                            # 파일 처리 실패 기록
                            file_results[fname] = {
                                "status": "실패",
                                "pipe_type": pipe_type,
                                "reason": f"범례에 없는 배관명: {pipe_name}"
                            }
                            # 배관종류별 실패 결과 추가
                            if pipe_type not in pipe_type_results:
                                pipe_type_results[pipe_type] = {"files": [], "failed": []}
                            pipe_type_results[pipe_type]["failed"].append({"file": fname, "reason": f"범례에 없는 배관명: {pipe_name}"})
                    
                    # 세대 위치에 배관번호 기록
                    building_info = {
                        101: {"start_col": "A", "lines": 4}, 102: {"start_col": "H", "lines": 4},
                        103: {"start_col": "O", "lines": 5}, 104: {"start_col": "W", "lines": 6},
                        105: {"start_col": "AF", "lines": 4}, 106: {"start_col": "AM", "lines": 4},
                        107: {"start_col": "AT", "lines": 4}, 108: {"start_col": "BA", "lines": 4},
                        109: {"start_col": "BH", "lines": 4}, 110: {"start_col": "BO", "lines": 4},
                    }
                    
                    print(f"  • 시트 구조 정보:")
                    for building, info in building_info.items():
                        print(f"    동={building}, 시작열={info['start_col']}, 라인수={info['lines']}")
                    
                    # 처리 결과 추적
                    processed_units = 0
                    
                    from openpyxl.utils import column_index_from_string
                    
                    for (building, unit_str), pipes in inspected_by_unit.items():
                        if building not in building_info:
                            print(f"  ⚠️ {building}동 정보가 없어 처리할 수 없습니다.")
                            continue
                            
                        # 층과 라인 계산
                        try:
                            if len(unit_str) >= 3:
                                floor = int(unit_str[:-2])  # 마지막 두 자리를 제외한 부분 (층)
                                line = int(unit_str[-2:])   # 마지막 두 자리 (라인)
                            else:
                                floor = int(unit_str[0])
                                line = int(unit_str[1:])  # (3자리 호수 처리)
                            
                            row = 41 - floor
                            start_col_index = column_index_from_string(building_info[building]["start_col"])
                            target_col_index = start_col_index + line  # (층 라벨열 + line)
                            
                            # 디버깅 정보
                            print(f"  • 계산 정보: 동={building}, 호={unit_str} -> 층={floor}, 라인={line}")
                            print(f"    -> 행={row}, 열 기준={building_info[building]['start_col']}({start_col_index}), 계산 열={target_col_index}")
                            
                            # 범위 검증
                            if row < 1 or row > ws_target.max_row:
                                print(f"  ⚠️ 계산된 행({row})이 범위를 벗어남: 1~{ws_target.max_row}")
                                continue
                                
                            if target_col_index < 1 or target_col_index > ws_target.max_column:
                                print(f"  ⚠️ 계산된 열({target_col_index})이 범위를 벗어남: 1~{ws_target.max_column}")
                                continue
                            
                            # 해당 열의 마지막행번호-1 행의 셀값 확인
                            last_row = ws_target.max_row
                            if last_row <= 1:
                                print(f"  ⚠️ 시트 행 범위 오류: max_row={last_row}")
                                continue
                                
                            reference_cell = ws_target.cell(row=last_row-1, column=target_col_index)
                            reference_value = reference_cell.value
                            print(f"    → 참조값: {reference_value}")
                            
                            # 셀 값 설정 및 스타일 적용
                            cell = ws_target.cell(row=row, column=target_col_index)
                            current_value = cell.value
                            print(f"    → 현재 셀값: {current_value}")
                            
                            if pipes:
                                # 번호들을 "/"로 연결하여 입력
                                pipe_str = "/".join(sorted(pipes, key=lambda x: int(x)))
                                
                                # 참조 셀 값과 비교
                                if pipe_str == reference_value:
                                    cell.value = "완료"
                                    cell.fill = blue_fill
                                    print(f"    ✓ '완료' 표시 (reference 일치)")
                                else:
                                    cell.value = pipe_str
                                    cell.fill = yellow_fill
                                    print(f"    ✓ '{pipe_str}' 표시 (reference 불일치)")
                                    
                                cell.font = black_font
                                print(f"  • [{pipe_type}] 동: {building}, 호수: {unit_str} -> {pipe_str} 처리")
                                updated_count += 1
                                total_updated_cells += 1
                                processed_units += 1
                                
                                # 해당 동-호에 해당하는 파일들을 모두 성공으로 표시
                                for f_name, f_info in processed_unit_files.items():
                                    if f_info["building"] == building and f_info["unit"] == unit_str:
                                        # 파일 처리 성공 기록
                                        file_results[f_name] = {
                                            "status": "성공",
                                            "pipe_type": pipe_type
                                        }
                                        processed_files.add(f_name)
                                        f_info["processed"] = True  # 처리 완료 표시
                                        
                                        # 배관종류별 성공 결과 추가
                                        if pipe_type not in pipe_type_results:
                                            pipe_type_results[pipe_type] = {"files": [], "failed": []}
                                        if f_name not in pipe_type_results[pipe_type]["files"]:
                                            pipe_type_results[pipe_type]["files"].append(f_name)
                            else:
                                cell.value = None  # 점검된 배관 없으면 비워둠
                                print(f"    ⚠️ 처리할 배관 정보 없음")
                        except Exception as e:
                            print(f"  ❌ 처리 중 오류 발생: {str(e)}")
                            continue
                    
                    # 처리 결과 요약
                    print(f"  • {pipe_type} 처리 결과: {processed_units}/{len(inspected_by_unit)} 처리됨")
            
            # 수정된 내용을 원본 파일로 저장
            output_path = excel_path
            try:
                wb.save(output_path)
                print(f"\n✓ 파일 저장 완료: {os.path.basename(output_path)}")
                
                # 파일별 성공/실패 요약 계산
                success_count = sum(1 for result in file_results.values() if result.get("status") == "성공")
                failed_count = sum(1 for result in file_results.values() if result.get("status") == "실패")
                skipped_count = total_videos - success_count - failed_count
                
                # 작업 결과 요약 표시
                print("\n[ 작업 결과 요약 ]")
                print(f"• 처리한 동영상 파일: {total_videos}개")
                print(f"• 업데이트된 세대/라인 수: {updated_count}개")
                print(f"• 업데이트된 셀 수: {total_updated_cells}개")
                print(f"• 처리된 배관종류: {', '.join(updated_pipe_types) if updated_pipe_types else '없음'}")
                
                # 파일별 성공/실패 요약 추가
                print(f"\n[ 파일별 처리 결과 ]")
                
                # ANSI 색상 코드 제거
                
                # 작업폴더의 모든 동영상 파일을 표시
                for fname in video_file_names:
                    if fname in file_results:
                        status = file_results[fname].get("status", "")
                        reason = file_results[fname].get("reason", "")
                        
                        if status == "성공":
                            print(f"• {fname}: (성공)")
                        elif status == "실패":
                            print(f"• {fname}: (실패) - {reason}")
                        elif status == "건너뜀":
                            if "배관명 불일치" in reason:
                                print(f"• {fname}: (오류) - {reason}")
                            else:
                                print(f"• {fname}: (건너뜀) - {reason}")
                    else:
                        # 결과가 없는 파일은 처리되지 않음으로 표시
                        print(f"• {fname}: (오류) - 처리되지 않음")
                        # 파일 정보 추가
                        file_results[fname] = {
                            "status": "건너뜀",
                            "reason": "처리되지 않음"
                        }
                
                # 추가 안내 메시지
                if updated_count == 0:
                    print("\n⚠️ 세대/라인 정보가 하나도 업데이트되지 않았습니다. 가능한 원인:")
                    print("  1. 동영상 파일명 형식이 '[동] [호] [배관종류] [배관명]' 형식과 다릅니다.")
                    print("  2. 시트 구조가 예상과 다릅니다 (행/열 위치, 형식 등).")
                    print("  3. 동영상에 표시된 배관 종류와 배관명이 시트에 없습니다.")
                
                # 마지막 확인: 모든 파일이 처리됐는지 확인
                for fname in video_file_names:
                    if fname not in file_results and fname not in processed_files:
                        file_results[fname] = {
                            "status": "건너뜀", 
                            "reason": "처리되지 않음"
                        }
                    elif fname not in processed_files and fname in file_results:
                        if file_results[fname].get("status") != "성공" and file_results[fname].get("status") != "실패":
                            file_results[fname]["status"] = "건너뜀"
                            if "reason" not in file_results[fname]:
                                file_results[fname]["reason"] = "처리되지 않음"
                
            except Exception as e:
                print(f"❌ 파일 저장 중 오류: {str(e)}")
                print("다른 프로그램에서 파일을 열고 있는지 확인하세요.")
            
            # COM 객체 정리 및 메모리 정리
            self.cleanup_com_objects()
            self.cleanup_memory()
            
            # 작업 완료 후 한번 더 엑셀 프로세스 종료 확인
            time.sleep(0.5)  # 잠시 대기 후 프로세스 확인
            self.close_excel_file(excel_path)
            
        except Exception as e:
            print(f"\n❌ 예상치 못한 오류 발생: {str(e)}")
            import traceback
            print(traceback.format_exc())  # 디버깅을 위한 스택 트레이스 출력
        finally:
            self.finish_process()

def main():
    """메인 함수: 애플리케이션 시작"""
    # 프로그램 시작 로그
    logger.info("TaskieX 애플리케이션 시작")
    
    # Tkinter 루트 윈도우 생성
    root = tk.Tk()
    
    try:
        # 애플리케이션 인스턴스 생성
        app = TaskieXApp(root)
        
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
        logger.info("TaskieX 애플리케이션 종료")

if __name__ == "__main__":
    main() 