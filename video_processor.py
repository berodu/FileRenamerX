#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import cv2
from datetime import datetime
import shutil
import numpy as np
from PIL import Image
import time
import subprocess
import json

class VideoProcessor:
    def __init__(self, video_dir, output_dir):
        """
        비디오 프로세서 초기화
        
        Args:
            video_dir (str): 비디오 파일이 있는 디렉토리 경로
            output_dir (str): 결과물을 저장할 디렉토리 경로
        """
        self.video_dir = video_dir
        self.output_dir = output_dir
        
        try:
            os.makedirs(output_dir, exist_ok=True)
        except PermissionError:
            raise PermissionError(f"결과 디렉토리 생성 권한이 없습니다: {output_dir}")
        except Exception as e:
            raise Exception(f"결과 디렉토리 생성 중 오류: {str(e)}")
    
    def validate_video_metadata(self, video_path):
        """
        동영상 파일의 메타데이터를 검증
        
        Args:
            video_path (str): 비디오 파일 경로
            
        Returns:
            tuple: (is_valid: bool, error_message: str)
        """
        try:
            # 파일 존재 여부 확인
            if not os.path.exists(video_path):
                return False, "파일을 찾을 수 없습니다"
            
            # 파일 읽기 권한 확인
            if not os.access(video_path, os.R_OK):
                return False, "파일 읽기 권한이 없습니다"
            
            # ffprobe 실행 가능 여부 확인
            try:
                # ffprobe 버전 확인으로 실행 가능 여부 체크
                check_result = subprocess.run(
                    ['ffprobe', '-version'],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=5
                )
                if check_result.returncode != 0:
                    return False, "ffprobe를 실행할 수 없습니다. FFmpeg가 설치되어 있는지 확인하세요."
            except FileNotFoundError:
                return False, "ffprobe를 찾을 수 없습니다. FFmpeg가 설치되어 있고 PATH에 등록되어 있는지 확인하세요."
            
            # ffprobe를 사용하여 메타데이터 조회
            # Windows 경로에 공백이나 특수문자가 있을 수 있으므로 따옴표로 감싸지 않고 직접 전달
            cmd = [
                'ffprobe',
                '-v', 'error',
                '-print_format', 'json',
                '-show_format',
                '-show_streams',
                video_path
            ]
            
            result = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=30,
                encoding='utf-8',
                errors='ignore'  # 인코딩 오류 무시
            )
            
            # stderr에 오류 메시지가 있는지 확인
            if result.stderr and result.stderr.strip():
                # 실제 오류 메시지 추출
                error_lines = [line.strip() for line in result.stderr.split('\n') 
                              if line.strip()]
                if error_lines:
                    # 첫 번째 의미있는 오류 메시지 반환
                    for line in error_lines:
                        if any(keyword in line.lower() for keyword in ['error', 'invalid', 'cannot', 'unable', 'failed']):
                            return False, f"ffprobe 오류: {line}"
                    # 오류 키워드가 없어도 stderr가 있으면 반환
                    return False, f"ffprobe 경고/오류: {error_lines[0]}"
            
            if result.returncode != 0:
                error_msg = result.stderr.strip() if result.stderr and result.stderr.strip() else "알 수 없는 오류"
                return False, f"메타데이터 조회 실패 (반환 코드: {result.returncode}): {error_msg}"
            
            # stdout이 비어있는지 확인
            if not result.stdout or not result.stdout.strip():
                # stderr에 더 자세한 정보가 있을 수 있음
                if result.stderr and result.stderr.strip():
                    return False, f"메타데이터 출력이 없습니다: {result.stderr.strip()}"
                else:
                    return False, "메타데이터 출력이 없습니다 (파일이 손상되었거나 지원하지 않는 형식일 수 있습니다)"
            
            # JSON 파싱
            try:
                metadata = json.loads(result.stdout)
            except json.JSONDecodeError as e:
                return False, f"메타데이터 파싱 실패: {str(e)}"
            except TypeError as e:
                return False, f"메타데이터 파싱 실패: {str(e)}"
            
            # 인코더 확인 (format 섹션에서)
            encoder = None
            if 'format' in metadata and 'tags' in metadata['format']:
                encoder = metadata['format']['tags'].get('encoder', '')
            
            # 코덱 확인 (streams 섹션에서)
            codec = None
            if 'streams' in metadata and len(metadata['streams']) > 0:
                # 비디오 스트림 찾기
                for stream in metadata['streams']:
                    if stream.get('codec_type') == 'video':
                        codec = stream.get('codec_name', '')
                        break
            
            # 검증: 인코더가 Lavf56.25.101인지 확인
            if not encoder or 'Lavf56.25.101' not in encoder:
                return False, f"인코더가 올바르지 않습니다. (현재: {encoder}, 요구: Lavf56.25.101)"
            
            # 검증: 코덱이 hevc인지 확인
            if not codec or codec.lower() not in ['hevc', 'h265', 'h.265']:
                return False, f"코덱이 올바르지 않습니다. (현재: {codec}, 요구: HEVC/H.265)"
            
            return True, ""
            
        except FileNotFoundError:
            return False, "ffprobe를 찾을 수 없습니다. FFmpeg가 설치되어 있는지 확인하세요."
        except subprocess.TimeoutExpired:
            return False, "메타데이터 조회 시간 초과"
        except Exception as e:
            return False, f"메타데이터 검증 중 오류: {str(e)}"
    
    def extract_frames(self, video_file, frame_times):
        """
        비디오에서 지정된 시간의 프레임을 추출
        
        Args:
            video_file (str): 비디오 파일명
            frame_times (list): 추출할 프레임 시간(초)의 리스트
        
        Returns:
            list: 추출된 프레임 이미지 파일 경로 리스트
        """
        if not frame_times:
            return []
            
        video_path = os.path.join(self.video_dir, video_file)
        
        # 비디오 파일 존재 확인
        if not os.path.exists(video_path):
            print(f"  ❌ 비디오 파일이 존재하지 않습니다: {video_path}")
            return []
            
        # 파일 크기 확인 (0바이트 파일인지)
        try:
            file_size = os.path.getsize(video_path)
            if file_size == 0:
                print(f"  ❌ 빈 비디오 파일입니다: {video_path}")
                return []
        except Exception as e:
            print(f"  ❌ 파일 크기 확인 중 오류: {str(e)}")
            return []
        
        # 비디오 파일 열기 시도 (최대 3번)
        video = None
        max_attempts = 3
        for attempt in range(max_attempts):
            try:
                video = cv2.VideoCapture(video_path)
                if video.isOpened():
                    break
                else:
                    print(f"  ⚠️ 비디오 파일 열기 실패 (시도 {attempt+1}/{max_attempts})")
                    time.sleep(0.5)  # 재시도 전 잠시 대기
            except Exception as e:
                print(f"  ❌ 비디오 파일 열기 중 오류: {str(e)}")
                if attempt < max_attempts - 1:
                    time.sleep(0.5)  # 재시도 전 잠시 대기
        
        # 모든 시도 후에도 열지 못한 경우
        if video is None or not video.isOpened():
            print(f"  ❌ {video_path} 파일을 열 수 없습니다.")
            return []
        
        try:
            # 비디오 정보 가져오기
            fps = video.get(cv2.CAP_PROP_FPS)
            frame_count = int(video.get(cv2.CAP_PROP_FRAME_COUNT))
            
            # 유효한 FPS 및 프레임 수 확인
            if fps <= 0 or frame_count <= 0:
                print(f"  ❌ 비디오 정보가 유효하지 않습니다: FPS={fps}, 프레임 수={frame_count}")
                video.release()
                return []
                
            duration = frame_count / fps
            
            frame_paths = []
            
            # 비디오 파일명에서 날짜 정보 추출 (예: 20250101_123030.mp4 -> 20250101_123030)
            base_name = os.path.splitext(video_file)[0]
            
            # 각 지정된 시간에 대해 프레임 추출
            for time_sec in frame_times:
                if time_sec > duration:
                    print(f"  ⚠️ {time_sec}초는 비디오 길이({duration:.2f}초)보다 깁니다. 이 프레임은 건너뜁니다.")
                    continue
                    
                frame_number = int(time_sec * fps)
                
                # 유효한 프레임 번호인지 확인
                if frame_number >= frame_count:
                    print(f"  ⚠️ 프레임 번호({frame_number})가 총 프레임 수({frame_count})보다 큽니다. 이 프레임은 건너뜁니다.")
                    continue
                
                # 비디오 위치 설정 재시도 (최대 3회)
                seek_success = False
                for _ in range(3):
                    try:
                        if video.set(cv2.CAP_PROP_POS_FRAMES, frame_number):
                            seek_success = True
                            break
                        else:
                            time.sleep(0.1)  # 잠시 대기 후 재시도
                    except Exception:
                        time.sleep(0.1)  # 오류 시 잠시 대기
                
                if not seek_success:
                    print(f"  ❌ {time_sec}초 위치로 이동할 수 없습니다.")
                    continue
                
                # 프레임 읽기 시도 (최대 3회)
                success = False
                frame = None
                for _ in range(3):
                    try:
                        success, frame = video.read()
                        if success and frame is not None:
                            break
                        else:
                            time.sleep(0.1)  # 잠시 대기 후 재시도
                    except Exception:
                        time.sleep(0.1)  # 오류 시 잠시 대기
                
                if success and frame is not None:
                    # 빈 프레임인지 확인 (흑백 또는 완전히 비어있는 프레임)
                    if frame.size == 0 or np.mean(frame) < 5:  # 평균 픽셀 값이 5 미만이면 거의 검은색
                        print(f"  ⚠️ {time_sec}초 위치의 프레임이 비어 있거나 검은색입니다.")
                        continue
                    
                    # 이미지 파일명 생성 (원본_시간초.jpg)
                    frame_filename = f"{base_name}_{time_sec}sec.jpg"
                    frame_path = os.path.abspath(os.path.join(self.output_dir, frame_filename))
                    
                    # 파일 저장 시도 (최대 3회)
                    saved = False
                    for attempt in range(3):
                        try:
                            # PIL을 사용하여 이미지 저장
                            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                            img = Image.fromarray(frame_rgb)
                            img.save(frame_path)
                            
                            # 파일이 실제로 생성되었는지 확인
                            if os.path.exists(frame_path) and os.path.getsize(frame_path) > 0:
                                frame_paths.append(frame_path)
                                saved = True
                                break
                            else:
                                print(f"  ⚠️ 이미지 저장 실패 (시도 {attempt+1}/3)")
                                time.sleep(0.3)  # 잠시 대기 후 재시도
                        except PermissionError:
                            print(f"  ❌ 이미지 저장 권한이 없습니다: {frame_path}")
                            break  # 권한 오류는 재시도해도 해결되지 않으므로 중단
                        except Exception as e:
                            print(f"  ❌ 이미지 저장 중 오류: {str(e)}")
                            if attempt < 2:  # 마지막 시도가 아니면 재시도
                                time.sleep(0.3)
                    
                    if not saved:
                        print(f"  ❌ {time_sec}초 프레임 저장에 모든 시도가 실패했습니다.")
                else:
                    print(f"  ❌ {video_path}에서 {time_sec}초 프레임을 추출할 수 없습니다.")
            
            # 비디오 파일 닫기
            video.release()
            
            return frame_paths
            
        except Exception as e:
            print(f"  ❌ 프레임 추출 중 오류: {str(e)}")
            if video is not None:
                video.release()
            return []
    
    def rename_image(self, image_path, new_name):
        """
        이미지 파일의 이름을 변경
        
        Args:
            image_path (str): 현재 이미지 파일 경로
            new_name (str): 새로운 파일명 (확장자 제외)
        
        Returns:
            str: 새로운 파일 경로
        """
        try:
            # 파일 존재 확인
            if not os.path.exists(image_path):
                print(f"  ❌ {image_path} 파일이 존재하지 않습니다.")
                return None
            
            # 입력 유효성 검사
            if not new_name or len(new_name.strip()) == 0:
                print(f"  ❌ 새 파일명이 비어있습니다.")
                return None
                
            # 파일명에 유효하지 않은 문자가 있는지 확인
            invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
            if any(char in new_name for char in invalid_chars):
                print(f"  ❌ 새 파일명에 유효하지 않은 문자가 포함되어 있습니다: {new_name}")
                return None
            
            # 원본 파일명과 확장자 분리
            dir_path = os.path.dirname(image_path)
            _, ext = os.path.splitext(image_path)
            
            # 새 파일 경로 생성
            new_path = os.path.join(dir_path, f"{new_name}{ext}")
            
            # 대상 파일이 이미 존재하는지 확인
            if os.path.exists(new_path) and os.path.abspath(image_path) != os.path.abspath(new_path):
                print(f"  ⚠️ 대상 파일이 이미 존재합니다: {new_path}")
                # 기존 파일 백업 또는 삭제 등의 처리 가능
                
                # 기존 파일 삭제 예시
                try:
                    os.remove(new_path)
                except Exception as e:
                    print(f"  ❌ 기존 파일 삭제 실패: {str(e)}")
                    return None
            
            # 파일명 변경 (최대 3회 시도)
            for attempt in range(3):
                try:
                    os.rename(image_path, new_path)
                    return new_path
                except PermissionError:
                    print(f"  ❌ 파일 이름 변경 권한이 없습니다: {image_path}")
                    return None  # 권한 오류는 재시도해도 해결되지 않으므로 중단
                except FileNotFoundError:
                    print(f"  ❌ 원본 파일을 찾을 수 없습니다: {image_path}")
                    return None  # 파일이 없으면 재시도해도 해결되지 않음
                except Exception as e:
                    if attempt < 2:  # 마지막 시도가 아니면 재시도
                        print(f"  ⚠️ 파일 이름 변경 실패 (시도 {attempt+1}/3): {str(e)}")
                        time.sleep(0.3)  # 잠시 대기 후 재시도
                    else:
                        print(f"  ❌ 파일 이름 변경에 모든 시도가 실패했습니다: {str(e)}")
                        return None
            
            return None
            
        except Exception as e:
            print(f"  ❌ 파일 이름 변경 중 예상치 못한 오류: {str(e)}")
            return None 