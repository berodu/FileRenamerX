#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import argparse
import sys
import re
import shutil
import time
from collections import Counter
from video_processor import VideoProcessor
from image_analyzer import GoogleVisionAnalyzer, ChatGPTVisionAnalyzer, GeminiAnalyzer

def is_valid_format(extracted_info):
    """추출 정보가 '[~동] [~호] [배관종류] [배관명]' 형식인지 확인"""
    pattern = r'\[\S+동\] \[\S+호\] \[\S+\] \[\S+\]'
    return bool(re.match(pattern, extracted_info))

def has_zero_prefix(extracted_info):
    """동, 호 정보가 0으로 시작하는지 확인"""
    # [0동], [01동], [02호] 등과 같은 패턴 찾기
    pattern = r'\[0\d*동\]|\[\d*0\d*호\]'
    return bool(re.search(pattern, extracted_info))

def has_no_spaces(extracted_info):
    """텍스트에 띄어쓰기가 없는지 확인 (대괄호 사이의 공백 제외)"""
    # 대괄호 사이의 공백을 임시로 다른 문자로 치환
    temp_text = re.sub(r'\] \[', ']#[', extracted_info)
    # 남은 공백이 있는지 확인
    return ' ' not in temp_text

def clear_directory(directory):
    """디렉토리 내용을 모두 삭제"""
    if os.path.exists(directory):
        try:
            for item in os.listdir(directory):
                item_path = os.path.join(directory, item)
                try:
                    if os.path.isfile(item_path):
                        os.unlink(item_path)
                    elif os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                except PermissionError:
                    print(f"  ❌ 권한 오류: {item_path} 접근이 거부되었습니다.")
                except FileNotFoundError:
                    # 이미 삭제된 파일인 경우 무시
                    pass
                except Exception as e:
                    print(f"  ❌ 파일 삭제 실패: {item_path}")
            return True
        except PermissionError:
            print(f"  ❌ 폴더 접근 권한이 없습니다: {directory}")
        except Exception as e:
            print(f"  ❌ 폴더 초기화 중 오류 발생: {str(e)}")
    return False

def validate_frame_times(frame_times_str):
    """프레임 시간 문자열 검증 및 파싱"""
    try:
        # 쉼표로 구분된 값을 숫자로 변환
        values = [v.strip() for v in frame_times_str.split(',')]
        frame_times = []
        
        for v in values:
            try:
                # 음수 검사
                time_val = float(v)
                if time_val < 0:
                    print(f"  ❌ 유효하지 않은 시간: {v} (음수)")
                    continue
                
                # 너무 큰 값 검사 (예: 10시간 이상)
                if time_val > 36000:  # 10시간(초)
                    print(f"  ❌ 유효하지 않은 시간: {v} (너무 큰 값)")
                    continue
                
                frame_times.append(int(time_val))
            except ValueError:
                print(f"  ❌ 유효하지 않은 시간 형식: {v}")
        
        if not frame_times:
            print("  ⚠️ 유효한 시간이 없습니다. 기본값 2,3,5초를 사용합니다.")
            return [2, 3, 5]  # 기본값
        
        return sorted(frame_times)  # 정렬된 값 반환
    except Exception as e:
        print(f"  ❌ 시간 파싱 오류: {str(e)}, 기본값 2,3,5초를 사용합니다.")
        return [2, 3, 5]  # 오류 시 기본값

def select_best_result(results):
    """우선순위에 따라 최상의 결과 선택"""
    if not results:
        return None
    
    # 유효한 형식인지 검사
    valid_format_results = [r for r in results if is_valid_format(r["extracted_info"])]
    
    # 유효한 형식의 결과가 없으면 원래 결과 사용
    if not valid_format_results:
        # 원래 결과도 없으면 None 반환
        if not results:
            return None
        # 모든 결과가 유효하지 않은 형식이면 원래 결과 중 첫 번째 사용
        return results[0]
    
    # 이후 분석은 유효한 형식의 결과만 대상으로 함
    results = valid_format_results
    
    # 1. 정보가 있는 것 (이미 results에 포함된 항목은 모두 정보가 있음)
    
    # 2. 0으로 시작하는 동, 호 정보가 있는 결과 제외
    valid_results = [r for r in results if not has_zero_prefix(r["extracted_info"])]
    
    # 유효한 결과가 없으면 원래 결과 사용
    if not valid_results:
        valid_results = results
    
    # 3. 다수결 - 가장 많이 나온 결과를 선택
    # 결과 텍스트별 빈도수 계산
    result_counter = Counter([r["extracted_info"] for r in valid_results])
    
    # 가장 많이 나온 결과 찾기
    most_common_results = result_counter.most_common()
    
    # 가장 많이 나온 결과에 해당하는 항목들
    most_common_items = [r for r in valid_results if r["extracted_info"] == most_common_results[0][0]]
    
    # 4. 띄어쓰기가 없는 결과 우선
    no_space_results = [r for r in most_common_items if has_no_spaces(r["extracted_info"])]
    
    # 띄어쓰기가 없는 결과가 있으면 해당 결과 중 첫 번째 것 선택
    if no_space_results:
        return no_space_results[0]
    
    # 없으면 가장 많이 나온 결과 중 첫 번째 것 선택
    return most_common_items[0]

def is_valid_video_file(file_path):
    """비디오 파일 유효성 검사 (확장자만으로 판단)"""
    # 확장자 확인
    _, ext = os.path.splitext(file_path)
    valid_exts = ['.mp4', '.avi', '.mov', '.wmv', '.mkv', '.flv']
    return ext.lower() in valid_exts

def check_api_keys():
    """API 키 파일 존재 여부 확인"""
    missing_keys = []
    
    # Google Vision API 키 확인
    vision_key_path = 'vision-api-key/vision-ocr-454121-572fb601794b.json'
    if not os.path.exists(vision_key_path):
        missing_keys.append("Google Vision API")
    
    return missing_keys

def main():
    try:
        parser = argparse.ArgumentParser(description='동영상에서 프레임을 추출하고 이미지를 분석합니다.')
        parser.add_argument('--method', type=str, choices=['google', 'chatgpt', 'gemini'], default='google',
                        help='이미지 분석 방법 (google: Vision API + ChatGPT, chatgpt: ChatGPT Vision, gemini: Gemini)')
        parser.add_argument('--frames', type=str, default='2,3,5',
                        help='캡처할 프레임 시간(초)을 쉼표로 구분하여 입력 (기본값: 2,3,5)')
        parser.add_argument('--video_dir', type=str, default='./작업폴더',
                        help='동영상 파일이 있는 디렉토리 (기본값: ./작업폴더)')
        parser.add_argument('--output_dir', type=str, default='./결과',
                        help='결과물을 저장할 디렉토리 (기본값: ./결과)')
        parser.add_argument('--debug', action='store_true', help='디버그 정보 출력')
        parser.add_argument('--verbose', action='store_true', help='상세 로그 출력')
        parser.add_argument('--retry', type=int, default=3, help='API 호출 실패 시 재시도 횟수 (기본값: 3)')
        args = parser.parse_args()

        # API 키 확인
        missing_keys = check_api_keys()
        if missing_keys:
            print(f"⚠️ 주의: 다음 API 키 파일이 없습니다: {', '.join(missing_keys)}")
            if args.method == 'google' and "Google Vision API" in missing_keys:
                print(f"❌ Google Vision API 키가 필요합니다. 분석 방법을 변경하거나 API 키를 추가해주세요.")
                return
            elif args.method == 'chatgpt' and "ChatGPT API" in missing_keys:
                print(f"❌ ChatGPT API 키가 필요합니다. 분석 방법을 변경하거나 API 키를 추가해주세요.")
                return
            elif args.method == 'gemini' and "Gemini API" in missing_keys:
                print(f"❌ Gemini API 키가 필요합니다. 분석 방법을 변경하거나 API 키를 추가해주세요.")
                return

        # 디버그 모드에서만 상세 정보 출력
        if args.debug:
            print(f"Python 버전: {sys.version}")
            print(f"운영체제: {sys.platform}")
            print(f"작업 디렉토리: {os.getcwd()}")

        # 비디오 디렉토리 확인
        video_dir_abs = os.path.abspath(args.video_dir)
        if not os.path.exists(video_dir_abs):
            print(f"⚠️ 비디오 디렉토리가 존재하지 않습니다.")
            try:
                os.makedirs(video_dir_abs, exist_ok=True)
                print(f"✓ 비디오 디렉토리를 생성했습니다.")
            except PermissionError:
                print(f"❌ 비디오 디렉토리 생성 권한이 없습니다.")
                return
            except Exception as e:
                print(f"❌ 비디오 디렉토리 생성 중 오류: {str(e)}")
                return

        # 결과 디렉토리 생성
        output_dir_abs = os.path.abspath(args.output_dir)
        
        # 결과 폴더 내용 삭제
        if os.path.exists(output_dir_abs):
            print(f"\n✓ 결과 폴더 초기화")
            clear_directory(output_dir_abs)
        
        # 결과 디렉토리가 없으면 생성
        try:
            os.makedirs(output_dir_abs, exist_ok=True)
        except PermissionError:
            print(f"❌ 결과 디렉토리 생성 권한이 없습니다: {output_dir_abs}")
            return
        except Exception as e:
            print(f"❌ 결과 디렉토리 생성 중 오류: {str(e)}")
            return
        
        # 간략한 설정 정보 출력
        print("\n✓ 분석 설정")
        print(f"• 방법: {args.method}")
        
        # 프레임 시간 검증 및 변환
        frame_times = validate_frame_times(args.frames)
        if not frame_times:
            return
        
        print(f"• 시간: {','.join(map(str, frame_times))}초")
        
        try:
            # 비디오 프로세서 초기화
            video_processor = VideoProcessor(video_dir_abs, output_dir_abs)
            
            # 분석 방법에 따라 적절한 분석기 선택
            try:
                if args.method == 'google':
                    analyzer = GoogleVisionAnalyzer()
                elif args.method == 'chatgpt':
                    analyzer = ChatGPTVisionAnalyzer()
                else:  # gemini
                    analyzer = GeminiAnalyzer()
            except Exception as e:
                print(f"❌ 분석기 초기화 중 오류: {str(e)}")
                return
            
            # 모든 비디오 파일 처리
            try:
                all_files = os.listdir(video_dir_abs)
                video_files = [f for f in all_files if f.lower().endswith(('.mp4', '.avi', '.mkv', '.mov'))]
                
                if not video_files:
                    print(f"\n❌ 비디오 파일이 없습니다.")
                    return
                
                # 유효한 비디오 파일만 필터링
                valid_video_files = []
                for video_file in video_files:
                    video_path = os.path.join(video_dir_abs, video_file)
                    if is_valid_video_file(video_path):
                        valid_video_files.append(video_file)
                    else:
                        print(f"⚠️ 유효하지 않은 비디오 파일 제외: {video_file}")
                
                if not valid_video_files:
                    print(f"\n❌ 유효한 비디오 파일이 없습니다.")
                    return
                    
                # 동영상별 최상의 결과를 저장할 리스트
                best_results = []
                
                for idx, video_file in enumerate(valid_video_files, 1):
                    print(f"\n[{idx}/{len(valid_video_files)}] {video_file}")
                    
                    # 비디오 처리 중 오류 발생 시 다음 비디오로 계속 진행
                    try:
                        # 비디오에서 프레임 추출
                        frame_paths = video_processor.extract_frames(video_file, frame_times)
                        
                        if not frame_paths:
                            print(f"  ❌ 프레임 추출 실패")
                            continue
                        
                        # 현재 비디오의 결과 저장
                        video_results = []
                        
                        # 각 프레임 분석 시 재시도 로직 추가
                        for frame_path in frame_paths:
                            # 파일 존재 확인
                            if not os.path.exists(frame_path):
                                continue
                            
                            # 분석 시도 (재시도 로직 포함)
                            retry_count = 0
                            extracted_info = None
                            
                            while retry_count < args.retry and extracted_info is None:
                                if retry_count > 0:
                                    print(f"  ⚠️ 재시도 중... ({retry_count}/{args.retry})")
                                    time.sleep(2)  # API 호출 간 딜레이
                                
                                try:
                                    extracted_info = analyzer.analyze_image(frame_path)
                                except Exception as e:
                                    if args.debug:
                                        print(f"  ❌ 분석 오류: {str(e)}")
                                    retry_count += 1
                                    continue
                                
                                # 유효한 형식인지 확인
                                if extracted_info and not is_valid_format(extracted_info):
                                    if args.debug:
                                        print(f"  ⚠️ 유효하지 않은 형식: {extracted_info}")
                                    # 재시도 횟수는 증가시키지만 None으로 설정하지는 않음
                                
                                retry_count += 1
                            
                            if extracted_info:
                                # 결과 저장
                                result = {
                                    "video_file": video_file,
                                    "image_path": frame_path,
                                    "extracted_info": extracted_info,
                                    "suggested_filename": extracted_info
                                }
                                video_results.append(result)
                        
                        # 현재 비디오의 최상의 결과 선택
                        if video_results:
                            try:
                                best_result = select_best_result(video_results)
                                if best_result:
                                    best_results.append(best_result)
                                    print(f"  ✓ 결과: {best_result['extracted_info']}")
                                else:
                                    print(f"  ⚠️ 최적 결과 선택 실패")
                            except Exception as e:
                                print(f"  ❌ 결과 선택 중 오류: {str(e)}")
                        else:
                            print(f"  ❌ 유효한 결과 없음")
                    
                    except Exception as e:
                        print(f"  ❌ 비디오 처리 중 오류: {str(e)}")
                        continue
                
                # 최종 결과 요약
                print("\n[ 최종 결과 ]")
                if best_results:
                    for i, result in enumerate(best_results, 1):
                        print(f"{i}. {result['video_file']} → {result['extracted_info']}")
                else:
                    print("❌ 선택된 결과가 없습니다.")
                    
            except PermissionError:
                print(f"❌ 비디오 디렉토리 접근 권한이 없습니다: {video_dir_abs}")
            except FileNotFoundError:
                print(f"❌ 비디오 디렉토리를 찾을 수 없습니다: {video_dir_abs}")
            except Exception as e:
                print(f"❌ 비디오 파일 리스트 조회 중 오류: {str(e)}")
        
        except Exception as e:
            print(f"❌ 처리 중 예상치 못한 오류 발생: {str(e)}")
    
    except KeyboardInterrupt:
        print("\n\n프로그램이 사용자에 의해 중단되었습니다.")
    except Exception as e:
        print(f"\n❌ 심각한 오류 발생: {str(e)}")

if __name__ == "__main__":
    main() 