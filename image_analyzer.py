#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import json
import base64
import requests
import time
import random
from abc import ABC, abstractmethod
from google.cloud import vision
from google.oauth2 import service_account
from openai import OpenAI
import google.generativeai as genai

class ImageAnalyzer(ABC):
    """이미지 분석기 추상 클래스"""
    
    def __init__(self):
        # prompt.txt 파일 로드
        try:
            with open('prompt.txt', 'r', encoding='utf-8') as f:
                self.prompt = f.read().strip()
                
            # 프롬프트가 비어있는지 확인
            if not self.prompt:
                raise ValueError("프롬프트 파일이 비어 있습니다.")
        except FileNotFoundError:
            raise FileNotFoundError("prompt.txt 파일을 찾을 수 없습니다. 프로젝트 루트 디렉토리에 파일이 있는지 확인하세요.")
        except UnicodeDecodeError:
            raise UnicodeDecodeError("prompt.txt 파일을 읽는 중 인코딩 오류가 발생했습니다. UTF-8 형식인지 확인하세요.")
        except Exception as e:
            raise Exception(f"프롬프트 파일 로드 중 오류 발생: {str(e)}")
    
    @abstractmethod
    def analyze_image(self, image_path):
        """
        이미지를 분석하여 추출된 정보 반환
        
        Args:
            image_path (str): 이미지 파일 경로
        
        Returns:
            str: 이미지에서 추출된 정보
        """
        pass
    
    def check_file_exists(self, image_path):
        """
        파일 존재 여부 확인 및 절대 경로 반환
        
        Args:
            image_path (str): 확인할 이미지 경로
            
        Returns:
            str or None: 절대 경로 또는 None (파일이 없는 경우)
        """
        if not image_path:
            return None
            
        try:
            # 절대 경로로 변환
            abs_path = os.path.abspath(image_path)
            
            # 파일 존재 확인
            if not os.path.isfile(abs_path):
                return None
                
            # 유효한 이미지 파일인지 확인 (크기 확인)
            file_size = os.path.getsize(abs_path)
            if file_size == 0:
                return None
                
            # 파일 확장자 확인
            _, ext = os.path.splitext(abs_path)
            valid_exts = ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp']
            if ext.lower() not in valid_exts:
                return None
                
            return abs_path
        except Exception:
            return None


class GoogleVisionAnalyzer(ImageAnalyzer):
    """Google Vision API와 ChatGPT API를 사용한 이미지 분석기"""
    
    def __init__(self):
        super().__init__()
        
        # Google Vision API 클라이언트 초기화
        try:
            credentials_path = 'vision-api-key/vision-ocr-454121-572fb601794b.json'
            
            # 키 파일 존재 확인
            if not os.path.exists(credentials_path):
                raise FileNotFoundError(f"Google Vision API 키 파일을 찾을 수 없습니다: {credentials_path}")
                
            credentials = service_account.Credentials.from_service_account_file(credentials_path)
            self.vision_client = vision.ImageAnnotatorClient(credentials=credentials)
        except FileNotFoundError as e:
            raise e
        except Exception as e:
            raise Exception(f"Google Vision API 초기화 중 오류: {str(e)}")
        
        # ChatGPT API 클라이언트 초기화
        try:
            api_key_path = 'chatgpt-api-key/chatgpt_api_key.txt'
            
            # 키 파일 존재 확인
            if not os.path.exists(api_key_path):
                raise FileNotFoundError(f"ChatGPT API 키 파일을 찾을 수 없습니다: {api_key_path}")
                
            with open(api_key_path, 'r') as f:
                api_key = f.read().strip()
                
            # API 키가 비어있는지 확인
            if not api_key:
                raise ValueError("ChatGPT API 키가 비어 있습니다.")
                
            self.openai_client = OpenAI(api_key=api_key)
        except FileNotFoundError as e:
            raise e
        except ValueError as e:
            raise e
        except Exception as e:
            raise Exception(f"ChatGPT API 초기화 중 오류: {str(e)}")
    
    def analyze_image(self, image_path):
        """
        Google Vision API로 OCR 수행 후 ChatGPT로 결과 분석
        
        Args:
            image_path (str): 이미지 파일 경로
        
        Returns:
            str: 이미지에서 추출된 정보
        """
        max_retries = 5  # 최대 재시도 횟수 증가 (3 -> 5)
        retry_count = 0
        base_delay = 2  # 기본 대기 시간
        
        while retry_count < max_retries:
            try:
                # 파일 존재 확인
                abs_path = self.check_file_exists(image_path)
                if not abs_path:
                    print(f"❌ [Vision API 오류] 유효하지 않은 이미지 파일: {image_path}")
                    return None
                    
                # 이미지 로드
                try:
                    with open(abs_path, 'rb') as image_file:
                        content = image_file.read()
                except PermissionError as e:
                    print(f"❌ [Vision API 오류] 파일 접근 권한이 없습니다: {str(e)}")
                    return None
                except Exception as e:
                    print(f"❌ [Vision API 오류] 이미지 로드 중 오류: {str(e)}")
                    return None
                
                # 지수 백오프 적용
                if retry_count > 0:
                    # 2^n 공식 적용 (2, 4, 8, 16, 32초)
                    current_delay = base_delay * (2 ** (retry_count - 1))
                    # 약간의 랜덤성 추가 (지터)
                    jitter = random.uniform(0.8, 1.2)
                    delay_with_jitter = current_delay * jitter
                    
                    print(f"  ⚠️ [Vision API 재시도] {retry_count}/{max_retries} (대기: {delay_with_jitter:.2f}초)")
                    time.sleep(delay_with_jitter)
                
                # Vision API로 OCR 수행
                try:
                    print(f"  ✓ Vision API OCR 요청 중...")
                    image = vision.Image(content=content)
                    response = self.vision_client.text_detection(image=image)
                    
                    if response.error.message:
                        # 네트워크 관련 오류는 재시도
                        if any(err in response.error.message.lower() for err in ['network', 'timeout', 'connection']):
                            retry_count += 1
                            print(f"❌ [Vision API 오류] 네트워크 오류: {response.error.message}")
                            continue
                        else:
                            print(f"❌ [Vision API 오류] OCR 실패: {response.error.message}")
                            return None
                except Exception as e:
                    retry_count += 1
                    print(f"❌ [Vision API 오류] API 호출 중 오류: {str(e)}")
                    continue
                
                # OCR 결과 추출
                texts = response.text_annotations
                if not texts:
                    print(f"❌ [Vision API 오류] 텍스트가 감지되지 않았습니다.")
                    return None
                
                # 전체 텍스트 추출 (첫 번째 항목은 전체 텍스트)
                detected_text = texts[0].description
                
                # 텍스트가 비어있는지 확인
                if not detected_text.strip():
                    print(f"❌ [Vision API 오류] 추출된 텍스트가 비어있습니다.")
                    return None
                
                print(f"  ✓ OCR 텍스트 추출 완료 ({len(detected_text)} 자)")
                
                # ChatGPT API를 사용하여 텍스트 분석
                system_msg = f"당신은 이미지에서 추출된 텍스트를 분석하여 필요한 정보를 추출하는 전문가입니다. 다음 지시에 따라 텍스트를 분석해주세요: {self.prompt}"
                
                try:
                    print(f"  ✓ ChatGPT 분석 요청 중...")
                    chat_response = self.openai_client.chat.completions.create(
                        model="gpt-4o",
                        messages=[
                            {"role": "system", "content": system_msg},
                            {"role": "user", "content": detected_text}
                        ],
                        temperature=0.3,
                        max_tokens=100
                    )
                    
                    if chat_response and hasattr(chat_response, 'choices') and len(chat_response.choices) > 0:
                        extracted_info = chat_response.choices[0].message.content.strip()
                        print(f"  ✓ [ChatGPT 분석 성공] 결과를 받았습니다.")
                        return extracted_info
                    else:
                        # ChatGPT API 응답 오류 재시도
                        retry_count += 1
                        print(f"❌ [ChatGPT 오류] 유효하지 않은 응답 형식입니다.")
                        continue
                except Exception as e:
                    # API 호출 오류 재시도
                    retry_count += 1
                    print(f"❌ [ChatGPT 오류] API 호출 중 오류: {str(e)}")
                    
                    # 429 에러 감지 (Too Many Requests)
                    if "429" in str(e) or "rate limit" in str(e).lower() or "too many requests" in str(e).lower():
                        print(f"⚠️ [ChatGPT 오류] 요청 한도 초과 (Rate Limit) - 지수 백오프 적용 중...")
                    continue
                
            except Exception as e:
                retry_count += 1
                print(f"❌ [Vision+ChatGPT 오류] 예상치 못한 오류 발생: {str(e)}")
                continue
        
        print(f"❌ [Vision+ChatGPT 오류] 최대 재시도 횟수({max_retries}회)를 초과했습니다. 분석에 실패했습니다.")
        return None


class ChatGPTVisionAnalyzer(ImageAnalyzer):
    """ChatGPT Vision API를 사용한 이미지 분석기"""
    
    def __init__(self):
        super().__init__()
        
        # ChatGPT API 클라이언트 초기화
        try:
            api_key_path = 'chatgpt-api-key/chatgpt_api_key.txt'
            
            # 키 파일 존재 확인
            if not os.path.exists(api_key_path):
                raise FileNotFoundError(f"ChatGPT API 키 파일을 찾을 수 없습니다: {api_key_path}")
                
            with open(api_key_path, 'r') as f:
                api_key = f.read().strip()
                
            # API 키가 비어있는지 확인
            if not api_key:
                raise ValueError("ChatGPT API 키가 비어 있습니다.")
                
            self.openai_client = OpenAI(api_key=api_key)
        except FileNotFoundError as e:
            raise e
        except ValueError as e:
            raise e
        except Exception as e:
            raise Exception(f"ChatGPT API 초기화 중 오류: {str(e)}")
    
    def analyze_image(self, image_path):
        """
        ChatGPT Vision API로 이미지 직접 분석
        
        Args:
            image_path (str): 이미지 파일 경로
        
        Returns:
            str: 이미지에서 추출된 정보
        """
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                # 파일 존재 확인
                abs_path = self.check_file_exists(image_path)
                if not abs_path:
                    return None
                    
                # 이미지 파일 크기 제한 확인 (ChatGPT Vision API의 파일 크기 제한)
                try:
                    file_size = os.path.getsize(abs_path)
                    max_size = 20 * 1024 * 1024  # 20MB
                    if file_size > max_size:
                        return None  # 파일이 너무 큼
                except Exception:
                    return None
                    
                # 이미지를 base64로 인코딩
                try:
                    with open(abs_path, "rb") as image_file:
                        base64_image = base64.b64encode(image_file.read()).decode('utf-8')
                except PermissionError:
                    return None
                except Exception:
                    return None
                
                # ChatGPT Vision API 호출
                try:
                    response = self.openai_client.chat.completions.create(
                        model="gpt-4o",
                        messages=[
                            {
                                "role": "system",
                                "content": f"당신은 이미지를 분석하여 필요한 정보를 추출하는 전문가입니다. 다음 지시에 따라 이미지를 분석해주세요: {self.prompt}"
                            },
                            {
                                "role": "user",
                                "content": [
                                    {"type": "text", "text": "이 이미지를 분석해서 요청한 정보를 추출해주세요."},
                                    {
                                        "type": "image_url",
                                        "image_url": {
                                            "url": f"data:image/jpeg;base64,{base64_image}"
                                        }
                                    }
                                ]
                            }
                        ],
                        temperature=0.3,
                        max_tokens=100
                    )
                    
                    if response and hasattr(response, 'choices') and len(response.choices) > 0:
                        extracted_info = response.choices[0].message.content.strip()
                        return extracted_info
                    else:
                        # API 응답 오류 재시도
                        retry_count += 1
                        if retry_count < max_retries:
                            time.sleep(1)  # 1초 대기 후 재시도
                            continue
                        return None
                except requests.exceptions.RequestException:
                    # 네트워크 오류 재시도
                    retry_count += 1
                    if retry_count < max_retries:
                        time.sleep(2)  # 2초 대기 후 재시도
                        continue
                    return None
                except Exception:
                    # 기타 오류 재시도
                    retry_count += 1
                    if retry_count < max_retries:
                        time.sleep(1)  # 1초 대기 후 재시도
                        continue
                    return None
                
            except Exception:
                retry_count += 1
                if retry_count < max_retries:
                    time.sleep(1)  # 1초 대기 후 재시도
                    continue
                return None
        
        return None


class GeminiAnalyzer(ImageAnalyzer):
    """Google Gemini API를 사용한 이미지 분석기"""
    
    def __init__(self):
        super().__init__()
        
        # Gemini API 초기화
        try:
            api_key_path = 'gemini-api-key/gemini-api-key.txt'
            
            # 키 파일 존재 확인
            if not os.path.exists(api_key_path):
                raise FileNotFoundError(f"Gemini API 키 파일을 찾을 수 없습니다: {api_key_path}")
                
            with open(api_key_path, 'r') as f:
                api_key = f.read().strip()
                
            # API 키가 비어있는지 확인
            if not api_key:
                raise ValueError("Gemini API 키가 비어 있습니다.")
                
            genai.configure(api_key=api_key)
            
            # 모델 존재 확인
            available_models = [m.name for m in genai.list_models()]
            if 'gemini-2.0-flash-lite' not in available_models and 'models/gemini-2.0-flash-lite' not in available_models:
                model_name = 'gemini-1.5-flash'  # 대체 모델
                self.model = genai.GenerativeModel(model_name)
            else:
                self.model = genai.GenerativeModel('gemini-2.0-flash-lite')
                
        except FileNotFoundError as e:
            raise e
        except ValueError as e:
            raise e
        except Exception as e:
            raise Exception(f"Gemini API 초기화 중 오류: {str(e)}")
    
    def analyze_image(self, image_path):
        """
        Gemini API로 이미지 직접 분석
        
        Args:
            image_path (str): 이미지 파일 경로
        
        Returns:
            str: 이미지에서 추출된 정보
        """
        max_retries = 5  # 최대 재시도 횟수를 3에서 5로 증가
        retry_count = 0
        base_delay = 2  # 기본 대기 시간 (초)
        
        while retry_count < max_retries:
            try:
                # 파일 존재 확인
                abs_path = self.check_file_exists(image_path)
                if not abs_path:
                    print(f"❌ [Gemini API 오류] 유효하지 않은 이미지 파일: {image_path}")
                    return None
                
                # 이미지 파일 크기 제한 확인 (Gemini API의 파일 크기 제한)
                try:
                    file_size = os.path.getsize(abs_path)
                    max_size = 10 * 1024 * 1024  # 10MB
                    if file_size > max_size:
                        print(f"❌ [Gemini API 오류] 이미지 파일이 너무 큽니다 ({file_size/1024/1024:.2f}MB > 10MB)")
                        return None  # 파일이 너무 큼
                except Exception as e:
                    print(f"❌ [Gemini API 오류] 파일 크기 확인 중 오류: {str(e)}")
                    return None
                
                # 이미지 로드
                try:
                    with open(abs_path, "rb") as image_file:
                        image_data = image_file.read()
                except PermissionError as e:
                    print(f"❌ [Gemini API 오류] 파일 접근 권한이 없습니다: {str(e)}")
                    return None
                except Exception as e:
                    print(f"❌ [Gemini API 오류] 이미지 로드 중 오류: {str(e)}")
                    return None
                
                # MIME 타입 결정
                _, ext = os.path.splitext(abs_path)
                mime_type = "image/jpeg"  # 기본값
                if ext.lower() == '.png':
                    mime_type = "image/png"
                elif ext.lower() in ['.gif']:
                    mime_type = "image/gif"
                elif ext.lower() in ['.webp']:
                    mime_type = "image/webp"
                elif ext.lower() in ['.bmp']:
                    mime_type = "image/bmp"
                
                image_parts = [
                    {
                        "mime_type": mime_type,
                        "data": image_data
                    }
                ]
                
                # 지수 백오프 적용 (재시도마다 대기 시간 증가)
                if retry_count > 0:
                    # 2^n 공식 적용 (2, 4, 8, 16, 32초...)
                    current_delay = base_delay * (2 ** (retry_count - 1))
                    # 약간의 랜덤성 추가 (지터)
                    jitter = random.uniform(0.8, 1.2)
                    delay_with_jitter = current_delay * jitter
                    
                    print(f"  ⚠️ [Gemini API 재시도] {retry_count}/{max_retries} (대기: {delay_with_jitter:.2f}초)")
                    time.sleep(delay_with_jitter)
                
                # Gemini에 프롬프트와 함께 전송
                prompt = f"다음 지시에 따라 이미지를 분석해주세요: {self.prompt}\n이미지를 분석하고 요청된 정보만 추출해주세요."
                
                try:
                    print(f"  ✓ Gemini API 요청 중...")
                    response = self.model.generate_content(
                        contents=[prompt, image_parts[0]]
                    )
                    
                    if response and hasattr(response, 'text'):
                        extracted_info = response.text.strip()
                        print(f"  ✓ [Gemini API 요청 성공] 결과를 받았습니다.")
                        return extracted_info
                    else:
                        # API 응답 오류 재시도
                        retry_count += 1
                        print(f"❌ [Gemini API 오류] 유효하지 않은 응답 형식")
                        continue
                except requests.exceptions.RequestException as e:
                    # 네트워크 오류 재시도
                    retry_count += 1
                    print(f"❌ [Gemini API 오류] 네트워크 요청 실패: {str(e)}")
                    continue
                except Exception as e:
                    # 모델 호출 오류 확인 및 재시도
                    retry_count += 1
                    print(f"❌ [Gemini API 오류] API 호출 중 오류: {str(e)}")
                    
                    # 429 에러 감지 (Too Many Requests)
                    if "429" in str(e) or "rate limit" in str(e).lower() or "too many requests" in str(e).lower():
                        print(f"⚠️ [Gemini API 오류] 요청 한도 초과 (Rate Limit) - 지수 백오프 적용 중...")
                    continue
                
            except Exception as e:
                retry_count += 1
                print(f"❌ [Gemini API 오류] 예상치 못한 오류 발생: {str(e)}")
                continue
        
        print(f"❌ [Gemini API 오류] 최대 재시도 횟수({max_retries}회)를 초과했습니다. 분석에 실패했습니다.")
        return None