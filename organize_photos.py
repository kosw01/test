"""
사진 파일을 촬영 날짜와 시간에 따라 정리하는 스크립트
날짜별 폴더를 만들고, 그 안에 시간별 폴더를 만들어 사진을 이동시킵니다.
"""

import os
import shutil
from datetime import datetime
from pathlib import Path
from PIL import Image
from PIL.ExifTags import TAGS
import sys

def get_image_datetime(image_path):
    """
    이미지 파일에서 촬영 날짜와 시간을 추출합니다.
    
    Args:
        image_path: 이미지 파일 경로
        
    Returns:
        datetime 객체 또는 None (날짜를 찾을 수 없는 경우)
    """
    try:
        with Image.open(image_path) as img:
            # EXIF 데이터 추출
            exifdata = img.getexif()
            
            if exifdata is None:
                return None
            
            # EXIF 태그에서 날짜/시간 찾기
            # DateTime (306): 촬영 날짜/시간
            # DateTimeOriginal (36867): 원본 촬영 날짜/시간
            # DateTimeDigitized (36868): 디지털화 날짜/시간
            
            date_time = None
            
            # DateTimeOriginal을 우선적으로 사용
            if 36867 in exifdata:
                date_time_str = exifdata[36867]
            elif 306 in exifdata:
                date_time_str = exifdata[306]
            elif 36868 in exifdata:
                date_time_str = exifdata[36868]
            else:
                return None
            
            # EXIF 날짜 형식: "YYYY:MM:DD HH:MM:SS"
            try:
                date_time = datetime.strptime(date_time_str, "%Y:%m:%d %H:%M:%S")
            except ValueError:
                # 다른 형식 시도
                try:
                    date_time = datetime.strptime(date_time_str, "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    return None
            
            return date_time
            
    except Exception as e:
        print(f"오류 발생 ({image_path}): {str(e)}")
        return None

def get_file_modified_time(file_path):
    """
    EXIF 데이터가 없는 경우 파일 수정 시간을 사용합니다.
    
    Args:
        file_path: 파일 경로
        
    Returns:
        datetime 객체
    """
    timestamp = os.path.getmtime(file_path)
    return datetime.fromtimestamp(timestamp)

def organize_photos(source_dir="."):
    """
    현재 디렉토리의 모든 사진을 날짜/시간별로 정리합니다.
    
    Args:
        source_dir: 소스 디렉토리 (기본값: 현재 디렉토리)
    """
    source_path = Path(source_dir)
    
    # 지원하는 이미지 확장자
    image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', 
                        '.heic', '.heif', '.raw', '.cr2', '.nef', '.orf', '.sr2'}
    
    # 이미지 파일 찾기
    image_files = []
    for ext in image_extensions:
        image_files.extend(source_path.glob(f"*{ext}"))
        image_files.extend(source_path.glob(f"*{ext.upper()}"))
    
    if not image_files:
        print("이미지 파일을 찾을 수 없습니다.")
        return
    
    print(f"총 {len(image_files)}개의 이미지 파일을 찾았습니다.\n")
    
    # 파일 정리
    organized_count = 0
    failed_count = 0
    
    for image_file in image_files:
        print(f"처리 중: {image_file.name}")
        
        # 촬영 날짜/시간 추출
        date_time = get_image_datetime(image_file)
        
        # EXIF 데이터가 없으면 파일 수정 시간 사용
        if date_time is None:
            print(f"  EXIF 데이터가 없어 파일 수정 시간을 사용합니다.")
            date_time = get_file_modified_time(image_file)
        
        # 날짜 및 시간 폴더명 생성 (10분 단위로 반올림)
        date_folder = date_time.strftime("%Y-%m-%d")
        # 분을 10분 단위로 내림 (예: 23분 → 20분, 37분 → 30분)
        rounded_minute = (date_time.minute // 10) * 10
        rounded_time = date_time.replace(minute=rounded_minute, second=0, microsecond=0)
        time_folder = rounded_time.strftime("%H-%M")
        
        # 대상 폴더 경로 생성
        target_dir = source_path / date_folder / time_folder
        target_dir.mkdir(parents=True, exist_ok=True)
        
        # 파일 이동
        target_file = target_dir / image_file.name
        
        # 같은 이름의 파일이 이미 존재하는 경우 처리
        if target_file.exists():
            # 파일명에 번호 추가
            base_name = image_file.stem
            extension = image_file.suffix
            counter = 1
            while target_file.exists():
                new_name = f"{base_name}_{counter}{extension}"
                target_file = target_dir / new_name
                counter += 1
        
        try:
            shutil.move(str(image_file), str(target_file))
            print(f"  이동 완료: {target_file}")
            organized_count += 1
        except Exception as e:
            print(f"  이동 실패: {str(e)}")
            failed_count += 1
        
        print()
    
    print(f"\n정리 완료!")
    print(f"성공: {organized_count}개")
    print(f"실패: {failed_count}개")

if __name__ == "__main__":
    # 현재 디렉토리에서 실행
    organize_photos(".")
