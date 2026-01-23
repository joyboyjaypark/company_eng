"""
Cloud Storage Module
이전 대화 내용/작업 데이터를 클라우드에 업로드하는 기능
"""
import os
import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict, Any


class CloudStorage:
    """클라우드 스토리지 인터페이스"""
    
    def __init__(self, storage_path: Optional[str] = None):
        """
        Initialize cloud storage
        
        Args:
            storage_path: Local path for cloud storage simulation.
                         If None, uses ~/.company_eng_cloud
        """
        if storage_path is None:
            storage_path = os.path.expanduser("~/.company_eng_cloud")
        
        self.storage_path = Path(storage_path)
        self.storage_path.mkdir(parents=True, exist_ok=True)
        
        # Create subdirectories for different data types
        self.drawings_path = self.storage_path / "drawings"
        self.weather_data_path = self.storage_path / "weather_data"
        self.history_path = self.storage_path / "history"
        
        for path in [self.drawings_path, self.weather_data_path, self.history_path]:
            path.mkdir(parents=True, exist_ok=True)
        
        self.history_file = self.history_path / "upload_history.json"
        self._load_history()
    
    def _load_history(self):
        """업로드 히스토리 로드"""
        if self.history_file.exists():
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    self.upload_history = json.load(f)
            except Exception:
                self.upload_history = []
        else:
            self.upload_history = []
    
    def _save_history(self):
        """업로드 히스토리 저장"""
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self.upload_history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"히스토리 저장 오류: {e}")
    
    def upload_file(self, source_path: str, data_type: str = "drawing", 
                   metadata: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        파일을 클라우드에 업로드
        
        Args:
            source_path: 업로드할 파일 경로
            data_type: 데이터 타입 ("drawing", "weather_data", "other")
            metadata: 추가 메타데이터
        
        Returns:
            업로드 결과 정보
        """
        if not os.path.exists(source_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {source_path}")
        
        # 대상 경로 결정
        if data_type == "drawing":
            dest_dir = self.drawings_path
        elif data_type == "weather_data":
            dest_dir = self.weather_data_path
        else:
            dest_dir = self.storage_path / "other"
            dest_dir.mkdir(parents=True, exist_ok=True)
        
        # 파일명 생성 (타임스탬프 포함)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        source_file = Path(source_path)
        filename = f"{timestamp}_{source_file.name}"
        dest_path = dest_dir / filename
        
        # 파일 복사
        shutil.copy2(source_path, dest_path)
        
        # 업로드 기록 생성
        upload_record = {
            "upload_id": f"{data_type}_{timestamp}",
            "timestamp": datetime.now().isoformat(),
            "source_path": str(source_path),
            "cloud_path": str(dest_path),
            "filename": filename,
            "data_type": data_type,
            "file_size": os.path.getsize(dest_path),
            "metadata": metadata or {}
        }
        
        self.upload_history.append(upload_record)
        self._save_history()
        
        return upload_record
    
    def list_uploads(self, data_type: Optional[str] = None, 
                    limit: int = 50) -> List[Dict[str, Any]]:
        """
        업로드된 파일 목록 조회
        
        Args:
            data_type: 필터링할 데이터 타입 (None이면 전체)
            limit: 반환할 최대 개수
        
        Returns:
            업로드 기록 리스트
        """
        history = self.upload_history
        
        if data_type:
            history = [h for h in history if h.get("data_type") == data_type]
        
        # 최신순 정렬
        history = sorted(history, key=lambda x: x.get("timestamp", ""), reverse=True)
        
        return history[:limit]
    
    def download_file(self, upload_id: str, dest_path: str) -> bool:
        """
        클라우드에서 파일 다운로드
        
        Args:
            upload_id: 업로드 ID
            dest_path: 다운로드할 로컬 경로
        
        Returns:
            성공 여부
        """
        # 업로드 기록 찾기
        record = None
        for r in self.upload_history:
            if r.get("upload_id") == upload_id:
                record = r
                break
        
        if not record:
            raise ValueError(f"업로드 ID를 찾을 수 없습니다: {upload_id}")
        
        cloud_path = record.get("cloud_path")
        if not os.path.exists(cloud_path):
            raise FileNotFoundError(f"클라우드 파일을 찾을 수 없습니다: {cloud_path}")
        
        # 파일 복사
        shutil.copy2(cloud_path, dest_path)
        return True
    
    def delete_upload(self, upload_id: str) -> bool:
        """
        업로드된 파일 삭제
        
        Args:
            upload_id: 업로드 ID
        
        Returns:
            성공 여부
        """
        # 업로드 기록 찾기
        record = None
        record_index = -1
        for i, r in enumerate(self.upload_history):
            if r.get("upload_id") == upload_id:
                record = r
                record_index = i
                break
        
        if not record:
            return False
        
        # 파일 삭제
        cloud_path = record.get("cloud_path")
        if os.path.exists(cloud_path):
            os.remove(cloud_path)
        
        # 히스토리에서 제거
        del self.upload_history[record_index]
        self._save_history()
        
        return True
    
    def get_storage_info(self) -> Dict[str, Any]:
        """
        스토리지 정보 조회
        
        Returns:
            스토리지 정보
        """
        total_size = 0
        file_count = 0
        
        for record in self.upload_history:
            cloud_path = record.get("cloud_path")
            if os.path.exists(cloud_path):
                total_size += os.path.getsize(cloud_path)
                file_count += 1
        
        return {
            "storage_path": str(self.storage_path),
            "total_files": file_count,
            "total_size_bytes": total_size,
            "total_size_mb": round(total_size / (1024 * 1024), 2),
            "upload_count": len(self.upload_history)
        }


# 싱글톤 인스턴스
_cloud_storage_instance = None


def get_cloud_storage() -> CloudStorage:
    """클라우드 스토리지 싱글톤 인스턴스 가져오기"""
    global _cloud_storage_instance
    if _cloud_storage_instance is None:
        _cloud_storage_instance = CloudStorage()
    return _cloud_storage_instance
