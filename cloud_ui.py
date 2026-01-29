"""
Cloud Upload/Download UI
클라우드 업로드/다운로드 GUI 다이얼로그
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from cloud_storage import get_cloud_storage


class CloudUploadDialog(tk.Toplevel):
    """클라우드 업로드 다이얼로그"""
    
    def __init__(self, parent, file_path=None, data_type="drawing"):
        """
        Initialize upload dialog
        
        Args:
            parent: 부모 윈도우
            file_path: 업로드할 파일 경로 (None이면 파일 선택)
            data_type: 데이터 타입
        """
        super().__init__(parent)
        self.title("클라우드 업로드")
        self.geometry("500x300")
        self.resizable(False, False)
        
        self.file_path = file_path
        self.data_type = data_type
        self.cloud_storage = get_cloud_storage()
        self.result = None
        
        self.create_widgets()
        
        # 모달 다이얼로그로 설정
        self.transient(parent)
        self.grab_set()
    
    def create_widgets(self):
        """위젯 생성"""
        # 파일 선택 프레임
        file_frame = ttk.LabelFrame(self, text="파일 선택", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.file_var = tk.StringVar(value=self.file_path or "")
        file_entry = ttk.Entry(file_frame, textvariable=self.file_var, state='readonly')
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        browse_btn = ttk.Button(file_frame, text="찾아보기", command=self.browse_file)
        browse_btn.pack(side=tk.RIGHT)
        
        # 데이터 타입 선택
        type_frame = ttk.LabelFrame(self, text="데이터 타입", padding=10)
        type_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.type_var = tk.StringVar(value=self.data_type)
        types = [
            ("도면 데이터", "drawing"),
            ("날씨 데이터", "weather_data"),
            ("기타", "other")
        ]
        
        for i, (label, value) in enumerate(types):
            rb = ttk.Radiobutton(type_frame, text=label, variable=self.type_var, value=value)
            rb.grid(row=0, column=i, padx=10)
        
        # 메모 입력
        memo_frame = ttk.LabelFrame(self, text="메모 (선택사항)", padding=10)
        memo_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.memo_text = tk.Text(memo_frame, height=5, wrap=tk.WORD)
        self.memo_text.pack(fill=tk.BOTH, expand=True)
        
        # 버튼
        button_frame = ttk.Frame(self)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        upload_btn = ttk.Button(button_frame, text="업로드", command=self.do_upload)
        upload_btn.pack(side=tk.RIGHT, padx=5)
        
        cancel_btn = ttk.Button(button_frame, text="취소", command=self.destroy)
        cancel_btn.pack(side=tk.RIGHT)
    
    def browse_file(self):
        """파일 찾아보기"""
        file_path = filedialog.askopenfilename(
            title="업로드할 파일 선택",
            filetypes=[
                ("JSON 파일", "*.json"),
                ("Excel 파일", "*.xlsx"),
                ("모든 파일", "*.*")
            ]
        )
        if file_path:
            self.file_var.set(file_path)
    
    def do_upload(self):
        """업로드 실행"""
        file_path = self.file_var.get()
        if not file_path:
            messagebox.showwarning("경고", "업로드할 파일을 선택해주세요.")
            return
        
        try:
            memo = self.memo_text.get("1.0", tk.END).strip()
            metadata = {
                "memo": memo
            }
            
            result = self.cloud_storage.upload_file(
                file_path,
                data_type=self.type_var.get(),
                metadata=metadata
            )
            
            self.result = result
            messagebox.showinfo(
                "업로드 완료",
                f"파일이 클라우드에 업로드되었습니다.\n\n"
                f"파일명: {result['filename']}\n"
                f"크기: {result['file_size']:,} bytes"
            )
            self.destroy()
            
        except Exception as e:
            messagebox.showerror("업로드 오류", f"업로드 중 오류가 발생했습니다:\n{e}")


class CloudBrowserDialog(tk.Toplevel):
    """클라우드 파일 브라우저 다이얼로그"""
    
    def __init__(self, parent):
        """
        Initialize cloud browser
        
        Args:
            parent: 부모 윈도우
        """
        super().__init__(parent)
        self.title("클라우드 스토리지 브라우저")
        self.geometry("800x500")
        
        self.cloud_storage = get_cloud_storage()
        self.selected_item = None
        
        self.create_widgets()
        self.refresh_list()
        
        # 모달 다이얼로그로 설정
        self.transient(parent)
        self.grab_set()
    
    def create_widgets(self):
        """위젯 생성"""
        # 상단 정보 프레임
        info_frame = ttk.Frame(self)
        info_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.info_label = ttk.Label(info_frame, text="클라우드 스토리지 정보를 불러오는 중...")
        self.info_label.pack(anchor=tk.W)
        
        # 필터 프레임
        filter_frame = ttk.Frame(self)
        filter_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(filter_frame, text="필터:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.filter_var = tk.StringVar(value="all")
        filters = [
            ("전체", "all"),
            ("도면", "drawing"),
            ("날씨 데이터", "weather_data"),
            ("기타", "other")
        ]
        
        for label, value in filters:
            rb = ttk.Radiobutton(
                filter_frame,
                text=label,
                variable=self.filter_var,
                value=value,
                command=self.refresh_list
            )
            rb.pack(side=tk.LEFT, padx=5)
        
        # 파일 리스트 프레임
        list_frame = ttk.Frame(self)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 트리뷰
        columns = ("timestamp", "filename", "type", "size", "memo")
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)
        
        self.tree.heading("timestamp", text="업로드 시간")
        self.tree.heading("filename", text="파일명")
        self.tree.heading("type", text="타입")
        self.tree.heading("size", text="크기")
        self.tree.heading("memo", text="메모")
        
        self.tree.column("timestamp", width=150)
        self.tree.column("filename", width=250)
        self.tree.column("type", width=100)
        self.tree.column("size", width=100)
        self.tree.column("memo", width=180)
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 더블클릭 이벤트
        self.tree.bind("<Double-1>", self.on_download)
        
        # 버튼 프레임
        button_frame = ttk.Frame(self)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        download_btn = ttk.Button(button_frame, text="다운로드", command=self.on_download)
        download_btn.pack(side=tk.RIGHT, padx=5)
        
        delete_btn = ttk.Button(button_frame, text="삭제", command=self.on_delete)
        delete_btn.pack(side=tk.RIGHT, padx=5)
        
        refresh_btn = ttk.Button(button_frame, text="새로고침", command=self.refresh_list)
        refresh_btn.pack(side=tk.RIGHT, padx=5)
        
        close_btn = ttk.Button(button_frame, text="닫기", command=self.destroy)
        close_btn.pack(side=tk.LEFT)
    
    def refresh_list(self):
        """파일 목록 새로고침"""
        # 트리뷰 클리어
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 필터 적용
        data_type = None if self.filter_var.get() == "all" else self.filter_var.get()
        
        try:
            uploads = self.cloud_storage.list_uploads(data_type=data_type, limit=100)
            
            for upload in uploads:
                timestamp = upload.get("timestamp", "")
                try:
                    dt = datetime.fromisoformat(timestamp)
                    timestamp = dt.strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    pass
                
                filename = upload.get("filename", "")
                file_type = upload.get("data_type", "")
                size = upload.get("file_size", 0)
                size_str = f"{size:,} B" if size < 1024 else f"{size/1024:.1f} KB"
                memo = upload.get("metadata", {}).get("memo", "")[:50]
                
                self.tree.insert(
                    "",
                    tk.END,
                    values=(timestamp, filename, file_type, size_str, memo),
                    tags=(upload.get("upload_id"),)
                )
            
            # 스토리지 정보 업데이트
            info = self.cloud_storage.get_storage_info()
            self.info_label.config(
                text=f"전체 파일: {info['total_files']}개 | "
                     f"전체 크기: {info['total_size_mb']:.2f} MB | "
                     f"저장 경로: {info['storage_path']}"
            )
            
        except Exception as e:
            messagebox.showerror("오류", f"파일 목록을 불러오는 중 오류가 발생했습니다:\n{e}")
    
    def on_download(self, event=None):
        """다운로드 버튼 클릭"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("경고", "다운로드할 파일을 선택해주세요.")
            return
        
        item = selection[0]
        upload_id = self.tree.item(item, "tags")[0]
        filename = self.tree.item(item, "values")[1]
        
        # 저장 경로 선택
        dest_path = filedialog.asksaveasfilename(
            title="다운로드 위치 선택",
            initialfile=filename,
            defaultextension=".json",
            filetypes=[("모든 파일", "*.*")]
        )
        
        if not dest_path:
            return
        
        try:
            self.cloud_storage.download_file(upload_id, dest_path)
            messagebox.showinfo("다운로드 완료", f"파일이 다운로드되었습니다:\n{dest_path}")
        except Exception as e:
            messagebox.showerror("다운로드 오류", f"다운로드 중 오류가 발생했습니다:\n{e}")
    
    def on_delete(self):
        """삭제 버튼 클릭"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("경고", "삭제할 파일을 선택해주세요.")
            return
        
        item = selection[0]
        upload_id = self.tree.item(item, "tags")[0]
        filename = self.tree.item(item, "values")[1]
        
        if not messagebox.askyesno("삭제 확인", f"'{filename}' 파일을 삭제하시겠습니까?"):
            return
        
        try:
            self.cloud_storage.delete_upload(upload_id)
            messagebox.showinfo("삭제 완료", "파일이 삭제되었습니다.")
            self.refresh_list()
        except Exception as e:
            messagebox.showerror("삭제 오류", f"삭제 중 오류가 발생했습니다:\n{e}")
