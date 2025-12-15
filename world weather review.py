import tkinter as tk
from tkinter import ttk, messagebox
import requests


# API 호출 및 데이터를 표에 표시
def fetch_data():
    try:
        # 사용자 입력값 가져오기
        year = year_combo.get()
        month = month_combo.get().zfill(2)
        day = day_combo.get().zfill(2)
        hour = hour_combo.get().zfill(2)
        minute = minute_combo.get().zfill(2)
        tm = f"{year}{month}{day}{hour}{minute}"
        dtm = dtm_entry.get()
        stn = stn_entry.get()
        auth_key = authkey_entry.get()

        # URL 및 요청 매개변수 생성
        base_url = "https://apihub.kma.go.kr/api/typ01/url/gts_syn1.php"
        params = {
            "authKey": auth_key,
            "tm": tm,
            "dtm": dtm,
            "stn": stn,
            "help": "1",
        }

        # API 요청
        response = requests.get(base_url, params=params)
        if response.status_code == 200:
            data = response.json()
            display_data(data)  # 데이터를 표에 표시
        else:
            messagebox.showerror("API 오류", f"API 호출 실패: HTTP {response.status_code}")
    except Exception as e:
        messagebox.showerror("에러", f"데이터 가져오는 중 오류 발생:\n{str(e)}")


# 데이터를 표에 출력
def display_data(data):
    try:
        # 기존 데이터 삭제
        for row in tree.get_children():
            tree.delete(row)

        # 요구하는 값 추출 및 테이블에 추가
        for record in data.get("results", []):  # "results" 키 기반 데이터 처리
            tree.insert("", "end", values=(
                record.get("TM", "N/A"),       # 시간
                record.get("TA", "N/A"),       # 기온
                record.get("HM", "N/A"),       # 상대습도
                record.get("PA", "N/A"),       # 현지기압
                record.get("TD", "N/A"),       # 이슬점 온도
            ))

    except KeyError:
        messagebox.showerror("데이터 오류", "응답 데이터에 필요한 필드가 없습니다.")
    except Exception as e:
        messagebox.showerror("에러", f"표시 중 오류 발생: {str(e)}")


# GUI 구성
root = tk.Tk()
root.title("기상 데이터 조회 프로그램")

# 입력 섹션
input_frame = tk.Frame(root)
input_frame.pack(pady=10)

# 연월일시분 구성
years = [str(2023), str(2024), str(2025)]
months = [str(i).zfill(2) for i in range(1, 13)]
days = [str(i).zfill(2) for i in range(1, 32)]
hours = [str(i).zfill(2) for i in range(24)]
minutes = [str(i).zfill(2) for i in range(0, 60, 10)]

tk.Label(input_frame, text="연도 (Year): ").grid(row=0, column=0, padx=5, sticky="e")
year_combo = ttk.Combobox(input_frame, width=6, values=years)
year_combo.current(0)
year_combo.grid(row=0, column=1, padx=5)

tk.Label(input_frame, text="월 (Month): ").grid(row=0, column=2, padx=5, sticky="e")
month_combo = ttk.Combobox(input_frame, width=4, values=months)
month_combo.current(0)
month_combo.grid(row=0, column=3, padx=5)

tk.Label(input_frame, text="일 (Day): ").grid(row=1, column=0, padx=5, sticky="e")
day_combo = ttk.Combobox(input_frame, width=4, values=days)
day_combo.current(0)
day_combo.grid(row=1, column=1, padx=5)

tk.Label(input_frame, text="시 (Hour): ").grid(row=1, column=2, padx=5, sticky="e")
hour_combo = ttk.Combobox(input_frame, width=4, values=hours)
hour_combo.current(0)
hour_combo.grid(row=1, column=3, padx=5)

tk.Label(input_frame, text="분 (Minute): ").grid(row=2, column=0, padx=5, sticky="e")
minute_combo = ttk.Combobox(input_frame, width=4, values=minutes)
minute_combo.current(0)
minute_combo.grid(row=2, column=1, padx=5)

tk.Label(input_frame, text="과거 시간 범위 (dtm): ").grid(row=3, column=0, padx=5, sticky="e")
dtm_entry = tk.Entry(input_frame, width=20)
dtm_entry.grid(row=3, column=1, padx=5)

tk.Label(input_frame, text="지점 번호 (stn): ").grid(row=4, column=0, padx=5, sticky="e")
stn_entry = tk.Entry(input_frame, width=20)
stn_entry.grid(row=4, column=1, padx=5)

tk.Label(input_frame, text="인증키 (authKey): ").grid(row=5, column=0, padx=5, sticky="e")
authkey_entry = tk.Entry(input_frame, width=40)
authkey_entry.insert(0, "6MYsKhxBTqaGLCocQU6mXA")  # 기본 인증키
authkey_entry.grid(row=5, column=1, padx=5, columnspan=2)

# 조회 버튼
fetch_button = tk.Button(root, text="조회하기", command=fetch_data)
fetch_button.pack(pady=10)

# 출력 섹션
output_frame = tk.Frame(root)
output_frame.pack(pady=10)

tree = ttk.Treeview(output_frame, columns=("TM", "TA", "HM", "PA", "TD"), show="headings", height=10)
tree.pack()

# 테이블 헤더 설정
tree.heading("TM", text="시간")
tree.heading("TA", text="기온 (℃)")
tree.heading("HM", text="습도 (%)")
tree.heading("PA", text="기압 (hPa)")
tree.heading("TD", text="이슬점온도 (℃)")

# 열 너비 조절
for col in ("TM", "TA", "HM", "PA", "TD"):
    tree.column(col, width=120, anchor="center")

root.mainloop()
