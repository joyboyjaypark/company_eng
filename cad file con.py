import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import ezdxf
from math import isclose

# matplotlib를 Tkinter에 임베드하기 위한 import
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

ROUND_DIGITS = 4  # 좌표 반올림 자릿수


def normalize_point(p):
    if len(p) == 2:
        x, y = p
        z = 0.0
    else:
        x, y, z = p
    return (
        round(x, ROUND_DIGITS),
        round(y, ROUND_DIGITS),
        round(z, ROUND_DIGITS),
    )


def get_entities(doc):
    """
    DXF 문서에서 기본 엔티티들을 추출해서
    그래픽 표시에도 쓸 수 있는 구조로 반환
    return: [{'type': 'LINE', 'layer': ..., 'data': {...}}, ...]
    """
    msp = doc.modelspace()
    entities = []

    for e in msp:
        etype = e.dxftype()

        # LINE
        if etype == "LINE":
            start = normalize_point(e.dxf.start)
            end = normalize_point(e.dxf.end)
            entities.append({
                "type": "LINE",
                "layer": e.dxf.layer,
                "data": {
                    "start": start,
                    "end": end,
                }
            })

        # CIRCLE
        elif etype == "CIRCLE":
            center = normalize_point(e.dxf.center)
            radius = round(e.dxf.radius, ROUND_DIGITS)
            entities.append({
                "type": "CIRCLE",
                "layer": e.dxf.layer,
                "data": {
                    "center": center,
                    "radius": radius,
                }
            })

        # ARC
        elif etype == "ARC":
            center = normalize_point(e.dxf.center)
            radius = round(e.dxf.radius, ROUND_DIGITS)
            start_angle = round(e.dxf.start_angle, ROUND_DIGITS)
            end_angle = round(e.dxf.end_angle, ROUND_DIGITS)
            entities.append({
                "type": "ARC",
                "layer": e.dxf.layer,
                "data": {
                    "center": center,
                    "radius": radius,
                    "start_angle": start_angle,
                    "end_angle": end_angle,
                }
            })

        # LWPOLYLINE
        elif etype == "LWPOLYLINE":
            points = [normalize_point((p[0], p[1], 0.0)) for p in e.get_points()]
            entities.append({
                "type": "LWPOLYLINE",
                "layer": e.dxf.layer,
                "data": {
                    "points": points,
                    "closed": e.closed,
                }
            })

        # TEXT / MTEXT (좌표만 찍어서 대략적인 위치 확인용, 여기서는 단순 표시만)
        elif etype in ("TEXT", "MTEXT"):
            insert = normalize_point(
                e.dxf.insert if hasattr(e.dxf, "insert") else (0, 0, 0)
            )
            text_content = e.plain_text() if hasattr(e, "plain_text") else e.dxf.text
            entities.append({
                "type": etype,
                "layer": e.dxf.layer,
                "data": {
                    "insert": insert,
                    "text": text_content.strip()
                }
            })

    return entities


def entity_key(ent):
    """
    엔티티를 비교 가능한 불변 키로 변환
    (그래픽용 데이터는 그대로 ent에 두고, 비교는 key로만 수행)
    """
    t = ent["type"]
    layer = ent["layer"]
    d = ent["data"]

    if t == "LINE":
        return ("LINE", layer, d["start"], d["end"])
    elif t == "CIRCLE":
        return ("CIRCLE", layer, d["center"], d["radius"])
    elif t == "ARC":
        return ("ARC", layer, d["center"], d["radius"],
                d["start_angle"], d["end_angle"])
    elif t == "LWPOLYLINE":
        return ("LWPOLYLINE", layer, tuple(d["points"]), d["closed"])
    elif t in ("TEXT", "MTEXT"):
        return (t, layer, d["insert"], d["text"])
    else:
        # 그 외는 문자열로만 구분(단순화)
        return (t, layer, str(d))


def compare_entity_lists(old_entities, new_entities):
    """
    엔티티 리스트 두 개를 비교해서
    - old에만 있는 것들 (deleted)
    - new에만 있는 것들 (added)
    을 원본 엔티티 구조로 반환
    """
    old_map = {entity_key(e): e for e in old_entities}
    new_map = {entity_key(e): e for e in new_entities}

    old_keys = set(old_map.keys())
    new_keys = set(new_map.keys())

    deleted_keys = old_keys - new_keys
    added_keys = new_keys - old_keys

    deleted_entities = [old_map[k] for k in deleted_keys]
    added_entities = [new_map[k] for k in added_keys]

    return deleted_entities, added_entities


class CADCompareViewer:
    def __init__(self, master):
        self.master = master
        master.title("CAD 도면 비교 뷰어 (DXF)")

        self.old_path = tk.StringVar()
        self.new_path = tk.StringVar()

        # 상단 파일 선택 영역
        file_frame = ttk.LabelFrame(master, text="파일 선택")
        file_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(file_frame, text="Old DXF:").grid(row=0, column=0, sticky="e", padx=5, pady=3)
        ttk.Entry(file_frame, textvariable=self.old_path, width=50).grid(row=0, column=1, padx=5, pady=3)
        ttk.Button(file_frame, text="찾기", command=self.browse_old).grid(row=0, column=2, padx=5, pady=3)

        ttk.Label(file_frame, text="New DXF:").grid(row=1, column=0, sticky="e", padx=5, pady=3)
        ttk.Entry(file_frame, textvariable=self.new_path, width=50).grid(row=1, column=1, padx=5, pady=3)
        ttk.Button(file_frame, text="찾기", command=self.browse_new).grid(row=1, column=2, padx=5, pady=3)

        ttk.Button(file_frame, text="비교 실행", command=self.run_compare).grid(
            row=0, column=3, rowspan=2, padx=10, pady=3
        )

        # 요약 라벨
        self.summary_label = ttk.Label(master, text="결과: 아직 비교하지 않았습니다.")
        self.summary_label.pack(pady=3)

        # matplotlib Figure 생성
        self.fig = Figure(figsize=(7, 5), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_aspect("equal", adjustable="box")
        self.ax.grid(True, linestyle="--", alpha=0.3)

        # Tkinter에 matplotlib 캔버스 임베드
        self.canvas = FigureCanvasTkAgg(self.fig, master)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill="both", expand=True, padx=10, pady=5)

        # 범례 설명 라벨
        legend_frame = ttk.Frame(master)
        legend_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(legend_frame, text="파란색: New에만 있는 요소(추가됨)").pack(side="left", padx=5)
        ttk.Label(legend_frame, text="빨간색: Old에만 있는 요소(삭제됨)").pack(side="left", padx=5)

    def browse_old(self):
        path = filedialog.askopenfilename(
            title="Old 도면(DXF) 선택",
            filetypes=[("DXF files", "*.dxf"), ("All files", "*.*")]
        )
        if path:
            self.old_path.set(path)

    def browse_new(self):
        path = filedialog.askopenfilename(
            title="New 도면(DXF) 선택",
            filetypes=[("DXF files", "*.dxf"), ("All files", "*.*")]
        )
        if path:
            self.new_path.set(path)

    def run_compare(self):
        old_path = self.old_path.get()
        new_path = self.new_path.get()

        if not old_path or not new_path:
            messagebox.showwarning("경고", "Old와 New DXF 파일을 모두 선택해 주세요.")
            return

        try:
            old_doc = ezdxf.readfile(old_path)
            new_doc = ezdxf.readfile(new_path)
        except Exception as e:
            messagebox.showerror("에러", f"DXF 파일을 여는 중 오류가 발생했습니다:\n{e}")
            return

        # 엔티티 추출
        old_entities = get_entities(old_doc)
        new_entities = get_entities(new_doc)

        # 비교
        deleted, added = compare_entity_lists(old_entities, new_entities)

        # 그래픽 갱신
        self.draw_result(deleted, added)

        # 요약 텍스트
        self.summary_label.config(
            text=f"비교 완료 - 추가된 엔티티: {len(added)}개, 삭제된 엔티티: {len(deleted)}개"
        )

    def draw_result(self, deleted, added):
        # 축 초기화
        self.ax.clear()
        self.ax.set_aspect("equal", adjustable="box")
        self.ax.grid(True, linestyle="--", alpha=0.3)

        xs = []
        ys = []

        # 추가된(파란색) 먼저 그리기
        for ent in added:
            self._draw_entity(ent, color="blue", xs=xs, ys=ys)

        # 삭제된(빨간색) 그리기
        for ent in deleted:
            self._draw_entity(ent, color="red", xs=xs, ys=ys)

        # 데이터가 있으면 영역 맞추기
        if xs and ys:
            margin = 10
            xmin, xmax = min(xs) - margin, max(xs) + margin
            ymin, ymax = min(ys) - margin, max(ys) + margin
            # 좌표축 뒤집기(도면 좌표계와 유사하게 보이게 하고 싶다면 선택)
            self.ax.set_xlim(xmin, xmax)
            self.ax.set_ylim(ymin, ymax)
        else:
            self.ax.set_xlim(-100, 100)
            self.ax.set_ylim(-100, 100)

        self.ax.set_xlabel("X")
        self.ax.set_ylabel("Y")
        self.ax.set_title("도면 변경사항 그래픽 뷰어")

        self.canvas.draw()

    def _draw_entity(self, ent, color, xs, ys):
        t = ent["type"]
        d = ent["data"]

        if t == "LINE":
            x1, y1, _ = d["start"]
            x2, y2, _ = d["end"]
            self.ax.plot([x1, x2], [y1, y2], color=color, linewidth=1)
            xs.extend([x1, x2])
            ys.extend([y1, y2])

        elif t == "CIRCLE":
            cx, cy, _ = d["center"]
            r = d["radius"]
            circle = self._circle_patch(cx, cy, r, color=color)
            self.ax.add_patch(circle)
            xs.extend([cx - r, cx + r])
            ys.extend([cy - r, cy + r])

        elif t == "ARC":
            # 간단하게는 원호도 원 전체를 그리거나,
            # 혹은 numpy로 세그먼트 나눠서 그려도 됨. 여기서는 전체 원으로 단순화(예시).
            cx, cy, _ = d["center"]
            r = d["radius"]
            circle = self._circle_patch(cx, cy, r, color=color, linestyle="--")
            self.ax.add_patch(circle)
            xs.extend([cx - r, cx + r])
            ys.extend([cy - r, cy + r])

        elif t == "LWPOLYLINE":
            pts = d["points"]
            if len(pts) >= 2:
                xs_line = [p[0] for p in pts]
                ys_line = [p[1] for p in pts]
                if d["closed"]:
                    xs_line.append(pts[0][0])
                    ys_line.append(pts[0][1])
                self.ax.plot(xs_line, ys_line, color=color, linewidth=1)
                xs.extend(xs_line)
                ys.extend(ys_line)

        elif t in ("TEXT", "MTEXT"):
            x, y, _ = d["insert"]
            # 텍스트 위치에 작은 점 + 간단한 텍스트
            self.ax.plot(x, y, marker="o", color=color, markersize=3)
            self.ax.text(x, y, d["text"], color=color, fontsize=6)
            xs.append(x)
            ys.append(y)

    def _circle_patch(self, cx, cy, r, color="blue", linestyle="-"):
        from matplotlib.patches import Circle
        return Circle((cx, cy), r, edgecolor=color, facecolor="none", linestyle=linestyle, linewidth=1)


if __name__ == "__main__":
    root = tk.Tk()
    app = CADCompareViewer(root)
    root.geometry("1000x700")
    root.mainloop()
