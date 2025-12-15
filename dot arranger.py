import tkinter as tk
from tkinter import simpledialog, messagebox
import numpy as np
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.patches import Polygon, Circle
from matplotlib.path import Path
import math

GRID_STEP = 0.2  # 0.2 m

class DiffuserLayoutApp:
    def __init__(self, master):
        self.master = master
        master.title("Diffuser Auto Layout")

        self.fig = Figure(figsize=(6, 6))
        self.ax = self.fig.add_subplot(111)
        self.ax.set_aspect("equal")
        self.ax.set_xlim(0, 10)
        self.ax.set_ylim(0, 10)
        self.ax.grid(True, which="both", linestyle="--", linewidth=0.5)

        self.canvas = FigureCanvasTkAgg(self.fig, master=master)
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        frame = tk.Frame(master)
        frame.pack(side=tk.BOTTOM, fill=tk.X)

        self.btn_close_poly = tk.Button(frame, text="실 닫기(폴리곤 완료)", command=self.close_polygon)
        self.btn_close_poly.pack(side=tk.LEFT, padx=5, pady=5)

        self.btn_auto_layout = tk.Button(frame, text="디퓨저 자동 배치", command=self.auto_layout)
        self.btn_auto_layout.pack(side=tk.LEFT, padx=5, pady=5)

        self.btn_clear = tk.Button(frame, text="초기화", command=self.clear_all)
        self.btn_clear.pack(side=tk.LEFT, padx=5, pady=5)

        self.cid = self.canvas.mpl_connect("button_press_event", self.on_click)

        self.vertices = []      # polygon vertices
        self.poly_patch = None  # matplotlib Polygon
        self.diffuser_patches = []

    def snap_to_grid(self, x, y):
        gx = round(x / GRID_STEP) * GRID_STEP
        gy = round(y / GRID_STEP) * GRID_STEP
        return gx, gy

    def on_click(self, event):
        # 좌클릭으로 꼭짓점 추가
        if event.inaxes != self.ax:
            return
        if event.button == 1:
            x, y = self.snap_to_grid(event.xdata, event.ydata)
            self.vertices.append((x, y))
            self.draw_polygon_preview()

    def draw_polygon_preview(self):
        self.ax.clear()
        self.ax.set_aspect("equal")
        self.ax.set_xlim(0, 10)
        self.ax.set_ylim(0, 10)
        self.ax.grid(True, which="both", linestyle="--", linewidth=0.5)

        if len(self.vertices) > 0:
            xs, ys = zip(*self.vertices)
            self.ax.plot(xs, ys, "bo-")

        # 이미 배치된 디퓨저 다시 그림
        for c in self.diffuser_patches:
            self.ax.add_patch(c)

        self.canvas.draw()

    def close_polygon(self):
        if len(self.vertices) < 3:
            messagebox.showwarning("경고", "꼭짓점이 3개 이상이어야 합니다.")
            return
        # 폴리곤 닫기
        if self.poly_patch:
            self.poly_patch.remove()
        self.poly_patch = Polygon(self.vertices, closed=True, fill=False, edgecolor="k", linewidth=2)
        self.ax.add_patch(self.poly_patch)

        xs, ys = zip(*self.vertices)
        self.ax.plot(xs + (xs[0],), ys + (ys[0],), "ko-")

        self.canvas.draw()

    def point_in_poly(self, x, y):
        if not self.poly_patch:
            return False
        path = Path(self.poly_patch.get_xy())
        return path.contains_point((x, y))

    def _dist_point_to_segment(self, px, py, x1, y1, x2, y2):
        dx = x2 - x1
        dy = y2 - y1
        if dx == 0 and dy == 0:
            return math.hypot(px - x1, py - y1)
        t = ((px - x1) * dx + (py - y1) * dy) / (dx * dx + dy * dy)
        t = max(0.0, min(1.0, t))
        proj_x = x1 + t * dx
        proj_y = y1 + t * dy
        return math.hypot(px - proj_x, py - proj_y)

    def _dist_to_polygon_edges(self, px, py):
        if not self.vertices:
            return float('inf')
        min_d = float('inf')
        n = len(self.vertices)
        for i in range(n):
            x1, y1 = self.vertices[i]
            x2, y2 = self.vertices[(i + 1) % n]
            d = self._dist_point_to_segment(px, py, x1, y1, x2, y2)
            if d < min_d:
                min_d = d
        return min_d

    def _select_points_greedy_maxmin(self, candidates, n):
        if n <= 0 or not candidates:
            return []
        cand = list(candidates)
        cand_sorted = sorted(cand, key=lambda p: (p[0], p[1]))
        mid = len(cand_sorted) // 2
        selected = [cand_sorted[mid]]
        remaining = cand_sorted[:mid] + cand_sorted[mid+1:]
        while len(selected) < n and remaining:
            best = None
            best_min_d = -1
            for p in remaining:
                min_d = min(math.hypot(p[0]-s[0], p[1]-s[1]) for s in selected)
                if min_d > best_min_d:
                    best_min_d = min_d
                    best = p
            if best is None:
                break
            selected.append(best)
            remaining.remove(best)
        return selected

    def generate_candidate_grid(self, margin=0.4):
        # 폴리곤 bounding box
        vx, vy = zip(*self.vertices)
        xmin, xmax = min(vx), max(vx)
        ymin, ymax = min(vy), max(vy)

        xs = np.arange(xmin, xmax + GRID_STEP, GRID_STEP)
        ys = np.arange(ymin, ymax + GRID_STEP, GRID_STEP)

        candidates = []
        for x in xs:
            for y in ys:
                if not self.point_in_poly(x, y):
                    continue
                # 벽 이격 margin 반영: 실제 폴리곤 에지까지의 거리로 판단
                # margin 인자는 미터 단위
                if self._dist_to_polygon_edges(x, y) < margin:
                    continue
                candidates.append((x, y))
        return candidates

    def auto_layout(self):
        if not self.poly_patch:
            messagebox.showwarning("경고", "먼저 실(폴리곤)을 닫아 주세요.")
            return

        n = simpledialog.askinteger("디퓨저 개수", "짝수 개의 디퓨저 수를 입력하세요:", minvalue=2)
        if n is None:
            return
        if n % 2 != 0:
            messagebox.showwarning("경고", "짝수만 허용됩니다.")
            return

        # 후보 격자점 생성 (폴리곤 에지로부터 일정 거리 이상)
        min_edge_clearance = 0.8  # meters
        candidates = self.generate_candidate_grid(margin=min_edge_clearance)
        # if insufficient, retry with relaxed margin
        if len(candidates) < n:
            candidates = self.generate_candidate_grid(margin=min_edge_clearance * 0.5)
        if len(candidates) < n:
            messagebox.showwarning("오류", "후보 점이 디퓨저 수보다 적습니다. 실 크기나 margin을 조정하세요.")
            return

        xs = [p[0] for p in candidates]
        ys = [p[1] for p in candidates]
        xmin, xmax = min(xs), max(xs)
        ymin, ymax = min(ys), max(ys)

        # 행/열 수 근사 결정 (r*c ~ n)
        best_r, best_c = 1, n
        best_err = float('inf')
        max_r_candidate = max(1, int(math.sqrt(n)) + 2)
        for r in range(1, max_r_candidate + 1):
            c = max(1, int(math.ceil(n / r)))
            ratio_grid = c / r
            ratio_room = (xmax - xmin) / (ymax - ymin + 1e-6)
            err = abs(ratio_grid - ratio_room)
            if err < best_err:
                best_err = err
                best_r, best_c = r, c

        r, c = best_r, best_c

        # 이상적인 grid 생성
        if r > 1:
            dy = (ymax - ymin) / (r - 1)
        else:
            dy = 0.0
        if c > 1:
            dx = (xmax - xmin) / (c - 1)
        else:
            dx = 0.0

        ideal_points = []
        for i in range(r):
            for j in range(c):
                if len(ideal_points) >= n:
                    break
                x = xmin + j * dx
                y = ymin + i * dy
                ideal_points.append((x, y))
            if len(ideal_points) >= n:
                break

        # ideal -> nearest candidate snap
        selected = []
        used_idx = set()
        for ix, iy in ideal_points:
            best_idx = None
            best_d2 = float('inf')
            for k, (cx, cy) in enumerate(candidates):
                if k in used_idx:
                    continue
                d2 = (cx - ix) ** 2 + (cy - iy) ** 2
                if d2 < best_d2:
                    best_d2 = d2
                    best_idx = k
            if best_idx is not None:
                used_idx.add(best_idx)
                selected.append(candidates[best_idx])

        # 부족하면 greedy max-min 으로 보충
        if len(selected) < n:
            remaining = [p for i, p in enumerate(candidates) if i not in used_idx]
            more = self._select_points_greedy_maxmin(remaining, n - len(selected))
            selected.extend(more)

        # 최종 검증: pairwise spacing 및 edge clearance
        ok = True
        for i in range(len(selected)):
            for j in range(i + 1, len(selected)):
                d = math.hypot(selected[i][0] - selected[j][0], selected[i][1] - selected[j][1])
                if d < min_edge_clearance:
                    ok = False
        if not ok:
            messagebox.showinfo("알림", "일부 디퓨저 간 거리가 최소 이격거리보다 작을 수 있습니다.\n실 크기/개수/spacing을 조정하세요.")

        # 화면에 디퓨저 표시
        for c in self.diffuser_patches:
            c.remove()
        self.diffuser_patches = []

        for x, y in selected[:n]:
            circ = Circle((x, y), 0.1, color="red")
            self.diffuser_patches.append(circ)
            self.ax.add_patch(circ)

        self.canvas.draw()

    def clear_all(self):
        self.vertices = []
        if self.poly_patch:
            self.poly_patch.remove()
            self.poly_patch = None
        for c in self.diffuser_patches:
            c.remove()
        self.diffuser_patches = []
        self.draw_polygon_preview()


if __name__ == "__main__":
    root = tk.Tk()
    app = DiffuserLayoutApp(root)
    root.mainloop()