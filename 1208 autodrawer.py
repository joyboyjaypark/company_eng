"""
duct_auto_router_gui.py

덕트 자동 루팅 GUI 프로토타입

기능:
- GUI 캔버스에서 마우스로:
    - 팬(공급기기) 1개
    - 디퓨저 여러 개
    - No-Go(장애물) 사각형 여러 개
  를 배치
- [자동 루팅] 버튼을 누르면:
    - 격자 그래프 생성 (No-Go 영역 회피)
    - 팬 → 각 디퓨저까지 최소비용 경로 계산
    - 자동 생성된 덕트 네트워크를 화면에 표시

필요 패키지:
    pip install networkx shapely matplotlib
"""

import math
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional

import networkx as nx
from shapely.geometry import Polygon, LineString

import tkinter as tk
from tkinter import ttk, messagebox

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt


# =========================
# 1. 기본 데이터 구조
# =========================

@dataclass
class DuctPoint:
    id: str
    x: float
    y: float
    flow: float = 0.0
    kind: str = "terminal"  # "fan", "terminal"


@dataclass
class NoGoArea:
    polygon: Polygon


@dataclass
class FloorPlanData:
    width: float
    height: float
    supply_fans: List[DuctPoint]
    supply_terminals: List[DuctPoint]
    no_go_areas: List[NoGoArea]
    grid_step: float = 1.0


# =========================
# 2. ML 모델 인터페이스 (더미)
# =========================

class EdgeScoreModel:
    def __init__(self):
        pass

    def predict_edge_score(
        self,
        p1: Tuple[float, float],
        p2: Tuple[float, float],
        context: Dict
    ) -> float:
        # 지금은 모든 엣지에 대해 동일 점수
        return 0.5


# =========================
# 3. 라우터 (그래프 기반 자동 루팅)
# =========================

class DuctRouter:
    def __init__(self, ml_model: Optional[EdgeScoreModel] = None):
        self.ml_model = ml_model or EdgeScoreModel()

    @staticmethod
    def distance(a: Tuple[float, float], b: Tuple[float, float]) -> float:
        return math.hypot(a[0] - b[0], a[1] - b[1])

    def _is_segment_allowed(
        self,
        p1: Tuple[float, float],
        p2: Tuple[float, float],
        no_go_areas: List[NoGoArea]
    ) -> bool:
        line = LineString([p1, p2])
        for area in no_go_areas:
            if line.intersects(area.polygon):
                return False
        return True

    def build_grid_graph(self, plan: FloorPlanData) -> nx.Graph:
        g = nx.Graph()
        step = plan.grid_step

        max_i = int(plan.width / step)
        max_j = int(plan.height / step)

        # 노드
        for i in range(max_i + 1):
            for j in range(max_j + 1):
                x = i * step
                y = j * step
                g.add_node((x, y))

        # 엣지 (4방향)
        for (x, y) in list(g.nodes()):
            neighbors = [
                (x + step, y),
                (x - step, y),
                (x, y + step),
                (x, y - step),
            ]
            for nxp, nyp in neighbors:
                if (nxp, nyp) not in g.nodes():
                    continue
                if not self._is_segment_allowed(
                    (x, y), (nxp, nyp), plan.no_go_areas
                ):
                    continue

                base_len = self.distance((x, y), (nxp, nyp))
                score = self.ml_model.predict_edge_score(
                    (x, y), (nxp, nyp), context={}
                )
                cost = base_len * (1 - 0.3 * score)

                g.add_edge(
                    (x, y), (nxp, nyp),
                    length=base_len,
                    ml_score=score,
                    weight=cost
                )
        return g

    def snap_to_grid(self, x: float, y: float, step: float) -> Tuple[float, float]:
        gx = round(x / step) * step
        gy = round(y / step) * step
        return gx, gy

    def route_supply(self, plan: FloorPlanData) -> Dict:
        if len(plan.supply_fans) != 1:
            raise ValueError("팬은 1개만 배치되어야 합니다.")

        fan = plan.supply_fans[0]
        step = plan.grid_step

        g = self.build_grid_graph(plan)

        fan_node = self.snap_to_grid(fan.x, fan.y, step)
        terminal_nodes = [
            self.snap_to_grid(t.x, t.y, step) for t in plan.supply_terminals
        ]

        paths: Dict[str, List[Tuple[float, float]]] = {}
        for idx, t_node in enumerate(terminal_nodes):
            term_id = plan.supply_terminals[idx].id
            try:
                path = nx.shortest_path(g, fan_node, t_node, weight="weight")
                paths[term_id] = path
            except nx.NetworkXNoPath:
                print(f"[경고] 터미널 {term_id}까지 경로 없음.")
                continue

        network_edges = set()
        for path in paths.values():
            for i in range(len(path) - 1):
                a = path[i]
                b = path[i + 1]
                edge = tuple(sorted([a, b]))
                network_edges.add(edge)

        total_length = 0.0
        for a, b in network_edges:
            total_length += self.distance(a, b)

        result = {
            "fan_node": fan_node,
            "terminal_map": {
                tid: paths[tid][-1] if tid in paths else None
                for tid in paths.keys()
            },
            "edges": list(network_edges),
            "total_length": total_length,
        }
        return result


# =========================
# 4. GUI 애플리케이션
# =========================

class DuctRouterGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("덕트 자동 루팅 GUI 프로토타입")

        # 평면 크기 (m)와 그리드 간격
        self.width = 30.0
        self.height = 20.0
        self.grid_step = 1.0

        # 캔버스 픽셀 크기
        self.canvas_w = 900
        self.canvas_h = 600

        # 좌표 변환 (실좌표[m] <-> 화면좌표[pixel])
        self.scale_x = self.canvas_w / self.width
        self.scale_y = self.canvas_h / self.height

        # 상태
        self.current_mode = tk.StringVar(value="fan")  # "fan", "terminal", "nogo"
        self.fan: Optional[DuctPoint] = None
        self.terminals: List[DuctPoint] = []
        self.nogo_rect_start: Optional[Tuple[float, float]] = None
        self.nogo_areas: List[NoGoArea] = []

        # 라우터
        self.router = DuctRouter()

        self._build_ui()

    # ----- 좌표 변환 -----

    def world_to_screen(self, x: float, y: float) -> Tuple[float, float]:
        sx = x * self.scale_x
        sy = self.canvas_h - y * self.scale_y
        return sx, sy

    def screen_to_world(self, sx: float, sy: float) -> Tuple[float, float]:
        x = sx / self.scale_x
        y = (self.canvas_h - sy) / self.scale_y
        return x, y

    # ----- UI 구성 -----

    def _build_ui(self):
        mainframe = ttk.Frame(self.master)
        mainframe.pack(fill=tk.BOTH, expand=True)

        # 왼쪽: 컨트롤 패널
        control_frame = ttk.Frame(mainframe)
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)

        ttk.Label(control_frame, text="배치 모드").pack(anchor="w", pady=(0, 5))

        ttk.Radiobutton(
            control_frame, text="팬(1개)", variable=self.current_mode,
            value="fan"
        ).pack(anchor="w")

        ttk.Radiobutton(
            control_frame, text="디퓨저", variable=self.current_mode,
            value="terminal"
        ).pack(anchor="w")

        ttk.Radiobutton(
            control_frame, text="No-Go 영역(사각형)", variable=self.current_mode,
            value="nogo"
        ).pack(anchor="w")

        ttk.Separator(control_frame, orient=tk.HORIZONTAL).pack(
            fill=tk.X, pady=10
        )

        ttk.Button(
            control_frame, text="자동 루팅 실행",
            command=self.on_route
        ).pack(fill=tk.X, pady=(0, 5))

        ttk.Button(
            control_frame, text="초기화",
            command=self.on_reset
        ).pack(fill=tk.X, pady=(0, 5))

        self.info_label = ttk.Label(
            control_frame,
            text="좌클릭: 팬/디퓨저 배치\n\n"
                 "No-Go 모드:\n처음 클릭 → 시작점\n두 번째 클릭 → 종료점(사각형)"
        )
        self.info_label.pack(anchor="w", pady=10)

        self.result_label = ttk.Label(control_frame, text="", foreground="blue")
        self.result_label.pack(anchor="w", pady=5)

        # 오른쪽: Matplotlib 캔버스
        fig, ax = plt.subplots(figsize=(9, 6))
        self.fig = fig
        self.ax = ax

        self.ax.set_xlim(0, self.width)
        self.ax.set_ylim(0, self.height)
        self.ax.set_aspect("equal", adjustable="box")
        self.ax.grid(True, linestyle="--", alpha=0.3)
        self.ax.set_title("평면 상에서 팬/디퓨저/장애물 배치 후 '자동 루팅' 실행")

        self.canvas = FigureCanvasTkAgg(self.fig, master=mainframe)
        self.canvas.get_tk_widget().pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.canvas.mpl_connect("button_press_event", self.on_canvas_click)

        self.redraw()

    # ----- GUI 이벤트 -----

    def on_canvas_click(self, event):
        if event.xdata is None or event.ydata is None:
            return

        x, y = event.xdata, event.ydata

        mode = self.current_mode.get()

        if mode == "fan":
            # 팬은 1개만 허용
            self.fan = DuctPoint(id="F1", x=x, y=y, kind="fan", flow=0.0)
        elif mode == "terminal":
            idx = len(self.terminals) + 1
            self.terminals.append(
                DuctPoint(id=f"T{idx}", x=x, y=y, kind="terminal", flow=0.0)
            )
        elif mode == "nogo":
            # 사각형: 첫 클릭은 시작점, 두 번째 클릭은 반대 코너
            if self.nogo_rect_start is None:
                self.nogo_rect_start = (x, y)
            else:
                x0, y0 = self.nogo_rect_start
                x1, y1 = x, y
                xmin, xmax = min(x0, x1), max(x0, x1)
                ymin, ymax = min(y0, y1), max(y0, y1)
                poly = Polygon([(xmin, ymin), (xmax, ymin),
                                (xmax, ymax), (xmin, ymax)])
                self.nogo_areas.append(NoGoArea(polygon=poly))
                self.nogo_rect_start = None

        self.redraw()

    def on_reset(self):
        self.fan = None
        self.terminals.clear()
        self.nogo_areas.clear()
        self.nogo_rect_start = None
        self.result_label.config(text="")
        self.redraw()

    def on_route(self):
        if self.fan is None:
            messagebox.showwarning("경고", "팬을 1개 배치해 주세요.")
            return
        if len(self.terminals) == 0:
            messagebox.showwarning("경고", "디퓨저를 1개 이상 배치해 주세요.")
            return

        plan = FloorPlanData(
            width=self.width,
            height=self.height,
            supply_fans=[self.fan],
            supply_terminals=self.terminals,
            no_go_areas=self.nogo_areas,
            grid_step=self.grid_step
        )

        try:
            result = self.router.route_supply(plan)
        except ValueError as e:
            messagebox.showerror("오류", str(e))
            return

        # 라우팅 결과를 저장해 그림에 반영
        self.last_route_result = result
        self.result_label.config(
            text=f"총 덕트 길이(격자 기준): {result['total_length']:.2f} m\n"
                 f"엣지 개수: {len(result['edges'])}"
        )
        self.redraw()

    # ----- 그리기 -----

    def redraw(self):
        self.ax.clear()
        self.ax.set_xlim(0, self.width)
        self.ax.set_ylim(0, self.height)
        self.ax.set_aspect("equal", adjustable="box")
        self.ax.grid(True, linestyle="--", alpha=0.3)
        self.ax.set_xlabel("X (m)")
        self.ax.set_ylabel("Y (m)")

        # No-Go 영역
        for area in self.nogo_areas:
            x, y = area.polygon.exterior.xy
            self.ax.fill(x, y, color="lightgray", alpha=0.5, label="No-Go")

        # 팬
        if self.fan is not None:
            self.ax.scatter(self.fan.x, self.fan.y,
                            c="red", s=80, marker="s", label="Fan")
            self.ax.text(self.fan.x, self.fan.y + 0.3, "F1",
                         color="red", ha="center")

        # 디퓨저
        for t in self.terminals:
            self.ax.scatter(t.x, t.y,
                            c="blue", s=40, marker="o")
            self.ax.text(t.x, t.y + 0.3, t.id,
                         color="blue", ha="center")

        # 마지막 자동 루팅 결과 표시
        result = getattr(self, "last_route_result", None)
        if result is not None:
            edges = result["edges"]
            xs = []
            ys = []
            for (a, b) in edges:
                xs.extend([a[0], b[0], None])
                ys.extend([a[1], b[1], None])
            self.ax.plot(xs, ys, color="green", linewidth=2,
                         label="자동 루트(격자기준)")

        self.ax.legend(loc="upper right")
        self.canvas.draw_idle()


# =========================
# 5. 실행
# =========================

def main():
    root = tk.Tk()
    app = DuctRouterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
