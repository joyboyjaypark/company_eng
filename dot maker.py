import tkinter as tk
from tkinter import simpledialog, messagebox
from math import hypot
import math
import random


class OrthogonalPolygonDrawer:
    def __init__(self, root):
        self.root = root
        self.root.title("직각 다각형 작도 (풍선 성장 시뮬레이션)")

        # 캔버스 설정
        self.CANVAS_WIDTH = 800
        self.CANVAS_HEIGHT = 600

        # 월드 좌표 기준: 0.2m 를 20px 로 가정 → 1px = 1cm
        self.base_grid_px = 20     # 0.2m = 20px
        self.scale_factor = 1.0    # 화면 확대/축소 배율
        self.offset_x = 0.0        # 월드 -> 화면 변환용 오프셋
        self.offset_y = 0.0

        # 첫 점에 이 거리(월드 좌표) 이내로 오면 도형을 닫음
        self.CLOSE_WORLD_THRESHOLD = self.base_grid_px * 0.5

        # 상태 변수 (월드 좌표 저장)
        self.points = []             # 다각형 꼭짓점들
        self.polygon_closed = False  # 도형 닫힘 여부

        self.placed_points_ids = []  # "점 찍기" 점
        self.rect_ids = []           # "사각형 배치" 사각형들

        # 풍선 데이터: 각 원소 {"cx","cy","r","vx","vy"}
        self.balloons = []
        self.balloon_ids = []
        self.balloon_running = False

        # 풍선 면적 성장 속도 (cm²/s) – 텍스트박스로 입력
        self.balloon_area_rate = 20.0  # 기본: 20 cm²/s

        self.create_widgets()

        self.grid_ids = []
        self.polygon_line_ids = []
        self.polygon_fill_id = None
        self.vertex_ids = []

        self.redraw_all()

    # ===========================
    # 좌표 변환
    # ===========================
    def world_to_screen(self, xw, yw):
        return xw * self.scale_factor + self.offset_x, yw * self.scale_factor + self.offset_y

    def screen_to_world(self, xs, ys):
        return (xs - self.offset_x) / self.scale_factor, (ys - self.offset_y) / self.scale_factor

    # ===========================
    # UI
    # ===========================
    def create_widgets(self):
        top_frame = tk.Frame(self.root)
        top_frame.pack(side=tk.TOP, fill=tk.X)

        info_label = tk.Label(
            top_frame,
            text=(
                "캔버스를 클릭해 직각 다각형을 그리세요.\n"
                "첫 점 근처에서 클릭하면 도형을 닫습니다.\n"
                "도형 완성 후 '점 찍기', '사각형 배치', '풍선 넣기'를 사용하세요.\n"
                "풍선은 다각형 경계 내부에서만 존재하며, 서로 및 벽과 겹치지 않도록 보정됩니다.\n"
                "우측 상단 '성장속도(cm²/s)'에 값을 입력 후 Enter 를 누르면 즉시 반영됩니다.\n"
                "각 풍선 면적이 (다각형면적/풍선개수)*0.785 를 넘으면 그 풍선의 성장은 멈춥니다.\n"
                "'풍선완료' 버튼: 성장 과정을 생략하고 거의 완성 상태를 바로 보여줍니다.\n"
                "'정렬하기' 버튼: 현재 풍선 중심을 다각형 내부에서 가로/세로 격자로 정렬합니다\n"
                "(정렬 시에는 풍선 경계가 다각형 밖으로 나갈 수도 있습니다)."
            ),
            justify="left"
        )
        info_label.pack(side=tk.LEFT, padx=5, pady=5)

        right_frame = tk.Frame(top_frame)
        right_frame.pack(side=tk.RIGHT, padx=5, pady=5)

        btn_frame = tk.Frame(right_frame)
        btn_frame.pack(side=tk.TOP, anchor="e")

        tk.Button(btn_frame, text="새 도형", command=self.reset).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame, text="줌 +", command=self.zoom_in).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame, text="줌 -", command=self.zoom_out).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame, text="점 찍기", command=self.place_points_in_polygon).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame, text="사각형 배치", command=self.place_rectangles_in_polygon).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame, text="풍선 넣기", command=self.start_balloons).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame, text="풍선완료", command=self.fast_forward_growth).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame, text="정렬하기", command=self.align_balloons_grid).pack(side=tk.LEFT, padx=3)

        # 성장속도 입력 (텍스트 박스)
        speed_frame = tk.Frame(right_frame)
        speed_frame.pack(side=tk.TOP, anchor="e", pady=(5, 0))

        tk.Label(speed_frame, text="성장속도(cm²/s)").pack(side=tk.LEFT)

        self.speed_entry_var = tk.StringVar()
        self.speed_entry_var.set(str(self.balloon_area_rate))
        self.speed_entry = tk.Entry(speed_frame, width=8, textvariable=self.speed_entry_var)
        self.speed_entry.pack(side=tk.LEFT, padx=3)
        self.speed_entry.bind("<Return>", self.on_speed_enter)

        self.canvas = tk.Canvas(self.root, width=self.CANVAS_WIDTH, height=self.CANVAS_HEIGHT, bg="white")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Button-1>", self.on_click)

    def on_speed_enter(self, event=None):
        """텍스트 박스에서 Enter 입력 시 성장속도(cm²/s)를 갱신."""
        text = self.speed_entry_var.get().strip()
        try:
            v = float(text)
            if v <= 0:
                raise ValueError
            self.balloon_area_rate = v
        except ValueError:
            messagebox.showwarning("입력 오류", "성장속도는 0보다 큰 숫자로 입력해주세요.")
            self.speed_entry_var.set(str(self.balloon_area_rate))

    # ===========================
    # 줌 / 리셋
    # ===========================
    def zoom_in(self):
        self.scale_factor *= 1.25
        self.redraw_all()

    def zoom_out(self):
        self.scale_factor /= 1.25
        if self.scale_factor < 0.1:
            self.scale_factor = 0.1
        self.redraw_all()

    def reset(self):
        self.points.clear()
        self.polygon_closed = False

        for pid in self.placed_points_ids:
            self.canvas.delete(pid)
        self.placed_points_ids.clear()

        for rid in self.rect_ids:
            self.canvas.delete(rid)
        self.rect_ids.clear()

        for bid in self.balloon_ids:
            self.canvas.delete(bid)
        self.balloon_ids.clear()

        self.balloons.clear()
        self.balloon_running = False

        self.redraw_all()

    def redraw_all(self):
        self.canvas.delete("all")
        self.grid_ids = []
        self.polygon_line_ids = []
        self.vertex_ids = []
        self.polygon_fill_id = None

        self.draw_grid()
        self.draw_polygon()
        if self.balloons:
            self.draw_balloons()

    # ===========================
    # 격자
    # ===========================
    def draw_grid(self):
        grid_screen = self.base_grid_px * self.scale_factor
        if grid_screen <= 2:
            return

        xw_min = (0 - self.offset_x) / self.scale_factor
        xw_max = (self.CANVAS_WIDTH - self.offset_x) / self.scale_factor
        yw_min = (0 - self.offset_y) / self.scale_factor
        yw_max = (self.CANVAS_HEIGHT - self.offset_y) / self.scale_factor

        gx_start = int(xw_min // self.base_grid_px) * self.base_grid_px
        gx_end   = int(xw_max // self.base_grid_px) * self.base_grid_px
        gy_start = int(yw_min // self.base_grid_px) * self.base_grid_px
        gy_end   = int(yw_max // self.base_grid_px) * self.base_grid_px

        x = gx_start
        while x <= gx_end:
            xs1, _ = self.world_to_screen(x, yw_min)
            xs2, _ = self.world_to_screen(x, yw_max)
            self.canvas.create_line(xs1, 0, xs2, self.CANVAS_HEIGHT, fill="#e0e0e0")
            x += self.base_grid_px

        y = gy_start
        while y <= gy_end:
            _, ys1 = self.world_to_screen(xw_min, y)
            _, ys2 = self.world_to_screen(xw_max, y)
            self.canvas.create_line(0, ys1, self.CANVAS_WIDTH, ys2, fill="#e0e0e0")
            y += self.base_grid_px

    # ===========================
    # 다각형 작도
    # ===========================
    def snap_to_grid_world(self, xw, yw):
        gx = round(xw / self.base_grid_px) * self.base_grid_px
        gy = round(yw / self.base_grid_px) * self.base_grid_px
        return gx, gy

    def on_click(self, event):
        if self.polygon_closed:
            return

        xw, yw = self.screen_to_world(event.x, event.y)
        sx, sy = self.snap_to_grid_world(xw, yw)

        if not self.points:
            self.points.append((sx, sy))
        else:
            last_x, last_y = self.points[-1]
            dx = sx - last_x
            dy = sy - last_y
            if abs(dx) >= abs(dy):
                nx, ny = sx, last_y
            else:
                nx, ny = last_x, sy

            first_x, first_y = self.points[0]
            if len(self.points) >= 2:
                dist_to_first = hypot(nx - first_x, ny - first_y)
                if dist_to_first <= self.CLOSE_WORLD_THRESHOLD:
                    nx, ny = first_x, first_y
                    self.points.append((nx, ny))
                    self.polygon_closed = True
                    self.redraw_all()
                    return

            self.points.append((nx, ny))

        self.redraw_all()

    def draw_polygon(self):
        if not self.points:
            return

        for i in range(len(self.points) - 1):
            x1w, y1w = self.points[i]
            x2w, y2w = self.points[i + 1]
            x1s, y1s = self.world_to_screen(x1w, y1w)
            x2s, y2s = self.world_to_screen(x2w, y2w)
            self.canvas.create_line(x1s, y1s, x2s, y2s, fill="black", width=2)

        for idx, (xw, yw) in enumerate(self.points):
            xs, ys = self.world_to_screen(xw, yw)
            r = 3
            color = "red" if idx == 0 else "blue"
            self.canvas.create_oval(xs - r, ys - r, xs + r, ys + r, fill=color, outline="")

        if self.polygon_closed and len(self.points) >= 3:
            coords = []
            for (xw, yw) in self.points:
                xs, ys = self.world_to_screen(xw, yw)
                coords.extend([xs, ys])
            self.polygon_fill_id = self.canvas.create_polygon(
                coords, fill="#cce5ff", outline="black", width=2
            )

    # ===========================
    # 포인트 인 폴리곤 / 내부 격자점 / 면적
    # ===========================
    def point_in_polygon_world(self, x, y, polygon):
        inside = False
        n = len(polygon)
        for i in range(n):
            x1, y1 = polygon[i]
            x2, y2 = polygon[(i + 1) % n]
            if ((y1 > y) != (y2 > y)):
                x_intersect = x1 + (y - y1) * (x2 - x1) / (y2 - y1)
                if x_intersect > x:
                    inside = not inside
        return inside

    def polygon_area(self, polygon):
        """Shoelace formula로 다각형 면적 계산 (px²)."""
        area = 0.0
        n = len(polygon)
        for i in range(n):
            x1, y1 = polygon[i]
            x2, y2 = polygon[(i + 1) % n]
            area += x1 * y2 - x2 * y1
        return abs(area) / 2.0

    def get_grid_points_inside_polygon_world(self, polygon):
        xs = [p[0] for p in polygon]
        ys = [p[1] for p in polygon]
        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)

        start_x = int((min_x // self.base_grid_px) * self.base_grid_px)
        start_y = int((min_y // self.base_grid_px) * self.base_grid_px)
        end_x   = int((max_x // self.base_grid_px) * self.base_grid_px)
        end_y   = int((max_y // self.base_grid_px) * self.base_grid_px)

        candidates = []
        for x in range(start_x, end_x + self.base_grid_px, self.base_grid_px):
            for y in range(start_y, end_y + self.base_grid_px, self.base_grid_px):
                if self.point_in_polygon_world(x, y, polygon):
                    candidates.append((x, y))
        return candidates

    def _dist_point_to_segment(self, px, py, x1, y1, x2, y2):
        # distance from point (px,py) to segment (x1,y1)-(x2,y2)
        dx = x2 - x1
        dy = y2 - y1
        if dx == 0 and dy == 0:
            return math.hypot(px - x1, py - y1)
        t = ((px - x1) * dx + (py - y1) * dy) / (dx * dx + dy * dy)
        t = max(0.0, min(1.0, t))
        proj_x = x1 + t * dx
        proj_y = y1 + t * dy
        return math.hypot(px - proj_x, py - proj_y)

    def distance_point_to_polygon_edges_world(self, px, py, polygon):
        # return minimum distance from point to polygon edges (world units)
        min_d = float('inf')
        n = len(polygon)
        for i in range(n):
            x1, y1 = polygon[i]
            x2, y2 = polygon[(i + 1) % n]
            d = self._dist_point_to_segment(px, py, x1, y1, x2, y2)
            if d < min_d:
                min_d = d
        return min_d

    def polygon_centroid_world(self, polygon):
        # area-weighted centroid (world units)
        area = 0.0
        cx = 0.0
        cy = 0.0
        n = len(polygon)
        for i in range(n):
            x0, y0 = polygon[i]
            x1, y1 = polygon[(i + 1) % n]
            a = x0 * y1 - x1 * y0
            area += a
            cx += (x0 + x1) * a
            cy += (y0 + y1) * a
        area *= 0.5
        if abs(area) < 1e-6:
            # fallback to arithmetic mean
            xs = [p[0] for p in polygon]
            ys = [p[1] for p in polygon]
            return sum(xs) / len(xs), sum(ys) / len(ys)
        cx /= (6.0 * area)
        cy /= (6.0 * area)
        return cx, cy

    def _find_nearest_unused_candidate(self, target, candidates, used_set):
        best_idx = None
        best_d2 = float('inf')
        for i, p in enumerate(candidates):
            if i in used_set:
                continue
            d2 = (p[0] - target[0]) ** 2 + (p[1] - target[1]) ** 2
            if d2 < best_d2:
                best_d2 = d2
                best_idx = i
        return best_idx

    def _generate_ideal_grid(self, polygon, spacing):
        # Generate a regular rectangular grid (world coords) covering polygon bbox
        xs = [p[0] for p in polygon]
        ys = [p[1] for p in polygon]
        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)

        # offset grid to center it within bbox
        nx = max(1, int(math.floor((max_x - min_x) / spacing)))
        ny = max(1, int(math.floor((max_y - min_y) / spacing)))
        if nx == 0 or ny == 0:
            return []

        used_spacing_x = (max_x - min_x) / nx
        used_spacing_y = (max_y - min_y) / ny
        sx = used_spacing_x
        sy = used_spacing_y

        points = []
        for i in range(nx + 1):
            cx = min_x + i * sx
            for j in range(ny + 1):
                cy = min_y + j * sy
                points.append((cx, cy))
        return points

    def _select_points_greedy_maxmin(self, candidates, n, min_edge_dist=None):
        # Greedy farthest-point sampling maximizing min distance between selected
        if n <= 0:
            return []
        if not candidates:
            return []
        # start with center-ish candidate
        candidates_sorted = sorted(candidates, key=lambda p: (p[0], p[1]))
        mid_idx = len(candidates_sorted) // 2
        selected = [candidates_sorted[mid_idx]]
        if n == 1:
            return selected
        remaining = candidates_sorted[:mid_idx] + candidates_sorted[mid_idx + 1:]
        while len(selected) < n and remaining:
            best_pt = None
            best_min_d = -1
            for pt in remaining:
                min_d = min(hypot(pt[0] - s[0], pt[1] - s[1]) for s in selected)
                if min_edge_dist is not None and min(hypot(pt[0] - s[0], pt[1] - s[1]) for s in selected) < 0:
                    pass
                if min_d > best_min_d:
                    best_min_d = min_d
                    best_pt = pt
            if best_pt is None:
                break
            selected.append(best_pt)
            remaining.remove(best_pt)
        return selected

    # ===========================
    # 점/사각형 (이전 코드와 동일)
    # ===========================
    def place_points_in_polygon(self):
        if not self.polygon_closed or len(self.points) < 3:
            messagebox.showinfo("알림", "먼저 도형을 완성해 주세요.")
            return
        n = simpledialog.askinteger("점 개수", "배치할 점의 개수를 입력하세요:",
                                    parent=self.root, minvalue=1)
        if not n:
            return
        # enforce even count
        if n % 2 != 0:
            # make it even by increasing by 1
            n += 1
            messagebox.showinfo("알림", f"점 개수는 짝수여야 합니다. 자동으로 {n}개로 조정합니다.")
        # clear previous
        for pid in self.placed_points_ids:
            self.canvas.delete(pid)
        self.placed_points_ids.clear()

        polygon = self.points[:-1]

        # parameters
        # world units: 1px = 1cm -> 0.8m = 80px
        min_edge_clearance_world = 0.8 * 100.0  # 0.8m -> 80cm -> 80px (1m=100px)
        # derive polygon area and ideal spacing
        poly_area = self.polygon_area(polygon)
        if poly_area <= 0:
            messagebox.showwarning("배치 실패", "다각형 면적이 유효하지 않습니다.")
            return

        # area per point -> ideal spacing
        area_per_point = poly_area / float(max(1, n))
        ideal_spacing = math.sqrt(area_per_point)

        # snap spacing to near multiples of base_grid_px but not smaller than base
        spacing = max(self.base_grid_px, round(ideal_spacing / self.base_grid_px) * self.base_grid_px)

        # attempt several spacing scales to find sufficient candidates
        candidates = []
        for scale in [1.0, 0.75, 0.5, 0.33, 0.25]:
            s = max(self.base_grid_px, spacing * scale)
            grid_pts = self._generate_ideal_grid(polygon, s)
            # keep only points inside polygon and sufficiently far from edges
            filtered = [p for p in grid_pts if self.point_in_polygon_world(p[0], p[1], polygon)
                        and self.distance_point_to_polygon_edges_world(p[0], p[1], polygon) >= min_edge_clearance_world]
            # dedupe
            pts_set = list(dict.fromkeys(filtered))
            if len(pts_set) >= n:
                candidates = pts_set
                break
            # accumulate candidates as fallback
            for p in pts_set:
                if p not in candidates:
                    candidates.append(p)

        # if still not enough candidates, expand to using the coarse grid inside polygon
        if len(candidates) < n:
            grid_coarse = self.get_grid_points_inside_polygon_world(polygon)
            for p in grid_coarse:
                if self.distance_point_to_polygon_edges_world(p[0], p[1], polygon) >= min_edge_clearance_world:
                    if p not in candidates:
                        candidates.append(p)

        if len(candidates) < n:
            messagebox.showwarning(
                "경고",
                f"유효한 배치 후보가 {len(candidates)}개입니다. {n}개 점을 배치할 수 없습니다."
            )

        # Prefer an r x c aligned grid (rows x cols) and snap ideal grid points to candidates
        selected = []
        if candidates:
            xs = [p[0] for p in candidates]
            ys = [p[1] for p in candidates]
            xmin, xmax = min(xs), max(xs)
            ymin, ymax = min(ys), max(ys)

            # choose r (rows) and c (cols) such that r*c >= n and ratio approximates bbox
            best_r, best_c = 1, n
            best_err = float('inf')
            for r_try in range(1, n + 1):
                c_try = int(math.ceil(n / r_try))
                ratio_grid = c_try / r_try
                ratio_room = (xmax - xmin) / (ymax - ymin + 1e-6)
                err = abs(ratio_grid - ratio_room)
                if r_try * c_try >= n and err < best_err:
                    best_err = err
                    best_r, best_c = r_try, c_try

            r, c = best_r, best_c

            # build ideal grid inside bounding box
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

            # snap ideal points to nearest candidates
            used_idx = set()
            for ip in ideal_points:
                best_i = None
                best_d2 = float('inf')
                for k, p in enumerate(candidates):
                    if k in used_idx:
                        continue
                    d2 = (p[0] - ip[0]) ** 2 + (p[1] - ip[1]) ** 2
                    if d2 < best_d2:
                        best_d2 = d2
                        best_i = k
                if best_i is not None:
                    used_idx.add(best_i)
                    selected.append(candidates[best_i])

        # if grid approach didn't yield enough, fallback to greedy selection
        if len(selected) < n:
            more = self._select_points_greedy_maxmin(candidates, n - len(selected))
            selected.extend(more)

        # final fallback: random sampling inside polygon ensuring min edge clearance
        attempts = 0
        max_attempts = 5000
        while len(selected) < n and attempts < max_attempts:
            attempts += 1
            xs_rand = random.uniform(min(p[0] for p in polygon), max(p[0] for p in polygon))
            ys_rand = random.uniform(min(p[1] for p in polygon), max(p[1] for p in polygon))
            if not self.point_in_polygon_world(xs_rand, ys_rand, polygon):
                continue
            if self.distance_point_to_polygon_edges_world(xs_rand, ys_rand, polygon) < min_edge_clearance_world:
                continue
            # ensure separation from already selected
            if any(hypot(xs_rand - s[0], ys_rand - s[1]) < min_edge_clearance_world for s in selected):
                continue
            selected.append((xs_rand, ys_rand))

        # draw selected
        for (xw, yw) in selected[:n]:
            xs, ys = self.world_to_screen(xw, yw)
            r = 4
            pid = self.canvas.create_oval(xs - r, ys - r, xs + r, ys + r,
                                          fill="green", outline="")
            self.placed_points_ids.append(pid)

    def select_far_points(self, candidates, n):
        if n <= 0:
            return []
        candidates_sorted = sorted(candidates, key=lambda p: (p[0], p[1]))
        mid_idx = len(candidates_sorted) // 2
        selected = [candidates_sorted[mid_idx]]
        if n == 1:
            return selected
        remaining = candidates_sorted[:mid_idx] + candidates_sorted[mid_idx + 1:]
        while len(selected) < n and remaining:
            best_point, best_min_dist = None, -1
            for pt in remaining:
                min_d = min(hypot(pt[0] - sp[0], pt[1] - sp[1]) for sp in selected)
                if min_d > best_min_dist:
                    best_min_dist = min_d
                    best_point = pt
            if best_point is None:
                break
            selected.append(best_point)
            remaining.remove(best_point)
        return selected

    def rectangle_fits_world(self, cx, cy, s, polygon):
        half = s / 2.0
        corners = [
            (cx - half, cy - half),
            (cx - half, cy + half),
            (cx + half, cy - half),
            (cx + half, cy + half),
        ]
        return all(self.point_in_polygon_world(x, y, polygon) for x, y in corners)

    def can_place_rectangles_world(self, candidates, polygon, N, s):
        valid = [p for p in candidates if self.rectangle_fits_world(p[0], p[1], s, polygon)]
        if len(valid) < N:
            return False, []
        selected = []
        valid_sorted = sorted(valid, key=lambda p: (p[0], p[1]))
        mid_idx = len(valid_sorted) // 2
        selected.append(valid_sorted[mid_idx])
        if N == 1:
            return True, selected
        remaining = valid_sorted[:mid_idx] + valid_sorted[mid_idx + 1:]
        while len(selected) < N and remaining:
            best_pt, best_min_dist = None, -1
            for p in remaining:
                min_d = min(hypot(p[0] - sp[0], p[1] - sp[1]) for sp in selected)
                if min_d < s:
                    continue
                if min_d > best_min_dist:
                    best_min_dist = min_d
                    best_pt = p
            if best_pt is None:
                return False, []
            selected.append(best_pt)
            remaining.remove(best_pt)
        if len(selected) < N:
            return False, []
        return True, selected

    def find_max_side_world(self, candidates, polygon, N):
        xs = [p[0] for p in polygon]
        ys = [p[1] for p in polygon]
        width = max(xs) - min(xs)
        height = max(ys) - min(ys)
        s_high = min(width, height)
        if s_high <= 0:
            return 0.0, []
        s_low = 0.0
        best_points = []
        for _ in range(25):
            s_mid = (s_low + s_high) / 2.0
            ok, pts = self.can_place_rectangles_world(candidates, polygon, N, s_mid)
            if ok:
                s_low = s_mid
                best_points = pts
            else:
                s_high = s_mid
        return s_low, best_points

    def place_rectangles_in_polygon(self):
        if not self.polygon_closed or len(self.points) < 3:
            messagebox.showinfo("알림", "먼저 도형을 완성해 주세요.")
            return
        N = simpledialog.askinteger("사각형 개수", "배치할 사각형 개수:",
                                    parent=self.root, minvalue=1)
        if not N:
            return
        for rid in self.rect_ids:
            self.canvas.delete(rid)
        self.rect_ids.clear()
        polygon = self.points[:-1]
        candidates = self.get_grid_points_inside_polygon_world(polygon)
        if len(candidates) < N:
            messagebox.showwarning(
                "경고",
                f"도형 내부 격자점 수가 {len(candidates)}개입니다.\n"
                f"{N}개의 사각형을 배치하기 충분하지 않을 수 없습니다."
            )
        s_world, centers = self.find_max_side_world(candidates, polygon, N)
        if not centers or s_world <= 0:
            messagebox.showwarning("배치 실패",
                                   "요청하신 개수의 사각형을 배치할 수 없습니다.")
            return

        half = s_world / 2.0
        for (xw, yw) in centers:
            x1w, y1w = xw - half, yw - half
            x2w, y2w = xw + half, yw + half
            x1s, y1s = self.world_to_screen(x1w, y1w)
            x2s, y2s = self.world_to_screen(x2w, y2w)
            rid = self.canvas.create_rectangle(x1s, y1s, x2s, y2s, outline="green", width=2)
            self.rect_ids.append(rid)

    # ===========================
    # 풍선 초기 배치
    # ===========================
    def start_balloons(self):
        if not self.polygon_closed or len(self.points) < 3:
            messagebox.showinfo("알림", "먼저 도형을 완성해 주세요.")
            return
        if self.balloon_running:
            messagebox.showinfo("알림", "이미 풍선 시뮬레이션이 진행 중입니다.")
            return

        N = simpledialog.askinteger("풍선 개수", "배치할 풍선(원) 개수:",
                                    parent=self.root, minvalue=1)
        if not N:
            return

        for bid in self.balloon_ids:
            self.canvas.delete(bid)
        self.balloon_ids.clear()
        self.balloons.clear()

        polygon = self.points[:-1]
        initial_r = self.base_grid_px / 2.0  # 반지름 10cm (지름 20cm)

        xs = [p[0] for p in polygon]
        ys = [p[1] for p in polygon]
        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)

        max_attempts = 3000
        attempts = 0
        while len(self.balloons) < N and attempts < max_attempts:
            attempts += 1
            cx = random.uniform(min_x, max_x)
            cy = random.uniform(min_y, max_y)
            if not self.point_in_polygon_world(cx, cy, polygon):
                continue
            if not self.circle_inside_polygon(cx, cy, initial_r, polygon):
                continue
            overlap = False
            for b in self.balloons:
                if hypot(cx - b["cx"], cy - b["cy"]) < (b["r"] + initial_r):
                    overlap = True
                    break
            if overlap:
                continue
            self.balloons.append(
                {"cx": cx, "cy": cy, "r": initial_r, "vx": 0.0, "vy": 0.0}
            )

        if not self.balloons:
            messagebox.showwarning(
                "배치 실패",
                "초기 풍선을 배치할 수 없습니다. 다각형이 너무 작거나 풍선 개수가 너무 많습니다."
            )
        else:
            self.draw_balloons()
            self.balloon_running = True
            self.root.after(30, self.step_balloons)

    def circle_inside_polygon(self, cx, cy, r, polygon):
        for (x, y) in [(cx + r, cy), (cx - r, cy), (cx, cy + r), (cx, cy - r)]:
            if not self.point_in_polygon_world(x, y, polygon):
                return False
        return True

    def draw_balloons(self):
        for bid in self.balloon_ids:
            self.canvas.delete(bid)
        self.balloon_ids.clear()
        for b in self.balloons:
            cx, cy, r = b["cx"], b["cy"], b["r"]
            xs, ys = self.world_to_screen(cx, cy)
            rs = r * self.scale_factor
            bid = self.canvas.create_oval(xs - rs, ys - rs, xs + rs, ys + rs,
                                          outline="red", width=2)
            self.balloon_ids.append(bid)

    # ===========================
    # 애니메이션용 step_balloons
    # ===========================
    def step_balloons(self):
        self._single_growth_step(animated=True)

    # ===========================
    # 풍선완료: 애니메이션 없이 빠르게 성장 끝까지 진행
    # ===========================
    def fast_forward_growth(self):
        if not self.polygon_closed or len(self.points) < 3 or not self.balloons:
            messagebox.showinfo("알림", "먼저 도형을 완성하고, 풍선을 배치해 주세요.")
            return
        # 애니메이션 중이라면 중단
        self.balloon_running = False

        max_steps = 200          # 최대 성장 스텝 수
        growth_epsilon = 1e-3    # 성장 멈춤 기준

        for _ in range(max_steps):
            max_change = self._single_growth_step(animated=False)
            if max_change < growth_epsilon:
                break

        # 최종 상태 그리기
        self.draw_balloons()
        messagebox.showinfo("성장 완료", "풍선 성장을 빠르게 완료했습니다.")

    # ===========================
    # 정렬하기: 풍선 중심을 가로/세로 격자로 재배치
    # ===========================
    def align_balloons_grid(self):
        if not self.polygon_closed or len(self.points) < 3 or not self.balloons:
            messagebox.showinfo("알림", "먼저 도형을 완성하고, 풍선을 배치해 주세요.")
            return

        polygon = self.points[:-1]
        xs = [p[0] for p in polygon]
        ys = [p[1] for p in polygon]
        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)

        n = len(self.balloons)
        cols = math.ceil(math.sqrt(n))
        rows = math.ceil(n / cols)

        # 바운딩 박스 안에 균일한 격자 생성
        # 약간 여유를 두고 안쪽으로
        margin_x = (max_x - min_x) * 0.05
        margin_y = (max_y - min_y) * 0.05
        gx_min = min_x + margin_x
        gx_max = max_x - margin_x
        gy_min = min_y + margin_y
        gy_max = max_y - margin_y

        if cols > 1:
            dx = (gx_max - gx_min) / (cols - 1)
        else:
            dx = 0
        if rows > 1:
            dy = (gy_max - gy_min) / (rows - 1)
        else:
            dy = 0

        candidates = []
        for r in range(rows):
            cy = gy_min + dy * r
            for c in range(cols):
                cx = gx_min + dx * c
                if self.point_in_polygon_world(cx, cy, polygon):
                    candidates.append((cx, cy))

        if len(candidates) < n:
            messagebox.showwarning(
                "정렬 실패",
                "다각형 내부의 정렬 격자점 수가 풍선 개수보다 적습니다.\n"
                "다각형이 너무 복잡하거나 좁을 수 있습니다."
            )
            return

        # 그냥 앞에서부터 n개 사용
        for i, b in enumerate(self.balloons):
            b["cx"], b["cy"] = candidates[i]

        self.draw_balloons()

    # ===========================
    # 공통 성장 스텝
    # ===========================
    def _single_growth_step(self, animated: bool) -> float:
        """한 번의 성장/물리 스텝을 수행.
        animated=True 이면 30ms 후 다시 예약, False면 예약 없이 한 번만.
        반환값: 이번 스텝에서의 최대 반지름 증가량.
        """
        if not self.polygon_closed or len(self.points) < 3:
            self.balloon_running = False
            return 0.0

        polygon = self.points[:-1]

        # 다각형 면적 및 풍선수에 따른 최대 허용 면적/반지름 계산
        poly_area = self.polygon_area(polygon)  # px²
        N = max(len(self.balloons), 1)
        target_area_each = poly_area / N * 0.785  # 조건: 다각형면적/풍선수 * 0.785

        # 1) 면적 기준 반지름 증가 (최대 허용 면적을 넘지 않도록 clamp)
        frame_dt = 0.03
        dA_nominal = self.balloon_area_rate * frame_dt

        max_r = max(self.CANVAS_WIDTH, self.CANVAS_HEIGHT)
        max_radius_change = 0.0

        for b in self.balloons:
            r_old = b["r"]
            if r_old <= 0:
                continue

            current_area = math.pi * (r_old ** 2)
            if current_area >= target_area_each:
                continue  # 이미 상한 도달

            dA_limit = target_area_each - current_area
            dA = min(dA_nominal, dA_limit)

            dr = dA / (2.0 * math.pi * r_old + 1e-6)
            r_new = min(r_old + dr, max_r)
            b["r"] = r_new

            delta_r = abs(r_new - r_old)
            if delta_r > max_radius_change:
                max_radius_change = delta_r

        # 2) 풍선-풍선 상호작용 (위치 보정 + 속도 반영)
        n = len(self.balloons)
        for i in range(n):
            for j in range(i + 1, n):
                bi = self.balloons[i]
                bj = self.balloons[j]
                dx = bj["cx"] - bi["cx"]
                dy = bj["cy"] - bi["cy"]
                dist = (dx*dx + dy*dy) ** 0.5
                if dist <= 1e-6:
                    continue
                target = bi["r"] + bj["r"]
                overlap = target - dist
                if overlap > 0:
                    nx = dx / dist
                    ny = dy / dist

                    mi = bi["r"] * bi["r"] + 1e-3
                    mj = bj["r"] * bj["r"] + 1e-3
                    inv_mi = 1.0 / mi
                    inv_mj = 1.0 / mj
                    total_inv_m = inv_mi + inv_mj
                    if total_inv_m == 0:
                        continue

                    move_i = inv_mi / total_inv_m
                    move_j = inv_mj / total_inv_m

                    corr = overlap * 1.0
                    bi["cx"] -= nx * corr * move_i
                    bi["cy"] -= ny * corr * move_i
                    bj["cx"] += nx * corr * move_j
                    bj["cy"] += ny * corr * move_j

                    # 속도 보정 (약한 반발)
                    rel_vx = bj["vx"] - bi["vx"]
                    rel_vy = bj["vy"] - bi["vy"]
                    rel_normal = rel_vx * nx + rel_vy * ny
                    e = 0.2
                    j_impulse = -(1 + e) * rel_normal / (inv_mi + inv_mj)
                    jx = j_impulse * nx
                    jy = j_impulse * ny
                    bi["vx"] -= jx * inv_mi
                    bi["vy"] -= jy * inv_mi
                    bj["vx"] += jx * inv_mj
                    bj["vy"] += jy * inv_mj

        # 3) 위치 업데이트
        for b in self.balloons:
            b["cx"] += b["vx"]
            b["cy"] += b["vy"]

        # 4) 벽 충돌: 위치 보정 + 속도 반사(1차)
        xs_poly = [p[0] for p in polygon]
        ys_poly = [p[1] for p in polygon]
        poly_min_x, poly_max_x = min(xs_poly), max(xs_poly)
        poly_min_y, poly_max_y = min(ys_poly), max(ys_poly)

        for b in self.balloons:
            cx, cy, r = b["cx"], b["cy"], b["r"]
            for i in range(len(polygon)):
                x1, y1 = polygon[i]
                x2, y2 = polygon[(i + 1) % len(polygon)]
                if x1 == x2:  # 수직 벽
                    if min(y1, y2) <= cy <= max(y1, y2):
                        d = cx - x1
                        if abs(d) < r:
                            sign = 1 if d > 0 else -1
                            b["cx"] = x1 + sign * r
                            b["vx"] *= -0.5
                elif y1 == y2:  # 수평 벽
                    if min(x1, x2) <= cx <= max(x1, x2):
                        d = cy - y1
                        if abs(d) < r:
                            sign = 1 if d > 0 else -1
                            b["cy"] = y1 + sign * r
                            b["vy"] *= -0.5

        # 5) 감쇠
        damping = 0.98
        for b in self.balloons:
            b["vx"] *= damping
            b["vy"] *= damping

        # 6) 최종 분리 단계(풍선-풍선)
        n = len(self.balloons)
        for _ in range(2):  # 두 번 반복해 안정화
            for i in range(n):
                for j in range(i + 1, n):
                    bi = self.balloons[i]
                    bj = self.balloons[j]
                    dx = bj["cx"] - bi["cx"]
                    dy = bj["cy"] - bi["cy"]
                    dist = (dx*dx + dy*dy) ** 0.5
                    if dist <= 1e-6:
                        dist = 1e-3
                        dx = dist
                        dy = 0.0
                    target = bi["r"] + bj["r"]
                    overlap = target - dist
                    if overlap > 0:
                        nx = dx / dist
                        ny = dy / dist
                        mi = bi["r"] * bi["r"] + 1e-3
                        mj = bj["r"] * bj["r"] + 1e-3
                        inv_mi = 1.0 / mi
                        inv_mj = 1.0 / mj
                        total_inv_m = inv_mi + inv_mj
                        if total_inv_m == 0:
                            continue
                        move_i = inv_mi / total_inv_m
                        move_j = inv_mj / total_inv_m
                        corr = overlap
                        bi["cx"] -= nx * corr * move_i
                        bi["cy"] -= ny * corr * move_i
                        bj["cx"] += nx * corr * move_j
                        bj["cy"] += ny * corr * move_j

        # 7) 최종 벽 보정(강제)
        for b in self.balloons:
            r = b["r"]
            min_cx = poly_min_x + r
            max_cx = poly_max_x - r
            min_cy = poly_min_y + r
            max_cy = poly_max_y - r
            b["cx"] = min(max(b["cx"], min_cx), max_cx)
            b["cy"] = min(max(b["cy"], min_cy), max_cy)

        # 8) 너무 작아진 풍선 제거
        self.balloons = [b for b in self.balloons if b["r"] > 1.0]

        # 9) 애니메이션 모드일 때만 다시 그리기 및 예약
        if animated:
            self.draw_balloons()
            growth_epsilon = 1e-3
            if self.balloons and max_radius_change < growth_epsilon:
                self.balloon_running = False
                messagebox.showinfo("성장 완료", "모든 풍선의 성장이 완료되었습니다.")
            elif self.balloons:
                self.root.after(30, self.step_balloons)
            else:
                self.balloon_running = False

        return max_radius_change


if __name__ == "__main__":
    root = tk.Tk()
    app = OrthogonalPolygonDrawer(root)
    root.mainloop()
