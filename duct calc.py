import tkinter as tk
from tkinter import messagebox, ttk, simpledialog
import math

# =========================
# 계산 함수들
# =========================

def calc_circular_diameter(q_m3h: float, dp_mmAq_per_m: float) -> float:
    """등가 원형 덕트 직경 D1 (mm) 계산"""
    if q_m3h <= 0:
        raise ValueError("풍량(m³/h)은 0보다 커야 합니다.")
    if dp_mmAq_per_m <= 0:
        raise ValueError("정압값(mmAq/m)은 0보다 커야 합니다.")

    C = 3.295e-10  # 경험식 상수
    D = ((C * q_m3h**1.9 / dp_mmAq_per_m)**0.199) * 1000  # mm
    return round(D, 0)


def round_step_up(x: float, step: float = 50) -> float:
    return math.ceil(x / step) * step


def round_step_down(x: float, step: float = 50) -> float:
    return math.floor(x / step) * step


def rect_equiv_diameter(a_mm: float, b_mm: float) -> float:
    """사각 덕트 a,b(mm)에 대한 등가 원형 직경 De(mm) (ASHRAE)"""
    if a_mm <= 0 or b_mm <= 0:
        raise ValueError("사각 덕트 변은 0보다 커야 합니다.")
    a, b = float(a_mm), float(b_mm)
    return 1.30 * (a*b)**0.625 / (a + b)**0.25


def size_rect_from_D1(D1: float, aspect_ratio: float, step: float = 50):
    """
    1번(등가원형 D1)을 기준으로:
      - 4번: 이론 사각 (조정 전) [항상 큰값, 작은값 순서]
      - 3번: 4번 기반 50mm 조정 (규칙 적용) [항상 큰값, 작은값 순서]
    """
    if D1 <= 0:
        raise ValueError("원형 덕트 직경은 0보다 커야 합니다.")
    if aspect_ratio <= 0:
        raise ValueError("종횡비(b/a)는 0보다 커야 합니다.")

    De_target = float(D1)
    r = float(aspect_ratio)

    # --- 4번: 이론 사각 (조정 전) ---
    a_theo = De_target * (1 + r)**0.25 / (1.30 * r**0.625)
    b_theo = r * a_theo
    theo_big, theo_small = max(a_theo, b_theo), min(a_theo, b_theo)

    # --- 후보1: 작은 값 올림, 큰 값 내림 ---
    small_up = round_step_up(theo_small, step)
    big_down = max(round_step_down(theo_big, step), step)
    De1 = rect_equiv_diameter(small_up, big_down)

    # --- 후보2: 둘 다 올림 ---
    a_up = round_step_up(a_theo, step)
    b_up = round_step_up(b_theo, step)
    De2 = rect_equiv_diameter(a_up, b_up)

    # --- 최종 선택 (3번) ---
    if De1 >= De_target:
        sel_big, sel_small = max(small_up, big_down), min(small_up, big_down)
        De_sel = De1
    else:
        sel_big, sel_small = max(a_up, b_up), min(a_up, b_up)
        De_sel = De2

    return (
        int(round(sel_big)),
        int(round(sel_small)),
        round(De_sel, 1),
        round(theo_big, 1),
        round(theo_small, 1),
    )

# =========================
# 팔레트(Canvas) 관련 (모델 좌표 기반)
# =========================

GRID_STEP_MODEL = 0.5
INITIAL_SCALE = 40.0

class AirPoint:
    def __init__(self, mx, my, kind, flow):
        self.mx = mx  # model x (m)
        self.my = my  # model y (m)
        self.kind = kind  # "inlet" or "outlet"
        self.flow = flow  # m3/h
        self.canvas_id = None
        self.text_id = None


class DuctSegment:
    """종합 사이징으로 생성되는 덕트 구간 데이터(모델 좌표 기준)"""
    def __init__(self, mx1, my1, mx2, my2, label_text, duct_w_mm, duct_h_mm,
                 flow_m3h, vertical_only=False):
        self.mx1 = mx1
        self.my1 = my1
        self.mx2 = mx2
        self.my2 = my2
        self.label_text = label_text
        self.duct_w_mm = duct_w_mm
        self.duct_h_mm = duct_h_mm
        self.flow = flow_m3h
        self.vertical_only = vertical_only   # True: 순수 수직, False: 순수 수평
        self.line_ids = []
        self.text_id = None
        self.leader_id = None

        # 상호작용 상태
        self.is_hovered = False
        self.is_dragging = False
        self.drag_start_model = None  # (mx, my) 드래그 시작점(모델좌표)

    def length_m(self):
        dx = self.mx2 - self.mx1
        dy = self.my2 - self.my1
        return math.sqrt(dx*dx + dy*dy)


class Palette:
    def __init__(self, parent):
        self.canvas = tk.Canvas(parent, bg="white")
        self.canvas.pack(expand=True, fill="both", padx=10, pady=10)

        self.points = []
        self.inlet_flow = 0.0

        self.scale_factor = INITIAL_SCALE
        self.offset_x = 0.0
        self.offset_y = 0.0

        self.grid_tag = "grid"
        self.pan_start_screen = None

        self.segments = []

        # Pencil drawing state
        self.mode = "pan"  # "pan" or "pencil"
        self._drawing = False
        self._draw_start = None
        self._preview_line_id = None

        # 라인 hover/drag 상태
        self.hovered_segment = None
        self.dragging_segment = None
        # hovered text id for tooltip bolding
        self.hovered_text_id = None

        # (추가) 마우스 커서 위치 추적용 (실시간 계산 툴팁 표시)
        self.current_mouse_model = None

        # (추가) points 변경 시 외부로 알리는 콜백 함수 저장용
        self.points_changed_callback = None

        self.canvas.bind("<Button-1>", self.on_left_click)
        self.canvas.bind("<Button-3>", self.on_right_click)
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        self.canvas.bind("<Button-2>", self.on_middle_press)
        self.canvas.bind("<B2-Motion>", self.on_middle_drag)
        self.canvas.bind("<Configure>", self.on_resize)

        # 마우스 이동 및 왼쪽 드래그
        self.canvas.bind("<Motion>", self.on_mouse_move)
        self.canvas.bind("<Leave>", self.on_mouse_leave)  # 마우스 나감 처리
        self.canvas.bind("<B1-Motion>", self.on_left_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_release)

        self.redraw_all()

    # ---------- (추가) points 변경 알림 ----------

    def _notify_points_changed(self):
        """점이 추가되거나 삭제되었을 때 등록된 콜백 함수를 호출"""
        cb = getattr(self, "points_changed_callback", None)
        if callable(cb):
            try:
                cb(self)
            except Exception:
                pass

    # ---------- 모드 전환 ----------

    def set_mode_pencil(self):
        # Toggle: if currently pencil, switch back to pan
        if self.mode == "pencil":
            self.set_mode_pan()
            return
        self.mode = "pencil"
        try:
            self.canvas.config(cursor="pencil")
        except tk.TclError:
            self.canvas.config(cursor="crosshair")

    def set_mode_pan(self):
        self.mode = "pan"
        self.canvas.config(cursor="arrow")

    # ---------- 좌표 변환 ----------

    def model_to_screen(self, mx, my):
        sx = mx * self.scale_factor + self.offset_x
        sy = my * self.scale_factor + self.offset_y
        return sx, sy

    def screen_to_model(self, sx, sy):
        mx = (sx - self.offset_x) / self.scale_factor
        my = (sy - self.offset_y) / self.scale_factor
        return mx, my

    # ---------- 격자/그리기 ----------

    def draw_grid(self):
        self.canvas.delete(self.grid_tag)

        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w <= 0 or h <= 0:
            return

        mx_min, my_min = self.screen_to_model(0, 0)
        mx_max, my_max = self.screen_to_model(w, h)

        x_start = math.floor(mx_min / GRID_STEP_MODEL) * GRID_STEP_MODEL
        x_end = math.ceil(mx_max / GRID_STEP_MODEL) * GRID_STEP_MODEL
        y_start = math.floor(my_min / GRID_STEP_MODEL) * GRID_STEP_MODEL
        y_end = math.ceil(my_max / GRID_STEP_MODEL) * GRID_STEP_MODEL

        x = x_start
        while x <= x_end:
            sx1, sy1 = self.model_to_screen(x, y_start)
            sx2, sy2 = self.model_to_screen(x, y_end)
            self.canvas.create_line(
                sx1, sy1, sx2, sy2,
                fill="#e0e0e0",
                tags=self.grid_tag
            )
            x += GRID_STEP_MODEL

        y = y_start
        while y <= y_end:
            sx1, sy1 = self.model_to_screen(x_start, y)
            sx2, sy2 = self.model_to_screen(x_end, y)
            self.canvas.create_line(
                sx1, sy1, sx2, sy2,
                fill="#e0e0e0",
                tags=self.grid_tag
            )
            y += GRID_STEP_MODEL

    def redraw_all(self):
        self.canvas.delete("all")
        self.draw_grid()

        # 점 + 풍량 라벨
        for p in self.points:
            sx, sy = self.model_to_screen(p.mx, p.my)
            color = "red" if p.kind == "inlet" else "blue"
            r = 5
            p.canvas_id = self.canvas.create_oval(
                sx - r, sy - r, sx + r, sy + r,
                fill=color, outline=""
            )
            label = f"{p.flow:.1f}"
            p.text_id = self.canvas.create_text(
                sx + 10, sy - 10,
                text=label,
                fill="black",
                font=("Arial", 8)
            )

        # 덕트 구간
        for seg in self.segments:
            seg.line_ids.clear()
            seg.text_id = None
            seg.leader_id = None

            mx1, my1, mx2, my2 = seg.mx1, seg.my1, seg.mx2, seg.my2
            sx1, sy1 = self.model_to_screen(mx1, my1)
            sx2, sy2 = self.model_to_screen(mx2, my2)

            # hover 시 굵게
            line_width = 3 if seg.is_hovered else 1
            line_color = "gray50"

            if seg.vertical_only:
                seg.line_ids.append(
                    self.canvas.create_line(
                        sx1, sy1, sx2, sy2,
                        fill=line_color,
                        width=line_width,
                        tags=("duct_line",)
                    )
                )
                horizontal_len = 0.0
                vertical_len = abs(my2 - my1)
            else:
                seg.line_ids.append(
                    self.canvas.create_line(
                        sx1, sy1, sx2, sy2,
                        fill=line_color,
                        width=line_width,
                        tags=("duct_line",)
                    )
                )
                horizontal_len = abs(mx2 - mx1)
                vertical_len = 0.0

            # 기준축 선택
            if vertical_len > horizontal_len:
                use_vertical = True
            else:
                use_vertical = False

            leader_length_px = 15
            text_offset_px = 5

            if use_vertical:
                # 세로 기준: 중앙점에서 수평 지시선
                mid_mx_v = mx1
                mid_my_v = (my1 + my2) / 2.0
                vx, vy = self.model_to_screen(mid_mx_v, mid_my_v)

                seg.leader_id = self.canvas.create_line(
                    vx, vy,
                    vx + leader_length_px, vy,
                    fill="blue"
                )

                tx = vx + leader_length_px + text_offset_px
                ty = vy

                seg.text_id = self.canvas.create_text(
                    tx, ty,
                    text=seg.label_text,
                    fill="blue",
                    font=("Arial", 8),
                    anchor="w"
                )
            else:
                # 가로 기준: 중앙점에서 수직 지시선
                mid_mx_h = (mx1 + mx2) / 2.0
                mid_my_h = my1
                hx, hy = self.model_to_screen(mid_mx_h, mid_my_h)

                seg.leader_id = self.canvas.create_line(
                    hx, hy,
                    hx, hy - leader_length_px,
                    fill="blue"
                )

                tx = hx
                ty = hy - leader_length_px - text_offset_px

                seg.text_id = self.canvas.create_text(
                    tx, ty,
                    text=seg.label_text,
                    fill="blue",
                    font=("Arial", 8),
                    anchor="s"
                )

        # (추가) 마우스 커서를 따라다니는 계산 결과 툴팁 그리기
        if self.current_mouse_model is not None:
            cur_mx, cur_my = self.current_mouse_model
            outlets = [p for p in self.points if p.kind == "outlet"]
            
            if outlets:
                sum_dx = 0.0
                sum_dy = 0.0
                sum_abs_dx = 0.0
                sum_abs_dy = 0.0
                sum_sq_dx = 0.0
                sum_sq_dy = 0.0

                for p in outlets:
                    dx = cur_mx - p.mx  # 마우스.x - outlet.x
                    dy = cur_my - p.my  # 마우스.y - outlet.y
                    sum_dx += dx
                    sum_dy += dy
                    sum_abs_dx += abs(dx)
                    sum_abs_dy += abs(dy)
                    sum_sq_dx += dx**2
                    sum_sq_dy += dy**2

                tooltip_text = (
                    f"Σ(Cursor.x - Outlet.x) : {sum_dx:.2f} m\n"
                    f"Σ(Cursor.y - Outlet.y) : {sum_dy:.2f} m\n"
                    f"Σ(|Cursor.x - Outlet.x|) : {sum_abs_dx:.2f} m\n"
                    f"Σ(|Cursor.y - Outlet.y|) : {sum_abs_dy:.2f} m\n"
                    f"Σ(Cursor.x - Outlet.x)²: {sum_sq_dx:.2f} m²\n"
                    f"Σ(Cursor.y - Outlet.y)²: {sum_sq_dy:.2f} m²"
                )

                # 마우스 화면 좌표
                msx, msy = self.model_to_screen(cur_mx, cur_my)

                # 텍스트 오프셋
                tx, ty = msx + 20, msy + 20

                # 배경 박스 크기 확장 (텍스트 줄 수 증가)
                box_w = 260
                box_h = 100
                self.canvas.create_rectangle(
                    tx - 5, ty - 5, tx + box_w, ty + box_h,
                    fill="#ffffe0", outline="black"
                )

                self.canvas.create_text(
                    tx, ty,
                    text=tooltip_text,
                    anchor="nw",
                    font=("Consolas", 9),
                    fill="black"
                )

    def on_resize(self, event):
        self.redraw_all()

    # ---------- 스냅 ----------

    def snap_model(self, mx, my):
        smx = round(mx / GRID_STEP_MODEL) * GRID_STEP_MODEL
        smy = round(my / GRID_STEP_MODEL) * GRID_STEP_MODEL
        return smx, smy

    # ---------- 팬 ----------

    def on_middle_press(self, event):
        self.pan_start_screen = (event.x, event.y)

    def on_middle_drag(self, event):
        if self.pan_start_screen is None:
            return
        sx0, sy0 = self.pan_start_screen
        dx = event.x - sx0
        dy = event.y - sy0
        self.pan_start_screen = (event.x, event.y)
        self.offset_x += dx
        self.offset_y += dy
        self.redraw_all()

    # ---------- 줌 ----------

    def on_mousewheel(self, event):
        factor = 1.1 if event.delta > 0 else 0.9
        new_scale = self.scale_factor * factor
        if not (10.0 <= new_scale <= 400.0):
            return

        sx, sy = event.x, event.y
        mx, my = self.screen_to_model(sx, sy)

        self.scale_factor = new_scale
        self.offset_x = sx - mx * self.scale_factor
        self.offset_y = sy - my * self.scale_factor

        self.redraw_all()

    # ---------- 덕트 라인 히트 테스트 ----------

    def _hit_test_segment(self, mx, my, tol=0.1):
        """
        mx,my (모델좌표)가 어떤 DuctSegment 위에 있는지 판정.
        tol: 덕트 중심선에서의 허용 거리 (모델 좌표, m).
        수직 덕트면 x만, 수평 덕트면 y만 비교.
        """
        for seg in self.segments:
            if seg.vertical_only:
                # x = 상수, y 범위
                if min(seg.my1, seg.my2) - tol <= my <= max(seg.my1, seg.my2) + tol:
                    if abs(mx - seg.mx1) <= tol:
                        return seg
            else:
                # y = 상수, x 범위
                if min(seg.mx1, seg.mx2) - tol <= mx <= max(seg.mx1, seg.mx2) + tol:
                    if abs(my - seg.my1) <= tol:
                        return seg
        return None

    # ---------- 점/풍량 ----------

    def set_inlet_flow(self, flow):
        self.inlet_flow = float(flow)
        if self.points:
            p0 = self.points[0]
            if p0.kind == "inlet":
                p0.flow = self.inlet_flow
        self.redraw_all()
        self._notify_points_changed()

    def on_left_click(self, event):
        # Pencil 모드: 자유 라인 그리기 시작
        if self.mode == "pencil":
            mx, my = self.screen_to_model(event.x, event.y)
            smx, smy = self.snap_model(mx, my)
            self._draw_start = (smx, smy)
            self._drawing = True
            # 미리보기 라인 초기화
            if self._preview_line_id is not None:
                self.canvas.delete(self._preview_line_id)
                self._preview_line_id = None
            sx, sy = self.model_to_screen(smx, smy)
            self._preview_line_id = self.canvas.create_line(sx, sy, sx, sy, fill="black", width=2, dash=(4, 2))
            return

        # 먼저, 덕트 위인지 확인: 라인 위라면 점 생성 안 함
        mx, my = self.screen_to_model(event.x, event.y)
        seg = self._hit_test_segment(mx, my, tol=0.15)
        if seg is not None:
            return  # 라인 위 클릭은 점 추가하지 않음

        sx, sy = event.x, event.y
        mx, my = self.screen_to_model(sx, sy)
        mx, my = self.snap_model(mx, my)

        if not self.points:
            flow = self.inlet_flow if self.inlet_flow > 0 else 0.0
            p = AirPoint(mx, my, "inlet", flow)
            self.points.append(p)
        else:
            p = AirPoint(mx, my, "outlet", 0.0)
            self.points.append(p)

        self.segments.clear()
        self.redraw_all()
        # (추가) 점 추가 시 합계 갱신
        self._notify_points_changed()

    def on_right_click(self, event):
        sx, sy = event.x, event.y
        mx, my = self.screen_to_model(sx, sy)

        target = self._find_point_near_model(mx, my, tol_model=0.3)
        if target is None:
            return

        if target.kind == "inlet":
            messagebox.showinfo("정보", "Air inlet의 풍량은 좌측 텍스트 값을 사용합니다.")
            return

        remaining = self._calc_remaining_flow(exclude=target)
        current = target.flow

        msg = (
            f"Air inlet 풍량: {self.inlet_flow:.1f} m³/h\n"
            f"다른 outlet에 분배된 풍량 합계: {self._sum_outlet_flow(exclude=target):.1f} m³/h\n"
            f"남은 풍량(참고용): {remaining:.1f} m³/h\n"
            f"현 지점(outlet) 현재 풍량: {current:.1f} m³/h\n\n"
            f"이 outlet의 풍량을 입력하세요:"
        )

        answer = simpledialog.askstring("Air outlet 풍량 입력", msg)
        if answer is None:
            return

        try:
            new_flow = float(answer)
            if new_flow < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("입력 오류", "0 이상 숫자로 입력해주세요.")
            return

        target.flow = new_flow
        self.segments.clear()
        self.redraw_all()
        # (추가) 풍량 변경시 호출 (위치 변경은 없지만 갱신)
        self._notify_points_changed()

    def _find_point_near_model(self, mx, my, tol_model=0.3):
        for p in self.points:
            if (p.mx - mx)**2 + (p.my - my)**2 <= tol_model**2:
                return p
        return None

    def _sum_outlet_flow(self, exclude=None):
        s = 0.0
        for p in self.points:
            if p.kind == "outlet" and p is not exclude:
                s += p.flow
        return s

    def _calc_remaining_flow(self, exclude=None):
        return self.inlet_flow - self._sum_outlet_flow(exclude=exclude)

    # ---------- 마우스 이동 / 드래그 ----------

    def on_mouse_move(self, event):
        # 모델 좌표로 변환 및 저장 (실시간 툴팁용)
        mx, my = self.screen_to_model(event.x, event.y)
        self.current_mouse_model = (mx, my)

        # 이미 드래그 중이면 hover는 굳이 다시 안 바꿔도 됨
        if self.dragging_segment is not None:
            self.redraw_all() # 드래그 중에도 툴팁 갱신을 위해 호출
            return

        seg = self._hit_test_segment(mx, my, tol=0.15)

        # 기존 hover 해제
        if self.hovered_segment is not None and self.hovered_segment is not seg:
            self.hovered_segment.is_hovered = False
            self.hovered_segment = None

        # 새 hover 설정
        if seg is not None:
            seg.is_hovered = True
            self.hovered_segment = seg

        # 텍스트 hover 처리: point 텍스트(풍량) 위에 커서가 있으면 굵게
        hovered_text = None
        for p in self.points:
            if p.text_id is None:
                continue
            try:
                tx, ty = self.model_to_screen(p.mx, p.my)
                # canvas text is at (tx+10, ty-10)
                sx = tx + 10
                sy = ty - 10
            except Exception:
                continue
            if (event.x - sx)**2 + (event.y - sy)**2 <= 64:  # within 8px
                hovered_text = p.text_id
                break

        # 변경 적용
        if hovered_text != self.hovered_text_id:
            # 이전 굵게 해제
            if self.hovered_text_id is not None:
                try:
                    self.canvas.itemconfigure(self.hovered_text_id, font=("Arial", 8))
                except Exception:
                    pass
            self.hovered_text_id = hovered_text
            if self.hovered_text_id is not None:
                try:
                    self.canvas.itemconfigure(self.hovered_text_id, font=("Arial", 8, "bold"))
                except Exception:
                    pass

        self.redraw_all()

    def on_mouse_leave(self, event):
        """마우스가 캔버스를 벗어나면 툴팁을 숨김"""
        self.current_mouse_model = None
        self.redraw_all()

    def on_left_drag(self, event):
        # Pencil 모드: 미리보기 라인 업데이트 (격자 스냅, 수평/수직 우선)
        if self.mode == "pencil" and self._drawing and self._draw_start is not None:
            mx, my = self.screen_to_model(event.x, event.y)
            smx, smy = self.snap_model(mx, my)
            sx1, sy1 = self.model_to_screen(*self._draw_start)
            sx2, sy2 = self.model_to_screen(smx, smy)
            # 수평/수직 정렬: 더 큰 이동축을 우선으로 고정
            dx = abs(smx - self._draw_start[0])
            dy = abs(smy - self._draw_start[1])
            if dx >= dy:
                sy2 = sy1
            else:
                sx2 = sx1
            if self._preview_line_id is not None:
                self.canvas.coords(self._preview_line_id, sx1, sy1, sx2, sy2)
            return

        mx, my = self.screen_to_model(event.x, event.y)
        # 드래그 중에도 마우스 위치 업데이트 (툴팁용)
        self.current_mouse_model = (mx, my)

        # 드래그 시작
        if self.dragging_segment is None:
            seg = self._hit_test_segment(mx, my, tol=0.15)
            if seg is None:
                return
            self.dragging_segment = seg
            seg.is_dragging = True
            seg.drag_start_model = (mx, my)

        seg = self.dragging_segment
        if seg is None:
            return

        cur_mx, cur_my = mx, my
        start_mx, start_my = seg.drag_start_model

        # 덕트 방향에 따라 한 축만 이동 (격자 스냅 포함)
        if seg.vertical_only:
            # 수직 세그먼트: x만 이동
            dx_raw = cur_mx - start_mx
            base_x = (seg.mx1 + seg.mx2) / 2.0
            new_x = base_x + dx_raw
            # 격자 스냅
            snapped_x, _ = self.snap_model(new_x, 0.0)
            dx = snapped_x - base_x
            if abs(dx) < 1e-9:
                return

            # 연결된 세그먼트 전체를 x 방향으로 이동
            self._move_connected_segments(seg, dx=dx, dy=0.0)
            # 이동 후 직교 정렬
            self._orthogonalize_segments()
            # inlet 연결 보정
            self._ensure_inlet_connected()

            seg.drag_start_model = (cur_mx, start_my)

        else:
            # 수평 세그먼트: y만 이동
            dy_raw = cur_my - start_my
            base_y = (seg.my1 + seg.my2) / 2.0
            new_y = base_y + dy_raw
            # 격자 스냅
            _, snapped_y = self.snap_model(0.0, new_y)
            dy = snapped_y - base_y
            if abs(dy) < 1e-9:
                return

            # 연결된 세그먼트 전체를 y 방향으로 이동
            self._move_connected_segments(seg, dx=0.0, dy=dy)
            # 이동 후 직교 정렬
            self._orthogonalize_segments()
            # inlet 연결 보정
            self._ensure_inlet_connected()

            seg.drag_start_model = (start_mx, cur_my)

        self.redraw_all()

    def on_left_release(self, event):
        # Pencil 모드: 라인 확정 생성
        if self.mode == "pencil" and self._drawing and self._draw_start is not None:
            mx, my = self.screen_to_model(event.x, event.y)
            smx, smy = self.snap_model(mx, my)
            x1, y1 = self._draw_start
            x2, y2 = smx, smy
            # 정렬: 수평/수직 중 택1
            if abs(x2 - x1) >= abs(y2 - y1):
                y2 = y1
                vertical = False
            else:
                x2 = x1
                vertical = True

            # 세그먼트 생성 (수동 라벨/사이즈는 간단 표기)
            label_text = ""
            seg = DuctSegment(
                x1, y1, x2, y2,
                label_text,
                duct_w_mm=0,
                duct_h_mm=0,
                flow_m3h=0.0,
                vertical_only=vertical
            )
            self.segments.append(seg)
            # 미리보기 라인 제거 및 상태 초기화
            if self._preview_line_id is not None:
                self.canvas.delete(self._preview_line_id)
                self._preview_line_id = None
            self._drawing = False
            self._draw_start = None
            # 정렬/연결 보정 및 화면 갱신
            self._orthogonalize_segments()
            self._ensure_inlet_connected()
            self.redraw_all()
            return

        if self.dragging_segment is not None:
            self.dragging_segment.is_dragging = False
            self.dragging_segment.drag_start_model = None
            self.dragging_segment = None
        # 드래그 끝난 후 hover 재판정
        self.on_mouse_move(event)

    # ---------- 자동완성 ----------

    def auto_complete(self):
        """엔드포인트를 스냅/연결하고, outlet에 분기하여 풍량/사이즈 표기."""
        if not self.segments:
            return

        # 1) 끝점 스냅(근접한 끝점은 동일 좌표로 병합)
        eps = 1e-6
        def key_of(x, y, eps=eps):
            return (round(x/eps)*eps, round(y/eps)*eps)

        reps: dict[tuple[float,float], tuple[float,float]] = {}
        for seg in self.segments:
            for (x, y) in ((seg.mx1, seg.my1), (seg.mx2, seg.my2)):
                k = key_of(x, y)
                if k not in reps:
                    reps[k] = (x, y)
        # 스냅 재할당
        for seg in self.segments:
            k1 = key_of(seg.mx1, seg.my1)
            k2 = key_of(seg.mx2, seg.my2)
            seg.mx1, seg.my1 = reps.get(k1, (seg.mx1, seg.my1))
            seg.mx2, seg.my2 = reps.get(k2, (seg.mx2, seg.my2))

        # 2) 직교 정렬 유지
        self._orthogonalize_segments()

        # 3) outlet 포인트마다 가장 가까운 덕트라인 끝점에 분기하여 연결
        outlet_points = [p for p in self.points if getattr(p, 'kind', None) == 'outlet']
        duct_endpoints = []
        for seg in self.segments:
            duct_endpoints.append((seg.mx1, seg.my1))
            duct_endpoints.append((seg.mx2, seg.my2))

        for outlet in outlet_points:
            ox, oy = outlet.mx, outlet.my
            # outlet과 이미 연결된 덕트가 있으면 skip
            already_connected = False
            for seg in self.segments:
                if (abs(seg.mx1 - ox) < eps and abs(seg.my1 - oy) < eps) or (abs(seg.mx2 - ox) < eps and abs(seg.my2 - oy) < eps):
                    already_connected = True
                    break
            if already_connected:
                continue
            # 가장 가까운 duct endpoint 찾기
            min_dist = float('inf')
            nearest = None
            for dx, dy in duct_endpoints:
                dist = (dx - ox) ** 2 + (dy - oy) ** 2
                if dist < min_dist:
                    min_dist = dist
                    nearest = (dx, dy)
            if nearest is None:
                continue
            # outlet과 duct endpoint를 직선으로 연결 (수직/수평 우선)
            if abs(nearest[0] - ox) >= abs(nearest[1] - oy):
                # 수평 우선
                mid_x, mid_y = ox, nearest[1]
            else:
                # 수직 우선
                mid_x, mid_y = nearest[0], oy
            # 두 구간으로 분기: duct endpoint→mid, mid→outlet
            # 풍량은 outlet.flow
            q = getattr(outlet, 'flow', 0.0)
            try:
                D1 = calc_circular_diameter(q, 0.1)
                sel_big, sel_small, De_sel, theo_big, theo_small = size_rect_from_D1(D1, 2.0, 50)
                label_text = f"{sel_big}x{sel_small} {int(q)}m³/h"
            except Exception:
                label_text = f"{int(q)}m³/h"
                sel_big, sel_small = 0, 0
            # duct endpoint → mid
            self.segments.append(DuctSegment(nearest[0], nearest[1], mid_x, mid_y, label_text, sel_big, sel_small, q, vertical_only=(nearest[0]==mid_x)))
            # mid → outlet
            self.segments.append(DuctSegment(mid_x, mid_y, ox, oy, label_text, sel_big, sel_small, q, vertical_only=(mid_x==ox)))

        # 4) 고아(dangling) 제거: 한 끝점의 차수(degree)==1이고 그 끝점이 inlet/outlet에 붙지 않으면 제거
        def endpoints(seg: DuctSegment):
            return [(seg.mx1, seg.my1), (seg.mx2, seg.my2)]
        deg: dict[tuple[float,float], int] = {}
        for seg in self.segments:
            for pt in endpoints(seg):
                deg[pt] = deg.get(pt, 0) + 1

        attached_points = {(p.mx, p.my) for p in self.points}

        def is_attached(pt: tuple[float,float]):
            return pt in attached_points

        kept: list[DuctSegment] = []
        for seg in self.segments:
            pts = endpoints(seg)
            ok = True
            for pt in pts:
                if deg.get(pt, 0) <= 1 and not is_attached(pt):
                    ok = False
                    break
            if ok:
                kept.append(seg)
        self.segments = kept

        # 5) inlet 자동 연결 보정
        self._ensure_inlet_connected()

        # 6) 라벨 갱신(풍량/사이즈 표기)
        for seg in self.segments:
            q = seg.flow
            try:
                D1 = calc_circular_diameter(q, 0.1)
                sel_big, sel_small, De_sel, theo_big, theo_small = size_rect_from_D1(D1, 2.0, 50)
                seg.label_text = f"{sel_big}x{sel_small} {int(q)}m³/h"
                seg.duct_w_mm = sel_big
                seg.duct_h_mm = sel_small
            except Exception:
                seg.label_text = f"{int(q)}m³/h"
        self.redraw_all()

    # ---------- 연결된 세그먼트 이동 (inlet/outlet 고정) ----------

    def _move_connected_segments(self, base_seg: DuctSegment, dx: float, dy: float):
        """
        base_seg와 연결된 세그먼트 전체를 dx, dy만큼 평행 이동.
        다만 inlet / outlet 점과 '정확히 붙어 있는' 세그먼트 끝점은 고정.
        """
        # 세그먼트 간 연결은 '공통 끝점 좌표가 같은 경우'로 판단
        def seg_endpoints(seg: DuctSegment):
            return [(seg.mx1, seg.my1), (seg.mx2, seg.my2)]

        # BFS로 연결된 세그먼트 찾기
        connected = set()
        queue = [base_seg]
        connected.add(base_seg)

        while queue:
            cur = queue.pop(0)
            cur_ends = seg_endpoints(cur)
            for other in self.segments:
                if other in connected:
                    continue
                other_ends = seg_endpoints(other)
                if any(
                    (abs(ex1 - ex2) < 1e-9 and abs(ey1 - ey2) < 1e-9)
                    for (ex1, ey1) in cur_ends
                    for (ex2, ey2) in other_ends
                ):
                    connected.add(other)
                    queue.append(other)

        # inlet / outlet 점 위치 목록
        point_positions = [(p.mx, p.my) for p in self.points]

        def is_attached_to_point(x, y):
            for px, py in point_positions:
                if abs(px - x) < 1e-9 and abs(py - y) < 1e-9:
                    return True
            return False

        # 실제 이동
        for seg in connected:
            # 끝점 1
            if not is_attached_to_point(seg.mx1, seg.my1):
                seg.mx1 += dx
                seg.my1 += dy
            # 끝점 2
            if not is_attached_to_point(seg.mx2, seg.my2):
                seg.mx2 += dx
                seg.my2 += dy

    # ---------- 세그먼트 직교 정렬(orthogonal snap) ----------

    def _orthogonalize_segments(self):
        """
        모든 세그먼트 좌표를 수평/수직에 맞게 정렬한다.
        - vertical_only == True  → x1=x2, y1,y2는 그대로 (수직)
        - vertical_only == False → y1=y2, x1,x2는 그대로 (수평)
        또한, 끝점이 같은 위치에 있어야 하는 세그먼트들은 좌표를 공유하도록 보정한다.
        """
        if not self.segments:
            return

        # 1) 각 세그먼트 자체를 수평/수직에 맞춤
        for seg in self.segments:
            if seg.vertical_only:
                x_avg = (seg.mx1 + seg.mx2) / 2.0
                seg.mx1 = x_avg
                seg.mx2 = x_avg
            else:
                y_avg = (seg.my1 + seg.my2) / 2.0
                seg.my1 = y_avg
                seg.my2 = y_avg

        # 2) 이어지는 세그먼트들의 끝점이 정확히 일치하도록 보정
        def key_of(x, y, eps=1e-6):
            return (round(x/eps)*eps, round(y/eps)*eps)

        rep = {}  # (kx,ky) -> (rep_x, rep_y)

        for seg in self.segments:
            for (x, y) in ((seg.mx1, seg.my1), (seg.mx2, seg.my2)):
                k = key_of(x, y)
                if k not in rep:
                    rep[k] = (x, y)

        for seg in self.segments:
            k1 = key_of(seg.mx1, seg.my1)
            if k1 in rep:
                seg.mx1, seg.my1 = rep[k1]

            k2 = key_of(seg.mx2, seg.my2)
            if k2 in rep:
                seg.mx2, seg.my2 = rep[k2]

    # ---------- inlet과 메인 덕트 자동 연결 ----------

    def _ensure_inlet_connected(self):
        """
        inlet 점과 가장 가까운 덕트가 떨어져 있으면,
        inlet에서 그 덕트까지 수평/수직 리저 세그먼트를 자동 생성한다.
        이미 연결되어 있으면 아무 것도 하지 않는다.
        """
        if not self.points:
            return

        inlet = self.points[0]
        if inlet.kind != "inlet":
            return

        ix, iy = inlet.mx, inlet.my

        # 1) 이미 inlet 끝점과 정확히 붙어 있는 세그먼트가 있는지 검사
        def seg_endpoints(seg: DuctSegment):
            return [(seg.mx1, seg.my1), (seg.mx2, seg.my2)]

        for seg in self.segments:
            for (x, y) in seg_endpoints(seg):
                if abs(x - ix) < 1e-9 and abs(y - iy) < 1e-9:
                    return  # 이미 연결됨

        # 2) 가장 가까운 세그먼트를 찾음
        nearest_seg = None
        nearest_dist = float("inf")
        proj_point = None  # inlet의 투영점

        for seg in self.segments:
            if seg.vertical_only:
                x0 = seg.mx1
                y0 = min(seg.my1, seg.my2)
                y1 = max(seg.my1, seg.my2)
                py = min(max(iy, y0), y1)
                px = x0
            else:
                y0 = seg.my1
                x0 = min(seg.mx1, seg.mx2)
                x1 = max(seg.mx1, seg.mx2)
                px = min(max(ix, x0), x1)
                py = y0

            dist = math.hypot(px - ix, py - iy)
            if dist < nearest_dist:
                nearest_dist = dist
                nearest_seg = seg
                proj_point = (px, py)

        if nearest_seg is None or proj_point is None:
            return

        px, py = proj_point

        if nearest_dist < 1e-9:
            return  # 이미 겹침

        # 3) inlet에서 투영점까지 수평+수직 리저 생성
        cur_x, cur_y = ix, iy

        # 수평 리저 (필요 시)
        if abs(px - cur_x) > 1e-9:
            seg_h = DuctSegment(
                cur_x, cur_y,
                px, cur_y,
                label_text="",
                duct_w_mm=0,
                duct_h_mm=0,
                flow_m3h=0.0,
                vertical_only=False
            )
            self.segments.append(seg_h)
            cur_x = px

        # 수직 리저 (필요 시)
        if abs(py - cur_y) > 1e-9:
            seg_v = DuctSegment(
                cur_x, cur_y,
                cur_x, py,
                label_text="",
                duct_w_mm=0,
                duct_h_mm=0,
                flow_m3h=0.0,
                vertical_only=True
            )
            self.segments.append(seg_v)

    # ---------- 종합 사이징 ----------

    def draw_duct_network(self, dp_mmAq_per_m: float, aspect_ratio: float):
        if len(self.points) < 2:
            messagebox.showwarning("경고", "점이 2개 이상 있어야 종합 사이징을 할 수 있습니다.")
            return

        self.segments.clear()

        inlet = self.points[0]
        if inlet.kind != "inlet":
            messagebox.showwarning("경고", "첫 번째 점은 Air inlet이어야 합니다.")
            return

        Q_inlet = inlet.flow
        if Q_inlet <= 0:
            messagebox.showwarning("경고", "Air inlet 풍량이 0 이하입니다. 좌측 풍량 값을 확인해주세요.")
            return

        outlets = [p for p in self.points[1:] if p.kind == "outlet"]
        if not outlets:
            messagebox.showwarning("경고", "Air outlet 점이 없습니다.")
            return

        total_out = sum(ot.flow for ot in outlets)
        if abs(total_out - Q_inlet) > 1e-6:
            messagebox.showwarning(
                "경고",
                f"Air outlet 총 풍량({total_out:.1f} m³/h)가 "
                f"inlet 풍량({Q_inlet:.1f} m³/h)과 다릅니다.\n"
                "그래도 계산을 계속 진행합니다."
            )

        # --- 1) 메인선 y좌표: inlet 높이로 고정 ---
        y_main = inlet.my

        # --- 2) 이론상 inlet↔메인 리저 (필요 시) ---
        if abs(inlet.my - y_main) > 1e-9:
            Q_riser = Q_inlet
            try:
                D1 = calc_circular_diameter(Q_riser, dp_mmAq_per_m)
                sel_big, sel_small, De_sel, theo_big, theo_small = size_rect_from_D1(
                    D1, aspect_ratio, 50
                )
                label_text = f"{sel_big}x{sel_small}"
                seg_riser = DuctSegment(
                    inlet.mx, inlet.my, inlet.mx, y_main,
                    label_text,
                    duct_w_mm=sel_big,
                    duct_h_mm=sel_small,
                    flow_m3h=Q_riser,
                    vertical_only=True
                )
                self.segments.append(seg_riser)
            except ValueError:
                pass

        # --- 3) 좌우 outlet 분리 ---
        left_outlets = [ot for ot in outlets if ot.mx < inlet.mx - 1e-9]
        right_outlets = [ot for ot in outlets if ot.mx > inlet.mx + 1e-9]
        center_outlets = [ot for ot in outlets if abs(ot.mx - inlet.mx) <= 1e-9]

        # --- 5) 오른쪽 메인 ---
        if right_outlets:
            right_sorted = sorted(right_outlets, key=lambda p: p.mx)
            x_nodes_right = [inlet.mx] + [ot.mx for ot in right_sorted]
            x_nodes_right = sorted(set(x_nodes_right))

            def flow_downstream_from_right(x_pos: float) -> float:
                s = 0.0
                for ot in right_sorted:
                    if ot.mx > x_pos + 1e-9:
                        s += ot.flow
                return s

            for i in range(len(x_nodes_right) - 1):
                x1 = x_nodes_right[i]
                x2 = x_nodes_right[i + 1]
                if abs(x2 - x1) < 1e-9:
                    continue

                Q_main = flow_downstream_from_right(x1)
                if Q_main <= 0:
                    continue

                try:
                    D1 = calc_circular_diameter(Q_main, dp_mmAq_per_m)
                    sel_big, sel_small, De_sel, theo_big, theo_small = size_rect_from_D1(
                        D1, aspect_ratio, 50
                    )
                except ValueError:
                    continue

                label_text = f"{sel_big}x{sel_small}"

                seg = DuctSegment(
                    x1, y_main, x2, y_main,
                    label_text,
                    duct_w_mm=sel_big,
                    duct_h_mm=sel_small,
                    flow_m3h=Q_main,
                    vertical_only=False
                )
                self.segments.append(seg)

        # --- 6) 왼쪽 메인 ---
        if left_outlets:
            left_sorted = sorted(left_outlets, key=lambda p: p.mx, reverse=True)
            x_nodes_left = [inlet.mx] + [ot.mx for ot in left_sorted]
            x_nodes_left = sorted(set(x_nodes_left), reverse=True)

            def flow_downstream_from_left(x_pos: float) -> float:
                s = 0.0
                for ot in left_sorted:
                    if ot.mx < x_pos - 1e-9:
                        s += ot.flow
                return s

            for i in range(len(x_nodes_left) - 1):
                x1 = x_nodes_left[i]
                x2 = x_nodes_left[i + 1]
                if abs(x2 - x1) < 1e-9:
                    continue

                Q_main = flow_downstream_from_left(x1)
                if Q_main <= 0:
                    continue

                try:
                    D1 = calc_circular_diameter(Q_main, dp_mmAq_per_m)
                    sel_big, sel_small, De_sel, theo_big, theo_small = size_rect_from_D1(
                        D1, aspect_ratio, 50
                    )
                except ValueError:
                    continue

                label_text = f"{sel_big}x{sel_small}"

                seg = DuctSegment(
                    x1, y_main, x2, y_main,
                    label_text,
                    duct_w_mm=sel_big,
                    duct_h_mm=sel_small,
                    flow_m3h=Q_main,
                    vertical_only=False
                )
                self.segments.append(seg)

        # --- 7) 브랜치 수직 세그먼트 ---
        for ot in outlets:
            Q_branch = ot.flow
            if Q_branch <= 0:
                continue

            bx = ot.mx
            by = y_main
            ox = ot.mx
            oy = ot.my

            try:
                D1 = calc_circular_diameter(Q_branch, dp_mmAq_per_m)
                sel_big, sel_small, De_sel, theo_big, theo_small = size_rect_from_D1(
                    D1, aspect_ratio, 50
                )
            except ValueError:
                continue

            label_text = f"{sel_big}x{sel_small}"

            seg = DuctSegment(
                bx, by, ox, oy,
                label_text,
                duct_w_mm=sel_big,
                duct_h_mm=sel_small,
                flow_m3h=Q_branch,
                vertical_only=True
            )
            self.segments.append(seg)

        self.redraw_all()

    # ---------- Undo & Clear & 균등 배분 ----------

    def undo_last_point(self):
        if not self.points:
            messagebox.showinfo("Undo", "되돌릴 점이 없습니다.")
            return
        self.points.pop()
        self.segments.clear()
        self.redraw_all()
        # (추가) 삭제 시 갱신
        self._notify_points_changed()

    def clear_all(self):
        self.points.clear()
        self.segments.clear()
        self.redraw_all()
        # (추가) 전체 삭제 시 갱신
        self._notify_points_changed()

    def distribute_equal_flow(self):
        if not self.points:
            messagebox.showwarning("경고", "먼저 Air inlet을 포함한 점을 찍어주세요.")
            return
        if len(self.points) < 2:
            messagebox.showwarning("경고", "최소 2개 이상의 점이 있어야 균등 배분이 가능합니다.")
            return

        Q_in = self.inlet_flow
        if Q_in <= 0:
            messagebox.showwarning("경고", "Air inlet 풍량이 0 이하입니다. 좌측 풍량 값을 확인해주세요.")
            return

        n_out = len(self.points) - 1
        Q_each = Q_in / n_out

        for idx, p in enumerate(self.points):
            if idx == 0:
                p.flow = Q_in
            else:
                p.flow = Q_each

        self.segments.clear()
        self.redraw_all()
        # 좌표 변화는 없지만, 사용자 기대상 “지정” 후 갱신 원할 수 있어 호출
        self._notify_points_changed()


# =========================
# (추가) outlet 상대 위치 통계 표시 함수
# =========================

relpos_text_widget = None  # GUI 생성 후 할당됨

def update_outlet_calculations(pal: Palette):
    """
    inlet 기준 outlet들의 상대 위치 통계 계산
    1) Σ(inlet.x - outlet.x)
    2) Σ(inlet.y - outlet.y)
    3) Σ(inlet.x - outlet.x)^2
    4) Σ(inlet.y - outlet.y)^2
    """
    global relpos_text_widget
    if relpos_text_widget is None:
        return

    if not pal.points or pal.points[0].kind != "inlet":
        text = "inlet이 없습니다."
    else:
        inlet = pal.points[0]
        outlets = [p for p in pal.points[1:] if p.kind == "outlet"]
        
        sum_dx = 0.0
        sum_dy = 0.0
        sum_sq_dx = 0.0
        sum_sq_dy = 0.0
        
        for p in outlets:
            dx = inlet.mx - p.mx  # inlet - outlet
            dy = inlet.my - p.my  # inlet - outlet
            sum_dx += dx
            sum_dy += dy
            sum_sq_dx += dx**2
            sum_sq_dy += dy**2
            
        text = (
            f"Outlet 개수: {len(outlets)}\n"
            f"1. Σ(inlet.x - outlet.x) : {sum_dx:.2f} m\n"
            f"2. Σ(inlet.y - outlet.y) : {sum_dy:.2f} m\n"
            f"3. Σ(inlet.x - outlet.x)²: {sum_sq_dx:.2f} m²\n"
            f"4. Σ(inlet.y - outlet.y)²: {sum_sq_dy:.2f} m²"
        )

    relpos_text_widget.config(state="normal")
    relpos_text_widget.delete("1.0", "end")
    relpos_text_widget.insert("end", text)
    relpos_text_widget.config(state="disabled")


# =========================
# GUI 이벤트 함수
# =========================

def calculate():
    try:
        q = float(cubic_meter_hour_entry.get())
        dp = float(resistance_entry.get())

        D1 = calc_circular_diameter(q, dp)
        D2 = round_step_up(D1, 50)

        try:
            r = float(aspect_ratio_combo.get())
        except ValueError:
            r = 2.0

        sel_big, sel_small, De_sel, theo_big, theo_small = size_rect_from_D1(D1, r, 50)

        text = (
            f"1. 등가원형 직경 (mm) : {D1:.0f}\n"
            f"2. 원형직경 (mm)     : {D2}\n"
            f"3. 사각덕트 사이즈 (mm X mm, 50mm 조정) : {sel_big} X {sel_small}\n"
            f"4. 사각덕트 사이즈 (mm X mm, 조정 전)  : {theo_big:.1f} X {theo_small:.1f}\n"
            f"※ 팔레트 격자 1칸 = 0.5 m (가로/세로)\n"
            f"※ 덕트 사이즈 단위는 mm 입니다."
        )

        results_text_widget.config(state="normal")
        results_text_widget.delete("1.0", "end")
        results_text_widget.insert("end", text)
        results_text_widget.config(state="disabled")

        palette.set_inlet_flow(q)

        # inlet 유무/상태 표시도 즉시 갱신
        update_outlet_calculations(palette)

    except ValueError as e:
        messagebox.showerror("입력 오류", f"입력값을 확인하세요!\n\n{e}")
    except Exception as e:
        messagebox.showerror("알 수 없는 오류", f"알 수 없는 오류가 발생했습니다:\n\n{e}")


def total_sizing():
    try:
        dp = float(resistance_entry.get())
    except ValueError:
        messagebox.showerror("입력 오류", "정압값을 올바르게 입력해주세요.")
        return

    try:
        r = float(aspect_ratio_combo.get())
    except ValueError:
        r = 2.0

    palette.draw_duct_network(dp_mmAq_per_m=dp, aspect_ratio=r)

    total_area_m2 = 0.0
    for seg in palette.segments:
        L = seg.length_m()
        w_m = seg.duct_w_mm / 1000.0
        h_m = seg.duct_h_mm / 1000.0
        area = (w_m + h_m) * 2 * L
        total_area_m2 += area

    results_text_widget.config(state="normal")
    base = results_text_widget.get("1.0", "end").rstrip()
    if base:
        base += "\n"
    base += f"5. 덕트 철판 소요량 (m²) : {total_area_m2:.1f}"

    results_text_widget.delete("1.0", "end")
    results_text_widget.insert("end", base)
    results_text_widget.config(state="disabled")


def clear_palette():
    palette.clear_all()


def equal_distribution():
    try:
        q = float(cubic_meter_hour_entry.get())
        palette.set_inlet_flow(q)
    except ValueError:
        messagebox.showerror("입력 오류", "풍량 값을 올바르게 입력해주세요.")
        return

    palette.distribute_equal_flow()


def undo_point():
    palette.undo_last_point()

# =========================
# GUI 구성
# =========================

root = tk.Tk()
root.title("덕트 사이징 프로그램 (Grid 0.5m, mm 덕트, 철판 소요량)")

main_frame = tk.Frame(root)
main_frame.pack(fill="both", expand=True, padx=10, pady=10)

root.bind("<Control-z>", lambda event: undo_point())

# 왼쪽 정보 입력창 (외기/실내/급기 온도 및 발열량)
info_frame = tk.Frame(main_frame, width=180)
info_frame.pack(side="left", fill="y", padx=(0,10))

tk.Label(info_frame, text="외기/실내/급기/발열량", font=("Arial", 10, "bold")).pack(anchor="w", pady=(6,4))

tk.Label(info_frame, text="외기온도 (°C):").pack(anchor="w")
outdoor_temp_entry = tk.Entry(info_frame, width=10)
outdoor_temp_entry.pack(anchor="w", pady=2)
outdoor_temp_entry.insert(0, "-5.0")

tk.Label(info_frame, text="실내온도 (°C):").pack(anchor="w")
indoor_temp_entry = tk.Entry(info_frame, width=10)
indoor_temp_entry.pack(anchor="w", pady=2)
indoor_temp_entry.insert(0, "25.0")

tk.Label(info_frame, text="급기온도 (°C):").pack(anchor="w")
supply_temp_entry = tk.Entry(info_frame, width=10)
supply_temp_entry.pack(anchor="w", pady=2)
supply_temp_entry.insert(0, "18.0")

tk.Label(info_frame, text="일반 발열량 (W/m²):").pack(anchor="w")
heat_norm_entry = tk.Entry(info_frame, width=10)
heat_norm_entry.pack(anchor="w", pady=2)
heat_norm_entry.insert(0, "0.00")

tk.Label(info_frame, text="장비 발열량 (W/m²):").pack(anchor="w")
heat_equip_entry = tk.Entry(info_frame, width=10)
heat_equip_entry.pack(anchor="w", pady=2)
heat_equip_entry.insert(0, "0.00")

left_frame = tk.Frame(main_frame)
left_frame.pack(side="left", anchor="w")

right_frame = tk.Frame(main_frame, bg="#f5f5f5", bd=1, relief="solid")
right_frame.configure(width=700)
right_frame.pack(side="right", fill="both", expand=True)

tk.Label(left_frame, text="풍량 (m³/h):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
cubic_meter_hour_entry = tk.Entry(left_frame, width=10)
cubic_meter_hour_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
cubic_meter_hour_entry.insert(0, "5000")

tk.Label(left_frame, text="정압값 (mmAq/m):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
resistance_entry = tk.Entry(left_frame, width=10)
resistance_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
resistance_entry.insert(0, "0.1")

tk.Label(left_frame, text="사각 덕트 종횡비 (b/a):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
aspect_ratio_combo = ttk.Combobox(left_frame, values=["1", "2", "3", "4"], state="readonly", width=5)
aspect_ratio_combo.current(1)
aspect_ratio_combo.grid(row=2, column=1, padx=5, pady=5, sticky="w")

tk.Button(left_frame, text="계산하기", command=calculate).grid(
    row=3, column=0, columnspan=2, pady=5, sticky="w"
)
tk.Button(left_frame, text="균등 풍량 배분", command=equal_distribution).grid(
    row=4, column=0, columnspan=2, pady=5, sticky="w"
)
tk.Button(left_frame, text="종합 사이징", command=total_sizing).grid(
    row=5, column=0, columnspan=2, pady=5, sticky="w"
)
tk.Button(left_frame, text="팔레트 전체 지우기", command=clear_palette).grid(
    row=6, column=0, columnspan=2, pady=5, sticky="w"
)

# 수동 그리기/자동완성 버튼
tk.Button(left_frame, text="펜슬 모드", command=lambda: palette.set_mode_pencil()).grid(
    row=7, column=0, padx=5, pady=5, sticky="w"
)
tk.Button(left_frame, text="자동완성", command=lambda: palette.auto_complete()).grid(
    row=7, column=1, padx=5, pady=5, sticky="w"
)

# 결과 출력 Text + Scrollbar
results_frame = tk.Frame(left_frame)
results_frame.grid(row=8, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")

left_frame.grid_rowconfigure(8, weight=1)
left_frame.grid_columnconfigure(0, weight=1)
left_frame.grid_columnconfigure(1, weight=1)

results_text_widget = tk.Text(
    results_frame,
    width=36,
    height=9,
    wrap="word",
    bg="white",
    relief="solid"
)
results_text_widget.pack(side="left", fill="both", expand=True)

results_scrollbar = tk.Scrollbar(results_frame, orient="vertical", command=results_text_widget.yview)
results_scrollbar.pack(side="right", fill="y")

results_text_widget.configure(yscrollcommand=results_scrollbar.set)
results_text_widget.config(state="disabled")

# (추가/수정) outlet 상대 위치 통계 표시용 별도 텍스트 창 (높이 증가)
relpos_frame = tk.Frame(left_frame)
relpos_frame.grid(row=9, column=0, columnspan=2, padx=5, pady=(0, 5), sticky="nsew")

tk.Label(relpos_frame, text="Outlet 상대 위치 통계 (inlet - outlet):").pack(anchor="w")

relpos_text_widget = tk.Text(
    relpos_frame,
    width=36,
    height=6,  # 4가지 항목 표시를 위해 높이 증가
    wrap="word",
    bg="white",
    relief="solid"
)
relpos_text_widget.pack(fill="both", expand=True)
relpos_text_widget.config(state="disabled")

palette = Palette(right_frame)

# (추가) points 변경 시 자동 갱신 연결 + 초기 1회 표시
palette.points_changed_callback = update_outlet_calculations
update_outlet_calculations(palette)

root.mainloop()
