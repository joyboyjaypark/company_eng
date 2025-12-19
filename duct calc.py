import tkinter as tk
from tkinter import messagebox, ttk, simpledialog
import copy
from collections import deque, defaultdict
import math

# =========================
# 1. 계산 함수들 (Engineering Logic)
# =========================

def calc_circular_diameter(q_m3h: float, dp_mmAq_per_m: float) -> float:
    """풍량(m3/h)과 정압(mmAq/m)을 이용해 등가 원형 덕트 직경 D1(mm) 계산"""
    if q_m3h <= 0:
        return 0.0
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
    """사각 덕트 a,b(mm)에 대한 등가 원형 직경 De(mm) (ASHRAE 공식)"""
    if a_mm <= 0 or b_mm <= 0:
        return 0.0
    a, b = float(a_mm), float(b_mm)
    return 1.30 * (a*b)**0.625 / (a + b)**0.25


def size_rect_from_D1(D1: float, aspect_ratio: float, step: float = 50):
    """원형 직경 D1과 종횡비를 기준으로 사각 덕트 사이즈 산출"""
    if D1 <= 0:
        return 0, 0, 0.0, 0.0, 0.0
    if aspect_ratio <= 0:
        raise ValueError("종횡비(b/a)는 0보다 커야 합니다.")

    De_target = float(D1)
    r = float(aspect_ratio)

    # 이론적 사각 사이즈
    a_theo = De_target * (1 + r)**0.25 / (1.30 * r**0.625)
    b_theo = r * a_theo
    theo_big, theo_small = max(a_theo, b_theo), min(a_theo, b_theo)

    # 후보군 계산
    small_up = round_step_up(theo_small, step)
    big_down = max(round_step_down(theo_big, step), step)
    De1 = rect_equiv_diameter(small_up, big_down)

    a_up = round_step_up(a_theo, step)
    b_up = round_step_up(b_theo, step)
    De2 = rect_equiv_diameter(a_up, b_up)

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

def calc_rect_other_side(D1: float, fixed_side_mm: float, step: float = 50):
    """한 변이 고정된 상태에서 D1을 만족하는 다른 한 변 계산"""
    if D1 <= 0:
        return 0, 0, 0.0
    if fixed_side_mm <= 0:
        raise ValueError("고정 변의 길이는 0보다 커야 합니다.")
    
    fixed = round_step_up(fixed_side_mm, step)
    other = 50.0
    # 50mm씩 증가시키며 최소 요구 직경을 만족하는지 확인
    while True:
        de = rect_equiv_diameter(fixed, other)
        if de >= D1 - 0.1: # 허용 오차 고려
            break
        other += step
        if other > 10000: # 무한 루프 방지
            break
            
    sel_big = max(fixed, other)
    sel_small = min(fixed, other)
    return int(sel_big), int(sel_small), round(de, 1)

def perform_sizing(q: float, dp: float, use_fixed: bool, fixed_val: float, aspect_r: float):
    """사이징 옵션(종횡비 vs 고정변)에 따라 적절한 함수를 호출하여 결과 반환"""
    if q <= 0:
        return 0, 0, f"{int(q)}m³/h"
    try:
        D1 = calc_circular_diameter(q, dp)
        if use_fixed:
            w, h, de = calc_rect_other_side(D1, fixed_val, 50)
            # 고정 변 모드에서는 이론치 계산 생략
        else:
            w, h, de, theo_big, theo_small = size_rect_from_D1(D1, aspect_r, 50)
        
        label = f"{w}x{h} {int(q)}m³/h"
        return w, h, label
    except Exception:
        return 0, 0, f"{int(q)}m³/h"


# =========================
# 2. 데이터 모델 및 팔레트 클래스
# =========================

GRID_STEP_MODEL = 0.5
INITIAL_SCALE = 40.0

class AirPoint:
    def __init__(self, mx, my, kind, flow):
        self.mx = mx
        self.my = my
        self.kind = kind
        self.flow = flow
        self.canvas_id = None
        self.text_id = None


class DuctSegment:
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
        self.vertical_only = vertical_only
        self.line_ids = []
        self.text_id = None
        self.leader_id = None
        
        # 상호작용 상태
        self.is_hovered = False
        self.is_dragging = False
        self.drag_start_model = None

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

        # 그리기 모드
        self.mode = "pan"  # "pan" or "pencil"
        self._drawing = False
        self._draw_start = None
        self._preview_line_id = None

        # 마우스 상호작용
        self.hovered_segment = None
        self.dragging_segment = None
        self.hovered_text_id = None
        self.current_mouse_model = None

        self.points_changed_callback = None

        # Undo stack: store snapshots of (points, segments, inlet_flow)
        self._undo_stack = []
        self._undo_limit = 100

        # 이벤트 바인딩
        self.canvas.bind("<Button-1>", self.on_left_click)
        self.canvas.bind("<Button-3>", self.on_right_click)
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        self.canvas.bind("<Button-2>", self.on_middle_press)
        self.canvas.bind("<B2-Motion>", self.on_middle_drag)
        self.canvas.bind("<Configure>", self.on_resize)

        self.canvas.bind("<Motion>", self.on_mouse_move)
        self.canvas.bind("<Leave>", self.on_mouse_leave)
        self.canvas.bind("<B1-Motion>", self.on_left_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_release)

        self.redraw_all()

    def _notify_points_changed(self):
        cb = getattr(self, "points_changed_callback", None)
        if callable(cb):
            try: cb(self)
            except: pass
        # notify sheet/segment-related changes as well
        cb2 = getattr(self, "sheet_changed_callback", None)
        if callable(cb2):
            try: cb2(self)
            except: pass

    def _snapshot(self):
        return {
            'points': copy.deepcopy(self.points),
            'segments': copy.deepcopy(self.segments),
            'inlet_flow': self.inlet_flow,
        }

    def push_undo(self):
        try:
            snap = self._snapshot()
            self._undo_stack.append(snap)
            if len(self._undo_stack) > self._undo_limit:
                self._undo_stack.pop(0)
        except Exception:
            pass

    def _restore_snapshot(self, snap):
        if not snap: return
        self.points = copy.deepcopy(snap.get('points', []))
        self.segments = copy.deepcopy(snap.get('segments', []))
        self.inlet_flow = snap.get('inlet_flow', 0.0)
        # clear canvas-related ids and transient flags
        for p in self.points:
            try:
                p.canvas_id = None; p.text_id = None
            except: pass
        for seg in self.segments:
            try:
                seg.line_ids = []
                seg.text_id = None
                seg.leader_id = None
                seg.is_hovered = False
                seg.is_dragging = False
                seg.drag_start_model = None
            except: pass
        self.redraw_all()
        self._notify_points_changed()

    def undo(self):
        if not self._undo_stack: return
        snap = self._undo_stack.pop()
        self._restore_snapshot(snap)

    def set_mode_pencil(self):
        if self.mode == "pencil":
            self.set_mode_pan()
            return
        self.mode = "pencil"
        try: self.canvas.config(cursor="pencil")
        except: self.canvas.config(cursor="crosshair")

    def set_mode_pan(self):
        self.mode = "pan"
        self.canvas.config(cursor="arrow")

    def model_to_screen(self, mx, my):
        sx = mx * self.scale_factor + self.offset_x
        sy = my * self.scale_factor + self.offset_y
        return sx, sy

    def screen_to_model(self, sx, sy):
        mx = (sx - self.offset_x) / self.scale_factor
        my = (sy - self.offset_y) / self.scale_factor
        return mx, my

    def draw_grid(self):
        self.canvas.delete(self.grid_tag)
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w <= 0 or h <= 0: return

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
            self.canvas.create_line(sx1, sy1, sx2, sy2, fill="#e0e0e0", tags=self.grid_tag)
            x += GRID_STEP_MODEL

        y = y_start
        while y <= y_end:
            sx1, sy1 = self.model_to_screen(x_start, y)
            sx2, sy2 = self.model_to_screen(x_end, y)
            self.canvas.create_line(sx1, sy1, sx2, sy2, fill="#e0e0e0", tags=self.grid_tag)
            y += GRID_STEP_MODEL

    def redraw_all(self):
        self.canvas.delete("all")
        self.draw_grid()

        # 점 그리기
        for p in self.points:
            sx, sy = self.model_to_screen(p.mx, p.my)
            color = "red" if p.kind == "inlet" else "blue"
            r = 5
            p.canvas_id = self.canvas.create_oval(sx - r, sy - r, sx + r, sy + r, fill=color, outline="")
            label = f"{p.flow:.1f}"
            p.text_id = self.canvas.create_text(sx + 10, sy - 10, text=label, fill="black", font=("Arial", 8))

        # 덕트 선 그리기
        for seg in self.segments:
            seg.line_ids.clear()
            seg.text_id = None
            seg.leader_id = None

            mx1, my1, mx2, my2 = seg.mx1, seg.my1, seg.mx2, seg.my2
            sx1, sy1 = self.model_to_screen(mx1, my1)
            sx2, sy2 = self.model_to_screen(mx2, my2)

            line_width = 3 if seg.is_hovered else 1
            line_color = "gray50"
            
            seg.line_ids.append(self.canvas.create_line(sx1, sy1, sx2, sy2, fill=line_color, width=line_width, tags=("duct_line",)))
            
            # 텍스트 위치 계산 (세로선이면 오른쪽, 가로선이면 아래쪽)
            is_vert_draw = abs(sx1 - sx2) < abs(sy1 - sy2)
            
            leader_length_px = 15
            text_offset_px = 5
            text_tags = ("duct_text",) # 클릭 감지용 태그

            if is_vert_draw: # 세로선
                mid_sx = sx1
                mid_sy = (sy1 + sy2) / 2.0
                vx, vy = mid_sx, mid_sy
                seg.leader_id = self.canvas.create_line(vx, vy, vx + leader_length_px, vy, fill="blue")
                tx, ty = vx + leader_length_px + text_offset_px, vy
                seg.text_id = self.canvas.create_text(tx, ty, text=seg.label_text, fill="blue", font=("Arial", 8), anchor="w", tags=text_tags)
            else: # 가로선
                mid_sx = (sx1 + sx2) / 2.0
                mid_sy = sy1
                hx, hy = mid_sx, mid_sy
                seg.leader_id = self.canvas.create_line(hx, hy, hx, hy - leader_length_px, fill="blue")
                tx, ty = hx, hy - leader_length_px - text_offset_px
                seg.text_id = self.canvas.create_text(tx, ty, text=seg.label_text, fill="blue", font=("Arial", 8), anchor="s", tags=text_tags)

    # 툴팁 기능 제거: 마우스 위치에 자동 계산되는 텍스트 박스는 더 이상 표시하지 않음

    # _draw_tooltip removed: mouse-over auto-calculation box deleted per user request

    def on_resize(self, event):
        self.redraw_all()

    def snap_model(self, mx, my):
        smx = round(mx / GRID_STEP_MODEL) * GRID_STEP_MODEL
        smy = round(my / GRID_STEP_MODEL) * GRID_STEP_MODEL
        return smx, smy

    def on_middle_press(self, event):
        self.pan_start_screen = (event.x, event.y)

    def on_middle_drag(self, event):
        if self.pan_start_screen is None: return
        sx0, sy0 = self.pan_start_screen
        dx = event.x - sx0
        dy = event.y - sy0
        self.pan_start_screen = (event.x, event.y)
        self.offset_x += dx
        self.offset_y += dy
        self.redraw_all()

    def on_mousewheel(self, event):
        factor = 1.1 if event.delta > 0 else 0.9
        new_scale = self.scale_factor * factor
        if not (10.0 <= new_scale <= 400.0): return
        sx, sy = event.x, event.y
        mx, my = self.screen_to_model(sx, sy)
        self.scale_factor = new_scale
        self.offset_x = sx - mx * self.scale_factor
        self.offset_y = sy - my * self.scale_factor
        self.redraw_all()

    def _hit_test_segment(self, mx, my, tol=0.1):
        for seg in self.segments:
            if seg.vertical_only:
                if min(seg.my1, seg.my2) - tol <= my <= max(seg.my1, seg.my2) + tol:
                    if abs(mx - seg.mx1) <= tol: return seg
            else:
                if min(seg.mx1, seg.mx2) - tol <= mx <= max(seg.mx1, seg.mx2) + tol:
                    if abs(my - seg.my1) <= tol: return seg
        return None

    def set_inlet_flow(self, flow):
        self.push_undo()
        self.inlet_flow = float(flow)
        if self.points:
            p0 = self.points[0]
            if p0.kind == "inlet": p0.flow = self.inlet_flow
        self.redraw_all()
        self._notify_points_changed()

    def on_left_click(self, event):
        # 1. 덕트 텍스트 클릭 감지 (사이즈 수정 다이얼로그)
        closest_items = self.canvas.find_closest(event.x, event.y, halo=2)
        if closest_items:
            item_id = closest_items[0]
            tags = self.canvas.gettags(item_id)
            if "duct_text" in tags:
                target_seg = None
                for seg in self.segments:
                    if seg.text_id == item_id:
                        target_seg = seg
                        break
                if target_seg:
                    self._edit_duct_size_dialog(target_seg)
                    return

        # 2. Pencil 모드 (라인 그리기 시작)
        if self.mode == "pencil":
            mx, my = self.screen_to_model(event.x, event.y)
            smx, smy = self.snap_model(mx, my)
            self._draw_start = (smx, smy)
            self._drawing = True
            if self._preview_line_id is not None:
                self.canvas.delete(self._preview_line_id)
                self._preview_line_id = None
            sx, sy = self.model_to_screen(smx, smy)
            self._preview_line_id = self.canvas.create_line(sx, sy, sx, sy, fill="black", width=2, dash=(4, 2))
            return

        # 3. 점 생성 모드 (Inlet/Outlet 추가)
        mx, my = self.screen_to_model(event.x, event.y)
        seg = self._hit_test_segment(mx, my, tol=0.15)
        if seg is not None: return # 라인 위에는 점 생성 안 함

        sx, sy = event.x, event.y
        mx, my = self.screen_to_model(sx, sy)
        mx, my = self.snap_model(mx, my)

        # 기록 후 점 추가
        self.push_undo()
        if not self.points:
            flow = self.inlet_flow if self.inlet_flow > 0 else 0.0
            p = AirPoint(mx, my, "inlet", flow)
            self.points.append(p)
        else:
            p = AirPoint(mx, my, "outlet", 0.0)
            self.points.append(p)

        self.segments.clear()
        self.redraw_all()
        self._notify_points_changed()

    def _edit_duct_size_dialog(self, seg):
        """덕트 라벨 클릭 시 호출: 고정 변 입력 받아 재계산"""
        try: dp_current = float(resistance_entry.get())
        except: dp_current = 0.1
        if dp_current <= 0:
            messagebox.showerror("오류", "정압값이 유효하지 않습니다.")
            return
        
        msg = (
            f"현재 구간 풍량: {seg.flow} m³/h\n"
            f"현재 사이즈: {seg.duct_w_mm} x {seg.duct_h_mm}\n\n"
            f"고정할 덕트 한 변의 길이(mm)를 입력하세요.\n"
            f"(입력한 값과 계산된 값 중 큰 값이 먼저 표기됩니다)"
        )
        val = simpledialog.askinteger("덕트 사이즈 변경", msg, minvalue=50, maxvalue=5000)
        if val:
            self.push_undo()
            w, h, label = perform_sizing(seg.flow, dp_current, True, float(val), 1.0)
            seg.duct_w_mm = w
            seg.duct_h_mm = h
            seg.label_text = label
            self.redraw_all()
            self._notify_points_changed()

    def on_right_click(self, event):
        sx, sy = event.x, event.y
        mx, my = self.screen_to_model(sx, sy)

        # 우클릭: 펜슬 모드에서 라인 근처면 해당 라인 삭제
        if self.mode == "pencil":
            seg_hit = self._hit_test_segment(mx, my, tol=0.15)
            if seg_hit is not None:
                self.push_undo()
                try:
                    self.segments.remove(seg_hit)
                except ValueError:
                    pass
                # after removal, merge colinear neighbors if branch node disappeared
                try:
                    self._merge_colinear_neighbors()
                except Exception:
                    pass
                self.redraw_all()
                self._notify_points_changed()
                return

        target = self._find_point_near_model(mx, my, tol_model=0.3)
        if target is None: return

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
        if answer is None: return
        try:
            new_flow = float(answer)
            if new_flow < 0: raise ValueError
        except ValueError:
            messagebox.showerror("입력 오류", "0 이상 숫자로 입력해주세요.")
            return
        self.push_undo()
        target.flow = new_flow
        self.segments.clear()
        self.redraw_all()
        self._notify_points_changed()

    def _find_point_near_model(self, mx, my, tol_model=0.3):
        for p in self.points:
            if (p.mx - mx)**2 + (p.my - my)**2 <= tol_model**2: return p
        return None

    def _sum_outlet_flow(self, exclude=None):
        s = 0.0
        for p in self.points:
            if p.kind == "outlet" and p is not exclude: s += p.flow
        return s

    def _calc_remaining_flow(self, exclude=None):
        return self.inlet_flow - self._sum_outlet_flow(exclude=exclude)

    def on_mouse_move(self, event):
        mx, my = self.screen_to_model(event.x, event.y)
        self.current_mouse_model = (mx, my)

        if self.dragging_segment is not None:
            self.redraw_all()
            return

        seg = self._hit_test_segment(mx, my, tol=0.15)
        if self.hovered_segment is not None and self.hovered_segment is not seg:
            self.hovered_segment.is_hovered = False
            self.hovered_segment = None

        if seg is not None:
            seg.is_hovered = True
            self.hovered_segment = seg

        hovered_text = None
        text_item = self.canvas.find_closest(event.x, event.y, halo=2)
        if text_item and "duct_text" in self.canvas.gettags(text_item[0]):
            hovered_text = text_item[0]
        
        for p in self.points:
            if p.text_id is None: continue
            try:
                tx, ty = self.model_to_screen(p.mx, p.my)
                sx = tx + 10; sy = ty - 10
            except: continue
            if (event.x - sx)**2 + (event.y - sy)**2 <= 64:
                hovered_text = p.text_id
                break

        if hovered_text != self.hovered_text_id:
            if self.hovered_text_id is not None:
                try: self.canvas.itemconfigure(self.hovered_text_id, font=("Arial", 8))
                except: pass
            self.hovered_text_id = hovered_text
            if self.hovered_text_id is not None:
                try: self.canvas.itemconfigure(self.hovered_text_id, font=("Arial", 8, "bold"))
                except: pass
        self.redraw_all()

    def on_mouse_leave(self, event):
        self.current_mouse_model = None
        self.redraw_all()

    def on_left_drag(self, event):
        if self.mode == "pencil" and self._drawing and self._draw_start is not None:
            mx, my = self.screen_to_model(event.x, event.y)
            smx, smy = self.snap_model(mx, my)
            sx1, sy1 = self.model_to_screen(*self._draw_start)
            sx2, sy2 = self.model_to_screen(smx, smy)
            dx = abs(smx - self._draw_start[0])
            dy = abs(smy - self._draw_start[1])
            if dx >= dy: sy2 = sy1
            else: sx2 = sx1
            if self._preview_line_id is not None:
                self.canvas.coords(self._preview_line_id, sx1, sy1, sx2, sy2)
            return

        mx, my = self.screen_to_model(event.x, event.y)
        self.current_mouse_model = (mx, my)

        if self.dragging_segment is None:
            seg = self._hit_test_segment(mx, my, tol=0.15)
            if seg is None: return
            # record state before starting a drag operation
            self.push_undo()
            self.dragging_segment = seg
            seg.is_dragging = True
            seg.drag_start_model = (mx, my)

        seg = self.dragging_segment
        if seg is None: return

        cur_mx, cur_my = mx, my
        start_mx, start_my = seg.drag_start_model

        if seg.vertical_only:
            dx_raw = cur_mx - start_mx
            base_x = (seg.mx1 + seg.mx2) / 2.0
            new_x = base_x + dx_raw
            snapped_x, _ = self.snap_model(new_x, 0.0)
            dx = snapped_x - base_x
            if abs(dx) < 1e-9: return
            self._move_connected_segments(seg, dx=dx, dy=0.0)
            self._orthogonalize_segments()
            self._ensure_inlet_connected()
            seg.drag_start_model = (cur_mx, start_my)
        else:
            dy_raw = cur_my - start_my
            base_y = (seg.my1 + seg.my2) / 2.0
            new_y = base_y + dy_raw
            _, snapped_y = self.snap_model(0.0, new_y)
            dy = snapped_y - base_y
            if abs(dy) < 1e-9: return
            self._move_connected_segments(seg, dx=0.0, dy=dy)
            self._orthogonalize_segments()
            self._ensure_inlet_connected()
            seg.drag_start_model = (start_mx, cur_my)

        self.redraw_all()

    def on_left_release(self, event):
        if self.mode == "pencil" and self._drawing and self._draw_start is not None:
            mx, my = self.screen_to_model(event.x, event.y)
            smx, smy = self.snap_model(mx, my)
            x1, y1 = self._draw_start
            x2, y2 = smx, smy
            if abs(x2 - x1) >= abs(y2 - y1):
                y2 = y1; vertical = False
            else:
                x2 = x1; vertical = True
            seg = DuctSegment(x1, y1, x2, y2, "", 0, 0, 0.0, vertical)
            self.push_undo()
            self.segments.append(seg)
            if self._preview_line_id is not None:
                self.canvas.delete(self._preview_line_id)
                self._preview_line_id = None
            self._drawing = False
            self._draw_start = None
            self._orthogonalize_segments()
            self._ensure_inlet_connected()
            self.redraw_all()
            self._notify_points_changed()
            return

        if self.dragging_segment is not None:
            self.dragging_segment.is_dragging = False
            self.dragging_segment.drag_start_model = None
            self.dragging_segment = None
        self.on_mouse_move(event)

    def auto_complete(self, dp: float, use_fixed: bool, fixed_val: float, aspect_r: float):
        # record state for undo
        self.push_undo()
        if not self.segments: return
        # 1. 스냅
        eps = 1e-6
        def key_of(x, y, eps=eps): return (round(x/eps)*eps, round(y/eps)*eps)
        reps = {}
        for seg in self.segments:
            for (x, y) in ((seg.mx1, seg.my1), (seg.mx2, seg.my2)):
                k = key_of(x, y)
                if k not in reps: reps[k] = (x, y)
        for seg in self.segments:
            k1 = key_of(seg.mx1, seg.my1); k2 = key_of(seg.mx2, seg.my2)
            seg.mx1, seg.my1 = reps.get(k1, (seg.mx1, seg.my1))
            seg.mx2, seg.my2 = reps.get(k2, (seg.mx2, seg.my2))

        self._orthogonalize_segments()

        # 1.5 Split intersections so crossing segments create branch nodes
        self._split_intersections()

        # 2. Outlet 연결 (Auto-Branch)
        outlet_points = [p for p in self.points if getattr(p, 'kind', None) == 'outlet']
        duct_endpoints = []
        for seg in self.segments:
            duct_endpoints.append((seg.mx1, seg.my1))
            duct_endpoints.append((seg.mx2, seg.my2))

        for outlet in outlet_points:
            ox, oy = outlet.mx, outlet.my
            already_connected = False
            for seg in self.segments:
                if (abs(seg.mx1 - ox) < eps and abs(seg.my1 - oy) < eps) or (abs(seg.mx2 - ox) < eps and abs(seg.my2 - oy) < eps):
                    already_connected = True
                    break
            if already_connected: continue
            
            min_dist = float('inf')
            nearest = None
            for dx, dy in duct_endpoints:
                dist = (dx - ox)**2 + (dy - oy)**2
                if dist < min_dist: min_dist = dist; nearest = (dx, dy)
            if nearest is None: continue
            
            if abs(nearest[0] - ox) >= abs(nearest[1] - oy):
                mid_x, mid_y = ox, nearest[1]
            else:
                mid_x, mid_y = nearest[0], oy
            
            q = getattr(outlet, 'flow', 0.0)
            w, h, label = perform_sizing(q, dp, use_fixed, fixed_val, aspect_r)
            self.segments.append(DuctSegment(nearest[0], nearest[1], mid_x, mid_y, label, w, h, q, vertical_only=(nearest[0]==mid_x)))
            self.segments.append(DuctSegment(mid_x, mid_y, ox, oy, label, w, h, q, vertical_only=(mid_x==ox)))

        # 3. 고아 제거
        def endpoints(seg): return [(seg.mx1, seg.my1), (seg.mx2, seg.my2)]
        deg = {}
        for seg in self.segments:
            for pt in endpoints(seg): deg[pt] = deg.get(pt, 0) + 1
        attached_points = {(p.mx, p.my) for p in self.points}
        def is_attached(pt): return pt in attached_points
        kept = []
        for seg in self.segments:
            pts = endpoints(seg)
            ok = True
            for pt in pts:
                if deg.get(pt, 0) <= 1 and not is_attached(pt):
                    ok = False; break
            if ok: kept.append(seg)
        self.segments = kept
        self._ensure_inlet_connected()
        
        # 4. 리사이징: 각 세그먼트의 실제 유량(아래쪽 outlet 합)을 계산하여 사이징
        # Build node mapping (use rounded keys to avoid float issues)
        def node_key(x, y, eps=1e-6):
            return (round(x/eps)*eps, round(y/eps)*eps)

        # map node -> list of (neighbor_node, segment)
        adj = defaultdict(list)
        seg_by_ends = {}
        for seg in self.segments:
            n1 = node_key(seg.mx1, seg.my1)
            n2 = node_key(seg.mx2, seg.my2)
            adj[n1].append((n2, seg))
            adj[n2].append((n1, seg))
            seg_by_ends[(n1, n2)] = seg
            seg_by_ends[(n2, n1)] = seg

        # locate inlet node
        inlet_node = None
        if self.points:
            inlet = self.points[0]
            inlet_node = node_key(inlet.mx, inlet.my)

        # accumulate flows per segment by finding path from each outlet to inlet
        seg_flow_acc = defaultdict(float)
        outlets = [p for p in self.points if getattr(p, 'kind', None) == 'outlet' and getattr(p, 'flow', 0.0) > 0]
        for out in outlets:
            start = node_key(out.mx, out.my)
            if inlet_node is None or start not in adj:
                continue
            # BFS to find path to inlet
            q = deque([start])
            parent = {start: None}
            parent_seg = {}
            found = False
            while q:
                cur = q.popleft()
                if cur == inlet_node:
                    found = True
                    break
                for (nbr, seg) in adj.get(cur, []):
                    if nbr in parent: continue
                    parent[nbr] = cur
                    parent_seg[nbr] = seg
                    q.append(nbr)
            if not found:
                continue
            # walk back from inlet to start, adding flow to each segment encountered
            node = inlet_node
            while True:
                prev = parent.get(node)
                if prev is None:
                    break
                seg = parent_seg.get(node)
                if seg is not None:
                    seg_flow_acc[seg] += out.flow
                node = prev

        # assign accumulated flows and compute sizing
        for seg in self.segments:
            f = seg_flow_acc.get(seg, 0.0)
            seg.flow = f
            w, h, label = perform_sizing(f, dp, use_fixed, fixed_val, aspect_r)
            seg.duct_w_mm = w; seg.duct_h_mm = h; seg.label_text = label
        self.redraw_all()
        self._notify_points_changed()

    def _split_intersections(self):
        # Split orthogonal segment intersections to create true graph nodes
        if not self.segments: return
        eps = 1e-9
        segs = list(self.segments)
        result = []
        for a in segs:
            split_points = [(a.mx1, a.my1), (a.mx2, a.my2)]
            for b in segs:
                if a is b: continue
                if a.vertical_only == b.vertical_only: continue
                if a.vertical_only:
                    vx = a.mx1
                    vy0, vy1 = min(a.my1, a.my2), max(a.my1, a.my2)
                    hy = b.my1
                    hx0, hx1 = min(b.mx1, b.mx2), max(b.mx1, b.mx2)
                    if (hx0 - eps) <= vx <= (hx1 + eps) and (vy0 - eps) <= hy <= (vy1 + eps):
                        split_points.append((vx, hy))
                else:
                    vx = b.mx1
                    vy0, vy1 = min(b.my1, b.my2), max(b.my1, b.my2)
                    hy = a.my1
                    hx0, hx1 = min(a.mx1, a.mx2), max(a.mx1, a.mx2)
                    if (hx0 - eps) <= vx <= (hx1 + eps) and (vy0 - eps) <= hy <= (vy1 + eps):
                        split_points.append((vx, hy))
            # unique and sorted
            if a.vertical_only:
                pts = sorted(set(split_points), key=lambda p: p[1])
            else:
                pts = sorted(set(split_points), key=lambda p: p[0])
            if len(pts) <= 1:
                continue
            # reuse original object for first piece
            x1,y1 = pts[0]
            x2,y2 = pts[1]
            a.mx1, a.my1, a.mx2, a.my2 = x1, y1, x2, y2
            result.append(a)
            # create additional pieces
            for k in range(1, len(pts)-1):
                xx1,yy1 = pts[k]
                xx2,yy2 = pts[k+1]
                if abs(xx1-xx2) < eps and abs(yy1-yy2) < eps: continue
                new_seg = DuctSegment(xx1, yy1, xx2, yy2, a.label_text, a.duct_w_mm, a.duct_h_mm, a.flow, a.vertical_only)
                result.append(new_seg)
        if result:
            self.segments = result

    def _move_connected_segments(self, base_seg, dx, dy):
        def seg_endpoints(seg): return [(seg.mx1, seg.my1), (seg.mx2, seg.my2)]
        connected = set([base_seg])
        queue = [base_seg]
        while queue:
            cur = queue.pop(0)
            cur_ends = seg_endpoints(cur)
            for other in self.segments:
                if other in connected: continue
                other_ends = seg_endpoints(other)
                if any((abs(ex1 - ex2) < 1e-9 and abs(ey1 - ey2) < 1e-9) for (ex1, ey1) in cur_ends for (ex2, ey2) in other_ends):
                    connected.add(other); queue.append(other)
        point_positions = [(p.mx, p.my) for p in self.points]
        def is_attached_to_point(x, y):
            for px, py in point_positions:
                if abs(px - x) < 1e-9 and abs(py - y) < 1e-9: return True
            return False
        for seg in connected:
            if not is_attached_to_point(seg.mx1, seg.my1): seg.mx1 += dx; seg.my1 += dy
            if not is_attached_to_point(seg.mx2, seg.my2): seg.mx2 += dx; seg.my2 += dy

    def _orthogonalize_segments(self):
        if not self.segments: return
        for seg in self.segments:
            if seg.vertical_only:
                x_avg = (seg.mx1 + seg.mx2) / 2.0
                seg.mx1 = x_avg; seg.mx2 = x_avg
            else:
                y_avg = (seg.my1 + seg.my2) / 2.0
                seg.my1 = y_avg; seg.my2 = y_avg
        def key_of(x, y, eps=1e-6): return (round(x/eps)*eps, round(y/eps)*eps)
        rep = {}
        for seg in self.segments:
            for (x, y) in ((seg.mx1, seg.my1), (seg.mx2, seg.my2)):
                k = key_of(x, y)
                if k not in rep: rep[k] = (x, y)
        for seg in self.segments:
            k1 = key_of(seg.mx1, seg.my1); k2 = key_of(seg.mx2, seg.my2)
            if k1 in rep: seg.mx1, seg.my1 = rep[k1]
            if k2 in rep: seg.mx2, seg.my2 = rep[k2]

    def _ensure_inlet_connected(self):
        if not self.points: return
        inlet = self.points[0]
        if inlet.kind != "inlet": return
        ix, iy = inlet.mx, inlet.my
        for seg in self.segments:
            if (abs(seg.mx1 - ix) < 1e-9 and abs(seg.my1 - iy) < 1e-9) or (abs(seg.mx2 - ix) < 1e-9 and abs(seg.my2 - iy) < 1e-9): return
        nearest_seg = None; nearest_dist = float("inf"); proj_point = None
        for seg in self.segments:
            if seg.vertical_only:
                x0 = seg.mx1; y0 = min(seg.my1, seg.my2); y1 = max(seg.my1, seg.my2)
                py = min(max(iy, y0), y1); px = x0
            else:
                y0 = seg.my1; x0 = min(seg.mx1, seg.mx2); x1 = max(seg.mx1, seg.mx2)
                px = min(max(ix, x0), x1); py = y0
            dist = math.hypot(px - ix, py - iy)
            if dist < nearest_dist: nearest_dist = dist; nearest_seg = seg; proj_point = (px, py)
        if nearest_seg is None or proj_point is None: return
        px, py = proj_point
        if nearest_dist < 1e-9: return
        cur_x, cur_y = ix, iy
        if abs(px - cur_x) > 1e-9:
            self.segments.append(DuctSegment(cur_x, cur_y, px, cur_y, "", 0, 0, 0.0, False))
            cur_x = px
        if abs(py - cur_y) > 1e-9:
            self.segments.append(DuctSegment(cur_x, cur_y, cur_x, py, "", 0, 0, 0.0, True))

    def _merge_colinear_neighbors(self):
        # Merge adjacent colinear segments when their connecting node has degree 2
        if not self.segments:
            return
        eps = 1e-9
        def node_key(x, y):
            return (round(x, 9), round(y, 9))

        attached_points = {(p.mx, p.my) for p in self.points}

        changed = True
        while changed:
            changed = False
            # build node -> segments map
            node_map = {}
            for seg in self.segments:
                n1 = node_key(seg.mx1, seg.my1)
                n2 = node_key(seg.mx2, seg.my2)
                node_map.setdefault(n1, []).append(seg)
                node_map.setdefault(n2, []).append(seg)

            # find candidate nodes to collapse
            for node, segs in list(node_map.items()):
                if node in attached_points: continue
                if len(segs) != 2: continue
                s1, s2 = segs[0], segs[1]
                # must be same orientation
                if s1.vertical_only != s2.vertical_only:
                    continue
                # compute endpoints other than node
                def other_end(s, n):
                    if abs(s.mx1 - n[0]) < eps and abs(s.my1 - n[1]) < eps:
                        return (s.mx2, s.my2)
                    else:
                        return (s.mx1, s.my1)

                o1 = other_end(s1, node)
                o2 = other_end(s2, node)

                # ensure colinear alignment (for vertical x equal, for horizontal y equal)
                if s1.vertical_only:
                    if abs(o1[0] - o2[0]) > 1e-6: continue
                    new_seg = DuctSegment(o1[0], o1[1], o2[0], o2[1], "", max(s1.duct_w_mm, s2.duct_w_mm), max(s1.duct_h_mm, s2.duct_h_mm), s1.flow + s2.flow, True)
                else:
                    if abs(o1[1] - o2[1]) > 1e-6: continue
                    new_seg = DuctSegment(o1[0], o1[1], o2[0], o2[1], "", max(s1.duct_w_mm, s2.duct_w_mm), max(s1.duct_h_mm, s2.duct_h_mm), s1.flow + s2.flow, False)

                # remove originals and add new
                try:
                    self.segments.remove(s1)
                    self.segments.remove(s2)
                except ValueError:
                    continue
                self.segments.append(new_seg)
                changed = True
                break

        if changed:
            self._orthogonalize_segments()
            self._ensure_inlet_connected()
            self.redraw_all()
            self._notify_points_changed()

    def draw_duct_network(self, dp_mmAq_per_m: float, use_fixed: bool, fixed_val: float, aspect_ratio: float):
        """종합 사이징: 자동 방향 감지 및 헤더 그룹화 적용"""
        # record state for undo
        self.push_undo()

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
            messagebox.showwarning("경고", "Air inlet 풍량 확인 필요.")
            return

        outlets = [p for p in self.points[1:] if p.kind == "outlet" and p.flow > 0]
        
        def create_seg(x1, y1, x2, y2, flow, is_vert):
            if abs(x1 - x2) < 1e-6 and abs(y1 - y2) < 1e-6: return 
            w, h, label = perform_sizing(flow, dp_mmAq_per_m, use_fixed, fixed_val, aspect_ratio)
            self.segments.append(DuctSegment(x1, y1, x2, y2, label, w, h, flow, is_vert))

        # 1. 방향 자동 감지 (Vertical or Horizontal Main?)
        sum_dx = sum(abs(p.mx - inlet.mx) for p in outlets)
        sum_dy = sum(abs(p.my - inlet.my) for p in outlets)
        is_vertical_main = sum_dy > sum_dx

        # 2. 그룹핑 (Groups)
        groups = {}
        tol = 0.1

        if is_vertical_main:
            # 수직 메인 -> X좌표가 비슷한 것끼리 그룹핑
            for ot in outlets:
                found_key = None
                for k in groups.keys():
                    if abs(k - ot.mx) < tol: found_key = k; break
                if found_key is None: groups[ot.mx] = [ot]
                else: groups[found_key].append(ot)
        else:
            # 수평 메인 -> Y좌표가 비슷한 것끼리 그룹핑
            for ot in outlets:
                found_key = None
                for k in groups.keys():
                    if abs(k - ot.my) < tol: found_key = k; break
                if found_key is None: groups[ot.my] = [ot]
                else: groups[found_key].append(ot)

        # 3. 메인 노드(Main Nodes) 생성
        main_nodes = [] 

        if is_vertical_main:
            # 메인 축: Inlet X (x_main)
            for grp_x, grp_outlets in groups.items():
                grp_flow = sum(o.flow for o in grp_outlets)
                # 분기점(Take-off): 그룹 내에서 Inlet Y와 가장 가까운 Y
                takeoff_y = min(grp_outlets, key=lambda o: abs(o.my - inlet.my)).my
                main_nodes.append({'pos': takeoff_y, 'flow': grp_flow, 'grp_coord': grp_x, 'outlets': grp_outlets})
            
            if not main_nodes: return
            x_main = inlet.mx
            
            up_nodes = [n for n in main_nodes if n['pos'] < inlet.my - 1e-6]
            down_nodes = [n for n in main_nodes if n['pos'] > inlet.my + 1e-6]
            
            # Upstream Main
            if up_nodes:
                up_nodes.sort(key=lambda n: n['pos'], reverse=True)
                curr_y = inlet.my
                current_main_flow = sum(n['flow'] for n in up_nodes)
                create_seg(x_main, curr_y, x_main, up_nodes[0]['pos'], current_main_flow, True)
                for i in range(len(up_nodes) - 1):
                    current_main_flow -= up_nodes[i]['flow']
                    if current_main_flow <= 0: break
                    create_seg(x_main, up_nodes[i]['pos'], x_main, up_nodes[i+1]['pos'], current_main_flow, True)

            # Downstream Main
            if down_nodes:
                down_nodes.sort(key=lambda n: n['pos'])
                curr_y = inlet.my
                current_main_flow = sum(n['flow'] for n in down_nodes)
                create_seg(x_main, curr_y, x_main, down_nodes[0]['pos'], current_main_flow, True)
                for i in range(len(down_nodes) - 1):
                    current_main_flow -= down_nodes[i]['flow']
                    if current_main_flow <= 0: break
                    create_seg(x_main, down_nodes[i]['pos'], x_main, down_nodes[i+1]['pos'], current_main_flow, True)

            # Branches & Headers
            for node in main_nodes:
                takeoff_y = node['pos']
                grp_x = node['grp_coord']
                grp_outlets = node['outlets']
                
                # Riser (Main -> Header)
                if abs(x_main - grp_x) > 1e-6:
                    create_seg(x_main, takeoff_y, grp_x, takeoff_y, node['flow'], False)
                
                # Header (Group Distribution)
                sorted_outs = sorted(grp_outlets, key=lambda o: o.my)
                h_up = [o for o in sorted_outs if o.my < takeoff_y - 1e-6]
                if h_up:
                    h_up.sort(key=lambda o: o.my, reverse=True)
                    curr_h_y = takeoff_y
                    h_flow = sum(o.flow for o in h_up)
                    create_seg(grp_x, curr_h_y, grp_x, h_up[0].my, h_flow, True)
                    for i in range(len(h_up)-1):
                        h_flow -= h_up[i].flow
                        create_seg(grp_x, h_up[i].my, grp_x, h_up[i+1].my, h_flow, True)

                h_down = [o for o in sorted_outs if o.my > takeoff_y + 1e-6]
                if h_down:
                    h_down.sort(key=lambda o: o.my)
                    curr_h_y = takeoff_y
                    h_flow = sum(o.flow for o in h_down)
                    create_seg(grp_x, curr_h_y, grp_x, h_down[0].my, h_flow, True)
                    for i in range(len(h_down)-1):
                        h_flow -= h_down[i].flow
                        create_seg(grp_x, h_down[i].my, grp_x, h_down[i+1].my, h_flow, True)

        else:
            # 메인 축: Inlet Y (y_main)
            for grp_y, grp_outlets in groups.items():
                grp_flow = sum(o.flow for o in grp_outlets)
                # 분기점: 그룹 내에서 Inlet X와 가장 가까운 X
                takeoff_x = min(grp_outlets, key=lambda o: abs(o.mx - inlet.mx)).mx
                main_nodes.append({'pos': takeoff_x, 'flow': grp_flow, 'grp_coord': grp_y, 'outlets': grp_outlets})

            y_main = inlet.my
            
            left_nodes = [n for n in main_nodes if n['pos'] < inlet.mx - 1e-6]
            right_nodes = [n for n in main_nodes if n['pos'] > inlet.mx + 1e-6]
            
            if left_nodes:
                left_nodes.sort(key=lambda n: n['pos'], reverse=True)
                curr_x = inlet.mx
                current_main_flow = sum(n['flow'] for n in left_nodes)
                create_seg(curr_x, y_main, left_nodes[0]['pos'], y_main, current_main_flow, False)
                for i in range(len(left_nodes) - 1):
                    current_main_flow -= left_nodes[i]['flow']
                    if current_main_flow <= 0: break
                    create_seg(left_nodes[i]['pos'], y_main, left_nodes[i+1]['pos'], y_main, current_main_flow, False)

            if right_nodes:
                right_nodes.sort(key=lambda n: n['pos'])
                curr_x = inlet.mx
                current_main_flow = sum(n['flow'] for n in right_nodes)
                create_seg(curr_x, y_main, right_nodes[0]['pos'], y_main, current_main_flow, False)
                for i in range(len(right_nodes) - 1):
                    current_main_flow -= right_nodes[i]['flow']
                    if current_main_flow <= 0: break
                    create_seg(right_nodes[i]['pos'], y_main, right_nodes[i+1]['pos'], y_main, current_main_flow, False)

            for node in main_nodes:
                takeoff_x = node['pos']
                grp_y = node['grp_coord']
                grp_outlets = node['outlets']
                
                if abs(y_main - grp_y) > 1e-6:
                    create_seg(takeoff_x, y_main, takeoff_x, grp_y, node['flow'], True)
                
                sorted_outs = sorted(grp_outlets, key=lambda o: o.mx)
                h_left = [o for o in sorted_outs if o.mx < takeoff_x - 1e-6]
                if h_left:
                    h_left.sort(key=lambda o: o.mx, reverse=True)
                    curr_h_x = takeoff_x
                    h_flow = sum(o.flow for o in h_left)
                    create_seg(curr_h_x, grp_y, h_left[0].mx, grp_y, h_flow, False)
                    for i in range(len(h_left)-1):
                        h_flow -= h_left[i].flow
                        create_seg(h_left[i].mx, grp_y, h_left[i+1].mx, grp_y, h_flow, False)

                h_right = [o for o in sorted_outs if o.mx > takeoff_x + 1e-6]
                if h_right:
                    h_right.sort(key=lambda o: o.mx)
                    curr_h_x = takeoff_x
                    h_flow = sum(o.flow for o in h_right)
                    create_seg(curr_h_x, grp_y, h_right[0].mx, grp_y, h_flow, False)
                    for i in range(len(h_right)-1):
                        h_flow -= h_right[i].flow
                        create_seg(h_right[i].mx, grp_y, h_right[i+1].mx, grp_y, h_flow, False)

        self.redraw_all()
        self._notify_points_changed()

    def undo_last_point(self):
        # backward compatibility: perform a full undo
        self.undo()

    def clear_all(self):
        self.push_undo()
        self.points.clear()
        self.segments.clear()
        self.redraw_all()
        self._notify_points_changed()

    def distribute_equal_flow(self):
        if len(self.points) < 2:
            messagebox.showwarning("경고", "최소 2개 이상의 점 필요.")
            return
        self.push_undo()
        Q_in = self.inlet_flow
        n_out = len(self.points) - 1
        Q_each = Q_in / n_out
        for idx, p in enumerate(self.points):
            if idx == 0: p.flow = Q_in
            else: p.flow = Q_each
        self.segments.clear()
        self.redraw_all()
        self._notify_points_changed()


# =========================
# 3. GUI 이벤트 처리
# =========================

# Outlet relative-position statistics and cursor tooltip removed per user request.
# The functions and widgets that displayed outlet relative statistics and the
# mouse-over tooltip were intentionally deleted to simplify the UI.


def update_sheet_area(pal: Palette):
    """Calculate total duct sheet area (m2) from palette segments and update results widget."""
    global results_text_widget
    if results_text_widget is None: return
    total_area_m2 = 0.0
    for seg in getattr(pal, 'segments', []):
        L = seg.length_m()
        w_m = getattr(seg, 'duct_w_mm', 0) / 1000.0
        h_m = getattr(seg, 'duct_h_mm', 0) / 1000.0
        area = (w_m + h_m) * 2 * L
        total_area_m2 += area

    # append a new numbered history line showing the latest sheet area
    results_text_widget.config(state="normal")
    existing = results_text_widget.get("1.0", "end").rstrip()
    # count existing numbered lines like 'N. '
    num = 0
    if existing:
        for ln in existing.splitlines():
            s = ln.lstrip()
            if not s: continue
            # check prefix like '1.' or '12.'
            parts = s.split(None, 1)
            if parts:
                prefix = parts[0]
                if prefix.endswith('.'):
                    try:
                        int(prefix[:-1])
                        num += 1
                    except Exception:
                        pass
    next_idx = num + 1
    base = existing + "\n" if existing else ""
    base += f"{next_idx}. 덕트 철판 소요량 (m²) : {total_area_m2:.1f}"
    results_text_widget.delete("1.0", "end")
    results_text_widget.insert("end", base)
    results_text_widget.config(state="disabled")
    try:
        results_text_widget.see("end")
    except Exception:
        try: results_text_widget.yview_moveto(1.0)
        except: pass

def get_sizing_params():
    try: dp = float(resistance_entry.get())
    except: dp = 0.1
    use_fixed = fixed_side_var.get()
    fixed_val = 0.0
    if use_fixed:
        try: fixed_val = float(fixed_side_entry.get())
        except: fixed_val = 0.0
    try: r = float(aspect_ratio_combo.get())
    except: r = 2.0
    return dp, use_fixed, fixed_val, r


def format_cubic_meter_entry(event=None):
    """Format `cubic_meter_hour_entry` value with thousand separators while typing."""
    global cubic_meter_hour_entry
    try:
        s = cubic_meter_hour_entry.get()
    except Exception:
        return
    if s is None: return
    # remove existing commas
    raw = s.replace(',', '')
    if raw in ('', '-', '.', '-.'): return
    neg = raw.startswith('-')
    if neg:
        raw2 = raw[1:]
    else:
        raw2 = raw
    parts = raw2.split('.')
    int_part = parts[0] if parts[0] != '' else '0'
    try:
        intval = int(int_part)
    except Exception:
        return
    int_fmt = f"{intval:,}"
    if neg: int_fmt = '-' + int_fmt
    if len(parts) > 1:
        frac = parts[1]
        new = int_fmt + '.' + frac
    else:
        new = int_fmt

    if new != s:
        # update entry and move cursor to end for simplicity
        cubic_meter_hour_entry.delete(0, 'end')
        cubic_meter_hour_entry.insert(0, new)
        try: cubic_meter_hour_entry.icursor('end')
        except: pass

def calculate():
    try:
        q = float(cubic_meter_hour_entry.get().replace(',', ''))
        dp, use_fixed, fixed_val, r = get_sizing_params()

        D1 = calc_circular_diameter(q, dp)
        D2 = round_step_up(D1, 50)

        w, h, label = perform_sizing(q, dp, use_fixed, fixed_val, r)

        if use_fixed:
            text = (
                f"- 원형덕트 (이론치) : {D1:.0f}\n"
                f"- 원형덕트(규격화) : {D2}\n"
                f"- 사각덕트(이론치) : {w} X {h}\n"
                f"※ 고정 변 모드 적용 중"
            )
        else:
            _, _, _, theo_big, theo_small = size_rect_from_D1(D1, r, 50)
            text = (
                f"- 원형덕트 (이론치) : {D1:.0f}\n"
                f"- 원형덕트(규격화) : {D2}\n"
                f"- 사각덕트(이론치) : {w} X {h}\n"
                f"- 사각덕트 (규격화) : {theo_big:.1f} X {theo_small:.1f}\n"
            )

        text = "[덕트 사이즈 결과]\n" + text + f"\n※ 팔레트 격자 1칸 = 0.5 m"
        results_text_widget.config(state="normal")
        results_text_widget.delete("1.0", "end")
        results_text_widget.insert("end", text)
        results_text_widget.config(state="disabled")

        # update inlet flow in palette
        try:
            palette.set_inlet_flow(q)
        except Exception:
            pass

    except ValueError as e:
        messagebox.showerror("입력 오류", f"입력값을 확인하세요!\n\n{e}")
    except Exception as e:
        messagebox.showerror("알 수 없는 오류", f"알 수 없는 오류:\n{e}")


def total_sizing():
    dp, use_fixed, fixed_val, r = get_sizing_params()
    if use_fixed and fixed_val <= 0:
        messagebox.showwarning("경고", "고정 변 길이(mm)를 올바르게 입력하세요.")
        return
    palette.draw_duct_network(dp, use_fixed, fixed_val, r)

    total_area_m2 = 0.0
    for seg in palette.segments:
        L = seg.length_m()
        w_m = seg.duct_w_mm / 1000.0
        h_m = seg.duct_h_mm / 1000.0
        area = (w_m + h_m) * 2 * L
        total_area_m2 += area

    results_text_widget.config(state="normal")
    base = results_text_widget.get("1.0", "end").rstrip()
    if base: base += "\n"
    base += f"5. 덕트 철판 소요량 (m²) : {total_area_m2:.1f}"
    results_text_widget.delete("1.0", "end")
    results_text_widget.insert("end", base)
    results_text_widget.config(state="disabled")


def auto_complete_action():
    dp, use_fixed, fixed_val, r = get_sizing_params()
    if use_fixed and fixed_val <= 0:
        messagebox.showwarning("경고", "고정 변 길이(mm)를 올바르게 입력하세요.")
        return
    palette.auto_complete(dp, use_fixed, fixed_val, r)


def clear_palette(): palette.clear_all()

def equal_distribution():
    try:
        q = float(cubic_meter_hour_entry.get().replace(',', ''))
        palette.set_inlet_flow(q)
    except:
        messagebox.showerror("입력 오류", "풍량 값을 확인하세요.")
        return
    palette.distribute_equal_flow()

def undo_point(): palette.undo_last_point()

def toggle_fixed_side():
    if fixed_side_var.get():
        aspect_ratio_combo.config(state="disabled")
        fixed_side_entry.config(state="normal")
    else:
        aspect_ratio_combo.config(state="readonly")
        fixed_side_entry.config(state="disabled")

# =========================
# 4. GUI 레이아웃 구성
# =========================

def create_app():
    root = tk.Tk()
    root.title("스마트 덕트 사이징 프로그램 (v3.0 - Auto Orientation & Grouping)")

    main_frame = tk.Frame(root)
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)
    root.bind("<Control-z>", lambda event: undo_point())
    # expose commonly used widgets as module-level globals so callbacks can access them
    global cubic_meter_hour_entry, resistance_entry, aspect_ratio_combo, fixed_side_var, fixed_side_entry
    global results_text_widget, relpos_text_widget, palette

    # 좌측 정보 입력창
    info_frame = tk.Frame(main_frame, width=180)
    info_frame.pack(side="left", fill="y", padx=(0,10))

    tk.Label(info_frame, text="외기/실내/급기/발열량", font=("Arial", 10, "bold")).pack(anchor="w", pady=(6,4))
    labels = ["외기온도 (°C):", "실내온도 (°C):", "급기온도 (°C):", "일반 발열량 (W/m²):", "장비 발열량 (W/m²):"]
    defaults = ["-5.0", "25.0", "18.0", "0.00", "0.00"]
    for lbl, dft in zip(labels, defaults):
        tk.Label(info_frame, text=lbl).pack(anchor="w")
        e = tk.Entry(info_frame, width=10)
        e.pack(anchor="w", pady=2)
        e.insert(0, dft)

    left_frame = tk.Frame(main_frame)
    # 왼쪽 컨트롤을 윈도우 세로 상단에 정렬
    left_frame.pack(side="left", anchor="n", fill="y")

    right_frame = tk.Frame(main_frame, bg="#f5f5f5", bd=1, relief="solid")
    right_frame.configure(width=700)
    right_frame.pack(side="right", fill="both", expand=True)

    # 제어 패널
    row_idx = 0
    tk.Label(left_frame, text="풍량 (m³/h):").grid(row=row_idx, column=0, padx=5, pady=5, sticky="w")
    cubic_meter_hour_entry = tk.Entry(left_frame, width=10)
    cubic_meter_hour_entry.grid(row=row_idx, column=1, padx=5, pady=5, sticky="w")
    cubic_meter_hour_entry.insert(0, "50000")
    # format while typing and on focus out
    cubic_meter_hour_entry.bind('<KeyRelease>', lambda e: format_cubic_meter_entry(e))
    cubic_meter_hour_entry.bind('<FocusOut>', lambda e: format_cubic_meter_entry(e))
    # format initial value
    try: format_cubic_meter_entry(None)
    except: pass
    row_idx += 1

    tk.Label(left_frame, text="정압값 (mmAq/m):").grid(row=row_idx, column=0, padx=5, pady=5, sticky="w")
    resistance_entry = tk.Entry(left_frame, width=10)
    resistance_entry.grid(row=row_idx, column=1, padx=5, pady=5, sticky="w")
    resistance_entry.insert(0, "0.1")
    row_idx += 1

    tk.Label(left_frame, text="사각 덕트 종횡비 (b/a):").grid(row=row_idx, column=0, padx=5, pady=5, sticky="w")
    aspect_ratio_combo = ttk.Combobox(left_frame, values=["1", "2", "3", "4"], state="readonly", width=5)
    aspect_ratio_combo.current(1)
    aspect_ratio_combo.grid(row=row_idx, column=1, padx=5, pady=5, sticky="w")
    row_idx += 1

    # 한 변 고정 옵션
    fixed_side_var = tk.BooleanVar(value=False)
    fixed_chk = tk.Checkbutton(left_frame, text="한 변 고정(mm):", variable=fixed_side_var, command=toggle_fixed_side)
    fixed_chk.grid(row=row_idx, column=0, padx=5, pady=5, sticky="w")

    fixed_side_entry = tk.Entry(left_frame, width=10, state="disabled")
    fixed_side_entry.grid(row=row_idx, column=1, padx=5, pady=5, sticky="w")
    fixed_side_entry.insert(0, "300")
    row_idx += 1

    # 버튼들 - 두 개씩 가로로 배치
    tk.Button(left_frame, text="계산하기", command=calculate).grid(row=row_idx, column=0, padx=5, pady=5, sticky="w")
    tk.Button(left_frame, text="균등 풍량 배분", command=equal_distribution).grid(row=row_idx, column=1, padx=5, pady=5, sticky="w")
    row_idx += 1
    tk.Button(left_frame, text="종합 사이징", command=total_sizing).grid(row=row_idx, column=0, padx=5, pady=5, sticky="w")
    tk.Button(left_frame, text="팔레트 전체 지우기", command=clear_palette).grid(row=row_idx, column=1, padx=5, pady=5, sticky="w")
    row_idx += 1

    tk.Button(left_frame, text="펜슬 모드", command=lambda: palette.set_mode_pencil()).grid(row=row_idx, column=0, padx=5, pady=5, sticky="w")
    tk.Button(left_frame, text="자동완성", command=auto_complete_action).grid(row=row_idx, column=1, padx=5, pady=5, sticky="w")
    row_idx += 1

    # 결과 출력창
    results_frame = tk.Frame(left_frame)
    results_frame.grid(row=row_idx, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
    left_frame.grid_rowconfigure(row_idx, weight=1)
    results_text_widget = tk.Text(results_frame, width=36, height=8, wrap="word", bg="white", relief="solid")
    results_text_widget.pack(side="left", fill="both", expand=True)
    results_scrollbar = tk.Scrollbar(results_frame, orient="vertical", command=results_text_widget.yview)
    results_scrollbar.pack(side="right", fill="y")
    results_text_widget.configure(yscrollcommand=results_scrollbar.set)
    results_text_widget.config(state="disabled")
    row_idx += 1

    # 팔레트 초기화
    global palette
    palette = Palette(right_frame)
    # Outlet relative-position statistics removed; no callback for it
    palette.sheet_changed_callback = update_sheet_area
    update_sheet_area(palette)

    # 시작 시: 전체 창 크기를 현재 값에서 가로 +20%, 세로 +20% 만큼 늘리고, 그 증가분을 팔레트(right_frame)에 할당
    try:
        root.update_idletasks()
        base_w = root.winfo_width() or root.winfo_reqwidth() or 1000
        base_h = root.winfo_height() or root.winfo_reqheight() or 700

        # 팔레트(right_frame)만 가로/세로 각각 20% 증가시키기
        try:
            curr_rf_w = right_frame.winfo_width()
            curr_rf_h = right_frame.winfo_height()
            if not curr_rf_w or curr_rf_w <= 1:
                curr_rf_w = 700
            if not curr_rf_h or curr_rf_h <= 1:
                curr_rf_h = 500
        except Exception:
            curr_rf_w = 700
            curr_rf_h = 500

        try:
            new_w = int(curr_rf_w * 1.2)
            new_h = int(curr_rf_h * 1.2)
            right_frame.configure(width=new_w, height=new_h)
            try:
                # prevent pack from shrinking the frame back to children size
                right_frame.pack_propagate(False)
            except Exception:
                pass
            try:
                # resize the canvas inside the palette to match (subtract padding)
                pad_x = 20
                pad_y = 20
                palette.canvas.config(width=max(10, new_w - pad_x), height=max(10, new_h - pad_y))
            except Exception:
                pass
        except Exception:
            try:
                right_frame.configure(width=int(curr_rf_w * 1.2))
            except Exception:
                pass
    except Exception:
        pass

    root.mainloop()


if __name__ == "__main__":
    create_app()