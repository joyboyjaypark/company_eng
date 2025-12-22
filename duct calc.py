import tkinter as tk
from tkinter import messagebox, ttk, simpledialog
import copy
from collections import deque, defaultdict
import math

# =========================
# 1. 계산 함수들 (Engineering Logic)
# =========================

def calc_circular_diameter(q_m3h: float, dp_mmAq_per_m: float) -> float:
    if q_m3h <= 0:
        return 0.0
    if dp_mmAq_per_m <= 0:
        raise ValueError("정압값(mmAq/m)은 0보다 커야 합니다.")
    C = 3.295e-10
    D = ((C * q_m3h**1.9 / dp_mmAq_per_m)**0.199) * 1000
    return round(D, 0)


def round_step_up(x: float, step: float = 50) -> float:
    return math.ceil(x / step) * step


def round_step_down(x: float, step: float = 50) -> float:
    return math.floor(x / step) * step


def rect_equiv_diameter(a_mm: float, b_mm: float) -> float:
    if a_mm <= 0 or b_mm <= 0:
        return 0.0
    a, b = float(a_mm), float(b_mm)
    return 1.30 * (a*b)**0.625 / (a + b)**0.25


def size_rect_from_D1(D1: float, aspect_ratio: float, step: float = 50):
    if D1 <= 0:
        return 0, 0, 0.0, 0.0, 0.0
    if aspect_ratio <= 0:
        raise ValueError("종횡비(b/a)는 0보다 커야 합니다.")

    De_target = float(D1)
    r = float(aspect_ratio)

    a_theo = De_target * (1 + r)**0.25 / (1.30 * r**0.625)
    b_theo = r * a_theo
    theo_big, theo_small = max(a_theo, b_theo), min(a_theo, b_theo)

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
    if D1 <= 0:
        return 0, 0, 0.0
    if fixed_side_mm <= 0:
        raise ValueError("고정 변의 길이는 0보다 커야 합니다.")
    
    fixed = round_step_up(fixed_side_mm, step)
    other = 50.0
    while True:
        de = rect_equiv_diameter(fixed, other)
        if de >= D1 - 0.1:
            break
        other += step
        if other > 10000:
            break
            
    sel_big = max(fixed, other)
    sel_small = min(fixed, other)
    return int(sel_big), int(sel_small), round(de, 1)


def perform_sizing(q: float, dp: float, use_fixed: bool, fixed_val: float, aspect_r: float):
    if q <= 0:
        return 0, 0, f"0x0\n(0 m³/h)"
    try:
        D1 = calc_circular_diameter(q, dp)
        if use_fixed:
            w, h, de = calc_rect_other_side(D1, fixed_val, 50)
        else:
            w, h, de, theo_big, theo_small = size_rect_from_D1(D1, aspect_r, 50)
        size_line = f"{w}x{h}"
        flow_fmt = f"({int(round(q)):,} m³/h)"
        label = f"{size_line}\n{flow_fmt}"
        return w, h, label
    except Exception:
        try:
            flow_fmt = f"({int(round(q)):,} m³/h)"
        except Exception:
            flow_fmt = "(0 m³/h)"
        return 0, 0, f"0x0\n{flow_fmt}"


# =========================
# 2. 데이터 모델 및 팔레트 클래스
# =========================

GRID_STEP_MODEL = 0.5
INITIAL_SCALE = 40.0
velocity_check_var = None


class AirPoint:
    def __init__(self, mx, my, kind, flow):
        self.mx = mx
        self.my = my
        self.kind = kind
        self.flow = flow
        self.canvas_id = None
        self.text_id = None
        self.number_id = None


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
        self.is_hovered = False
        self.is_dragging = False
        self.drag_start_model = None

    def length_m(self):
        dx = self.mx2 - self.mx1
        dy = self.my2 - self.my1
        return math.sqrt(dx*dx + dy*dy)


class Palette:
    def __init__(self, parent):
        self.container = tk.Frame(parent)
        self.container.pack(expand=True, fill="both", padx=10, pady=10)

        self.ruler_top = tk.Canvas(self.container, height=24, bg="#f0f0f0", highlightthickness=0)
        self.ruler_left = tk.Canvas(self.container, width=40, bg="#f0f0f0", highlightthickness=0)
        self.canvas = tk.Canvas(self.container, bg="white")

        spacer = tk.Frame(self.container, width=40, height=24, bg="#f0f0f0")
        spacer.grid(row=0, column=0, sticky="nsew")
        self.ruler_top.grid(row=0, column=1, sticky="nsew")
        self.ruler_left.grid(row=1, column=0, sticky="nsew")
        self.canvas.grid(row=1, column=1, sticky="nsew")

        self.container.grid_rowconfigure(1, weight=1)
        self.container.grid_columnconfigure(1, weight=1)

        self.points = []
        self.inlet_flow = 0.0
        self.scale_factor = INITIAL_SCALE
        self.offset_x = 0.0
        self.offset_y = 0.0
        self.grid_tag = "grid"
        self.pan_start_screen = None
        self.segments = []

        self.mode = "pan"
        self._drawing = False
        self._draw_start = None
        self._preview_line_id = None

        self.hovered_segment = None
        self.dragging_segment = None
        self.hovered_text_id = None
        self.current_mouse_model = None

        self.points_changed_callback = None
        self._undo_stack = []
        self._undo_limit = 100

        self.canvas.bind("<Button-1>", self.on_left_click)
        self.canvas.bind("<Button-3>", self.on_right_click)
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        self.canvas.bind("<Button-2>", self.on_middle_press)
        self.canvas.bind("<B2-Motion>", self.on_middle_drag)
        self.canvas.bind("<Configure>", self.on_resize)
        self.ruler_top.bind("<Configure>", self.on_resize)
        self.ruler_left.bind("<Configure>", self.on_resize)
        self.canvas.bind("<Motion>", self.on_mouse_move)
        self.canvas.bind("<Leave>", self.on_mouse_leave)
        self.canvas.bind("<B1-Motion>", self.on_left_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_release)

        self.redraw_all()

    def draw_rulers(self):
        try:
            cw = self.canvas.winfo_width()
            ch = self.canvas.winfo_height()
            if cw <= 0 or ch <= 0:
                return

            self.ruler_top.delete("all")
            self.ruler_left.delete("all")

            mx_min, _ = self.screen_to_model(0, 0)
            mx_max, _ = self.screen_to_model(cw, 0)
            _, my_min = self.screen_to_model(0, 0)
            _, my_max = self.screen_to_model(0, ch)

            if self.scale_factor * 0.5 >= 8:
                minor = 0.5
            else:
                minor = 1.0
            major = 1.0 if minor == 0.5 else 5.0

            start = math.floor(mx_min / minor) * minor
            x = start
            while x <= mx_max:
                px = x * self.scale_factor + self.offset_x
                if abs((x / major) - round(x / major)) < 1e-6:
                    self.ruler_top.create_line(px, 24, px, 15, fill="black")
                    label = f"{x:.0f}" if major >= 1 else f"{x:.1f}"
                    self.ruler_top.create_text(px + 2, 2, text=label, anchor="nw", font=("Arial", 8))
                else:
                    self.ruler_top.create_line(px, 24, px, 18, fill="black")
                x += minor

            start_y = math.floor(my_min / minor) * minor
            y = start_y
            while y <= my_max:
                py = y * self.scale_factor + self.offset_y
                if abs((y / major) - round(y / major)) < 1e-6:
                    self.ruler_left.create_line(40, py, 26, py, fill="black")
                    label = f"{y:.0f}" if major >= 1 else f"{y:.1f}"
                    self.ruler_left.create_text(2, py - 6, text=label, anchor="nw", font=("Arial", 8))
                else:
                    self.ruler_left.create_line(40, py, 30, py, fill="black")
                y += minor
        except Exception:
            pass

    def _notify_points_changed(self):
        cb = getattr(self, "points_changed_callback", None)
        if callable(cb):
            try: cb(self)
            except: pass
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
        for p in self.points:
            try:
                p.canvas_id = None
                p.text_id = None
                p.number_id = None
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
        try:
            self.draw_rulers()
        except Exception:
            pass

        for idx, p in enumerate(self.points):
            sx, sy = self.model_to_screen(p.mx, p.my)
            color = "red" if p.kind == "inlet" else "blue"
            try:
                base_r = 5.0
                r_px = max(2, int(round(base_r * (self.scale_factor / INITIAL_SCALE))))
            except Exception:
                r_px = 5
            
            p.canvas_id = self.canvas.create_oval(
                sx - r_px, sy - r_px, sx + r_px, sy + r_px, 
                fill=color, outline=""
            )
            
            point_number = idx + 1
            
            p.number_id = self.canvas.create_text(
                sx - r_px - 3, sy - r_px - 3,
                text=f"[{point_number}]",
                fill="darkgreen" if p.kind == "inlet" else "darkblue",
                font=("Arial", 9, "bold"),
                anchor="se"
            )
            
            label = f"{p.flow:.1f}"
            p.text_id = self.canvas.create_text(
                sx + r_px + 5, sy - 5,
                text=label,
                fill="black",
                font=("Arial", 8),
                anchor="w"
            )

        for seg in self.segments:
            seg.line_ids.clear()
            seg.text_id = None
            seg.leader_id = None

            mx1, my1, mx2, my2 = seg.mx1, seg.my1, seg.mx2, seg.my2
            sx1, sy1 = self.model_to_screen(mx1, my1)
            sx2, sy2 = self.model_to_screen(mx2, my2)

            line_width = 3 if seg.is_hovered else 1
            line_color = "gray50"
            
            seg.line_ids.append(self.canvas.create_line(
                sx1, sy1, sx2, sy2, 
                fill=line_color, width=line_width, tags=("duct_line",)
            ))
            
            is_vert_draw = abs(sx1 - sx2) < abs(sy1 - sy2)
            leader_length_px = 15
            text_offset_px = 5
            text_tags = ("duct_text",)

            if is_vert_draw:
                mid_sx = sx1
                mid_sy = (sy1 + sy2) / 2.0
                vx, vy = mid_sx, mid_sy
                seg.leader_id = self.canvas.create_line(vx, vy, vx + leader_length_px, vy, fill="blue")
                tx, ty = vx + leader_length_px + text_offset_px, vy
                try:
                    vel_text = ""
                    vvar = getattr(self, 'velocity_check_var', None)
                    if vvar is None:
                        vvar = globals().get('velocity_check_var', None)
                    if vvar is not None and vvar.get():
                        q_m3h = getattr(seg, 'flow', 0.0)
                        w_mm = getattr(seg, 'duct_w_mm', 0)
                        h_mm = getattr(seg, 'duct_h_mm', 0)
                        area_m2 = 0.0
                        if w_mm > 0 and h_mm > 0:
                            area_m2 = (w_mm/1000.0) * (h_mm/1000.0)
                        elif w_mm > 0:
                            r = (w_mm/1000.0) / 2.0
                            area_m2 = math.pi * r * r
                        if area_m2 > 0 and q_m3h > 0:
                            vel = (q_m3h/3600.0) / area_m2
                            vel_text = f" - {vel:.2f} m/s"
                except Exception:
                    vel_text = ""
                seg.text_id = self.canvas.create_text(
                    tx, ty, text=seg.label_text + vel_text, 
                    fill="blue", font=("Arial", 8), anchor="w", tags=text_tags
                )
            else:
                mid_sx = (sx1 + sx2) / 2.0
                mid_sy = sy1
                hx, hy = mid_sx, mid_sy
                seg.leader_id = self.canvas.create_line(hx, hy, hx, hy - leader_length_px, fill="blue")
                tx, ty = hx, hy - leader_length_px - text_offset_px
                try:
                    vel_text = ""
                    vvar = getattr(self, 'velocity_check_var', None)
                    if vvar is None:
                        vvar = globals().get('velocity_check_var', None)
                    if vvar is not None and vvar.get():
                        q_m3h = getattr(seg, 'flow', 0.0)
                        w_mm = getattr(seg, 'duct_w_mm', 0)
                        h_mm = getattr(seg, 'duct_h_mm', 0)
                        area_m2 = 0.0
                        if w_mm > 0 and h_mm > 0:
                            area_m2 = (w_mm/1000.0) * (h_mm/1000.0)
                        elif w_mm > 0:
                            r = (w_mm/1000.0) / 2.0
                            area_m2 = math.pi * r * r
                        if area_m2 > 0 and q_m3h > 0:
                            vel = (q_m3h/3600.0) / area_m2
                            vel_text = f" - {vel:.2f} m/s"
                except Exception:
                    vel_text = ""
                seg.text_id = self.canvas.create_text(
                    tx, ty, text=seg.label_text + vel_text, 
                    fill="blue", font=("Arial", 8), anchor="s", tags=text_tags
                )

        try:
            if self.current_mouse_model is not None and self.points:
                mx_mouse, my_mouse = self.current_mouse_model
                sx_mouse, sy_mouse = self.model_to_screen(mx_mouse, my_mouse)
                ALIGN_TOL_PX = 6

                for pt in self.points:
                    if getattr(pt, 'kind', None) not in ('inlet', 'outlet'):
                        continue
                    sx_p, sy_p = self.model_to_screen(pt.mx, pt.my)

                    if abs(sy_mouse - sy_p) <= ALIGN_TOL_PX:
                        snapped_x, _ = self.snap_model(mx_mouse, my_mouse)
                        sx_grid, _ = self.model_to_screen(snapped_x, pt.my)
                        self.canvas.create_line(
                            sx_p, sy_p, sx_grid, sy_p, 
                            fill="darkorange", dash=(4, 3), width=1, tags=("align_guide",)
                        )

                    if abs(sx_mouse - sx_p) <= ALIGN_TOL_PX:
                        _, snapped_y = self.snap_model(mx_mouse, my_mouse)
                        _, sy_grid = self.model_to_screen(pt.mx, snapped_y)
                        self.canvas.create_line(
                            sx_p, sy_p, sx_p, sy_grid, 
                            fill="darkorange", dash=(4, 3), width=1, tags=("align_guide",)
                        )
        except Exception:
            pass

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
        candidates = []
        for seg in self.segments:
            if seg.vertical_only:
                if min(seg.my1, seg.my2) - tol <= my <= max(seg.my1, seg.my2) + tol:
                    if abs(mx - seg.mx1) <= tol:
                        candidates.append(seg)
            else:
                if min(seg.mx1, seg.mx2) - tol <= mx <= max(seg.mx1, seg.mx2) + tol:
                    if abs(my - seg.my1) <= tol:
                        candidates.append(seg)
        if not candidates:
            return None
        try:
            shortest = min(candidates, key=lambda s: s.length_m())
            return shortest
        except Exception:
            return candidates[0]

    def set_inlet_flow(self, flow):
        self.push_undo()
        self.inlet_flow = float(flow)
        if self.points:
            p0 = self.points[0]
            if p0.kind == "inlet":
                p0.flow = self.inlet_flow
        self.redraw_all()
        self._notify_points_changed()

    def on_left_click(self, event):
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

        mx, my = self.screen_to_model(event.x, event.y)
        seg = self._hit_test_segment(mx, my, tol=0.15)
        if seg is not None: return

        sx, sy = event.x, event.y
        mx, my = self.screen_to_model(sx, sy)
        mx, my = self.snap_model(mx, my)

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
        try:
            dp_current = float(resistance_entry.get())
        except:
            dp_current = 0.1
        if dp_current <= 0:
            messagebox.showerror("오류", "정압값이 유효하지 않습니다.")
            return
        
        msg = (
            f"현재 구간 풍량: {seg.flow} m³/h\n"
            f"현재 사이즈: {seg.duct_w_mm} x {seg.duct_h_mm}\n\n"
            f"고정할 덕트 한 변의 길이(mm)를 입력하세요.\n"
        )
        val = simpledialog.askinteger("덕트 사이즈 변경", msg, minvalue=50, maxvalue=5000)
        if val:
            self.push_undo()
            w, h, label = perform_sizing(seg.flow, dp_current, True, float(val), 1.0)
            seg.duct_w_mm = w
            seg.duct_h_mm = h
            seg.label_text = label
            self.redraw_all()
            try:
                compute_and_display_thickness_breakdown()
            except Exception:
                pass

    def on_right_click(self, event):
        sx, sy = event.x, event.y
        mx, my = self.screen_to_model(sx, sy)

        if self.mode == "pencil":
            seg_hit = self._hit_test_segment(mx, my, tol=0.15)
            if seg_hit is not None:
                self.push_undo()
                try:
                    self.segments.remove(seg_hit)
                except ValueError:
                    pass
                self.redraw_all()
                self._notify_points_changed()
                return

        target = self._find_point_near_model(mx, my, tol_model=0.3)
        if target is None: return

        if target.kind == "inlet":
            messagebox.showinfo("정보", "Air inlet의 풍량은 좌측 텍스트 값을 사용합니다.")
            return

        point_number = self.points.index(target) + 1 if target in self.points else "?"

        remaining = self._calc_remaining_flow(exclude=target)
        current = target.flow
        msg = (
            f"포인트 번호: [{point_number}]\n"
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
        self.redraw_all()

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
            if dx >= dy:
                sy2 = sy1
            else:
                sx2 = sx1
            if self._preview_line_id is not None:
                self.canvas.coords(self._preview_line_id, sx1, sy1, sx2, sy2)
            return

        mx, my = self.screen_to_model(event.x, event.y)
        self.current_mouse_model = (mx, my)

        if self.dragging_segment is None:
            seg = self._hit_test_segment(mx, my, tol=0.15)
            if seg is None: return
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
                y2 = y1
                vertical = False
            else:
                x2 = x1
                vertical = True
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
        self.push_undo()
        if not self.segments: return
        eps = 1e-6
        
        def key_of(x, y):
            return (round(x/eps)*eps, round(y/eps)*eps)
        
        reps = {}
        for seg in self.segments:
            for (x, y) in ((seg.mx1, seg.my1), (seg.mx2, seg.my2)):
                k = key_of(x, y)
                if k not in reps:
                    reps[k] = (x, y)
        for seg in self.segments:
            k1 = key_of(seg.mx1, seg.my1)
            k2 = key_of(seg.mx2, seg.my2)
            seg.mx1, seg.my1 = reps.get(k1, (seg.mx1, seg.my1))
            seg.mx2, seg.my2 = reps.get(k2, (seg.mx2, seg.my2))

        self._orthogonalize_segments()
        self._split_intersections()

        outlet_points = [p for p in self.points if getattr(p, 'kind', None) == 'outlet']
        duct_endpoints = []
        for seg in self.segments:
            duct_endpoints.append((seg.mx1, seg.my1))
            duct_endpoints.append((seg.mx2, seg.my2))

        for outlet in outlet_points:
            ox, oy = outlet.mx, outlet.my
            already_connected = False
            for seg in self.segments:
                if (abs(seg.mx1 - ox) < eps and abs(seg.my1 - oy) < eps) or \
                   (abs(seg.mx2 - ox) < eps and abs(seg.my2 - oy) < eps):
                    already_connected = True
                    break
            if already_connected: continue
            
            min_dist = float('inf')
            nearest = None
            for dx, dy in duct_endpoints:
                dist = (dx - ox)**2 + (dy - oy)**2
                if dist < min_dist:
                    min_dist = dist
                    nearest = (dx, dy)
            if nearest is None: continue
            
            if abs(nearest[0] - ox) >= abs(nearest[1] - oy):
                mid_x, mid_y = ox, nearest[1]
            else:
                mid_x, mid_y = nearest[0], oy
            
            q = getattr(outlet, 'flow', 0.0)
            w, h, label = perform_sizing(q, dp, use_fixed, fixed_val, aspect_r)
            self.segments.append(DuctSegment(nearest[0], nearest[1], mid_x, mid_y, label, w, h, q, vertical_only=(nearest[0]==mid_x)))
            self.segments.append(DuctSegment(mid_x, mid_y, ox, oy, label, w, h, q, vertical_only=(mid_x==ox)))

        self._ensure_inlet_connected()
        self._recalculate_segment_flows(dp, use_fixed, fixed_val, aspect_r)

        # ---- 추가: 자동완성 후 최적화 3회 반복 ----
        total_opt = 0
        max_iter = 3
        for _ in range(max_iter):
            c = self._optimize_outlet_connections_safe(dp, use_fixed, fixed_val, aspect_r)
            total_opt += c
            if c == 0:
                break
        # -------------------------------------------

        self.redraw_all()
        self._notify_points_changed()

    def _split_intersections(self):
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
            if a.vertical_only:
                pts = sorted(set(split_points), key=lambda p: p[1])
            else:
                pts = sorted(set(split_points), key=lambda p: p[0])
            if len(pts) <= 1:
                continue
            x1, y1 = pts[0]
            x2, y2 = pts[1]
            a.mx1, a.my1, a.mx2, a.my2 = x1, y1, x2, y2
            result.append(a)
            for k in range(1, len(pts)-1):
                xx1, yy1 = pts[k]
                xx2, yy2 = pts[k+1]
                if abs(xx1-xx2) < eps and abs(yy1-yy2) < eps: continue
                new_seg = DuctSegment(xx1, yy1, xx2, yy2, a.label_text, a.duct_w_mm, a.duct_h_mm, a.flow, a.vertical_only)
                result.append(new_seg)
        if result:
            self.segments = result

    def _move_connected_segments(self, base_seg, dx, dy):
        def seg_endpoints(seg):
            return [(seg.mx1, seg.my1), (seg.mx2, seg.my2)]
        connected = set([base_seg])
        queue = [base_seg]
        while queue:
            cur = queue.pop(0)
            cur_ends = seg_endpoints(cur)
            for other in self.segments:
                if other in connected: continue
                other_ends = seg_endpoints(other)
                if any((abs(ex1 - ex2) < 1e-9 and abs(ey1 - ey2) < 1e-9) for (ex1, ey1) in cur_ends for (ex2, ey2) in other_ends):
                    connected.add(other)
                    queue.append(other)
        point_positions = [(p.mx, p.my) for p in self.points]
        def is_attached_to_point(x, y):
            for px, py in point_positions:
                if abs(px - x) < 1e-9 and abs(py - y) < 1e-9:
                    return True
            return False
        for seg in connected:
            if not is_attached_to_point(seg.mx1, seg.my1):
                seg.mx1 += dx
                seg.my1 += dy
            if not is_attached_to_point(seg.mx2, seg.my2):
                seg.mx2 += dx
                seg.my2 += dy

    def _orthogonalize_segments(self):
        if not self.segments: return
        for seg in self.segments:
            if seg.vertical_only:
                x_avg = (seg.mx1 + seg.mx2) / 2.0
                seg.mx1 = x_avg
                seg.mx2 = x_avg
            else:
                y_avg = (seg.my1 + seg.my2) / 2.0
                seg.my1 = y_avg
                seg.my2 = y_avg
        def key_of(x, y, eps=1e-6):
            return (round(x/eps)*eps, round(y/eps)*eps)
        rep = {}
        for seg in self.segments:
            for (x, y) in ((seg.mx1, seg.my1), (seg.mx2, seg.my2)):
                k = key_of(x, y)
                if k not in rep:
                    rep[k] = (x, y)
        for seg in self.segments:
            k1 = key_of(seg.mx1, seg.my1)
            k2 = key_of(seg.mx2, seg.my2)
            if k1 in rep:
                seg.mx1, seg.my1 = rep[k1]
            if k2 in rep:
                seg.mx2, seg.my2 = rep[k2]

    def _ensure_inlet_connected(self):
        if not self.points: return
        inlet = self.points[0]
        if inlet.kind != "inlet": return
        ix, iy = inlet.mx, inlet.my
        for seg in self.segments:
            if (abs(seg.mx1 - ix) < 1e-9 and abs(seg.my1 - iy) < 1e-9) or \
               (abs(seg.mx2 - ix) < 1e-9 and abs(seg.my2 - iy) < 1e-9):
                return
        nearest_seg = None
        nearest_dist = float("inf")
        proj_point = None
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
        if nearest_seg is None or proj_point is None: return
        px, py = proj_point
        if nearest_dist < 1e-9: return
        cur_x, cur_y = ix, iy
        if abs(px - cur_x) > 1e-9:
            self.segments.append(DuctSegment(cur_x, cur_y, px, cur_y, "", 0, 0, 0.0, False))
            cur_x = px
        if abs(py - cur_y) > 1e-9:
            self.segments.append(DuctSegment(cur_x, cur_y, cur_x, py, "", 0, 0, 0.0, True))

    def _verify_all_outlets_connected(self):
        if not self.segments or not self.points:
            return False
        
        eps = 1e-6
        def node_key(x, y):
            return (round(x / eps) * eps, round(y / eps) * eps)
        
        adj = defaultdict(set)
        for seg in self.segments:
            n1 = node_key(seg.mx1, seg.my1)
            n2 = node_key(seg.mx2, seg.my2)
            adj[n1].add(n2)
            adj[n2].add(n1)
        
        inlet = self.points[0]
        if inlet.kind != "inlet":
            return False
        inlet_node = node_key(inlet.mx, inlet.my)
        
        visited = set()
        queue = deque([inlet_node])
        visited.add(inlet_node)
        
        while queue:
            cur = queue.popleft()
            for nbr in adj[cur]:
                if nbr not in visited:
                    visited.add(nbr)
                    queue.append(nbr)
        
        for p in self.points:
            if p.kind == 'outlet' and p.flow > 0:
                outlet_node = node_key(p.mx, p.my)
                if outlet_node not in visited:
                    return False
        
        return True

    def _optimize_outlet_connections_safe(self, dp: float, use_fixed: bool, fixed_val: float, aspect_r: float):
        """안전한 최적화: 조건 완화 (0.1m 이상 단축 시 적용) 및 좌표 스냅 강화"""
        if not self.segments or not self.points:
            return 0
        
        eps = 1e-6
        
        # 좌표 스냅 함수 강화
        def snap(val):
            return round(val / GRID_STEP_MODEL) * GRID_STEP_MODEL

        def node_key(x, y):
            return (round(x / eps) * eps, round(y / eps) * eps)
        
        backup_segments = copy.deepcopy(self.segments)
        outlets = [p for p in self.points if getattr(p, 'kind', None) == 'outlet' and p.flow > 0]
        optimization_count = 0
        
        for outlet in outlets:
            ox, oy = outlet.mx, outlet.my
            outlet_node = node_key(ox, oy)
            outlet_flow = outlet.flow
            
            connected_segs = []
            for seg in self.segments:
                seg_n1 = node_key(seg.mx1, seg.my1)
                seg_n2 = node_key(seg.mx2, seg.my2)
                if seg_n1 == outlet_node or seg_n2 == outlet_node:
                    connected_segs.append(seg)
            
            if not connected_segs:
                continue
            
            current_length = sum(seg.length_m() for seg in connected_segs)
            
            # 최소 1.5m 이상이어야 최적화 대상 (너무 짧으면 무시)
            if current_length < 1.5:
                continue
            
            best_point = None
            best_distance = float('inf')
            best_seg_to_split = None
            
            for seg in self.segments:
                if seg in connected_segs:
                    continue
                
                # 끝점 확인
                for px, py in [(seg.mx1, seg.my1), (seg.mx2, seg.my2)]:
                    if abs(px - ox) < eps:
                        d = abs(py - oy)
                        if d < best_distance and d > eps:
                            best_distance = d
                            best_point = (px, py)
                            best_seg_to_split = None
                    elif abs(py - oy) < eps:
                        d = abs(px - ox)
                        if d < best_distance and d > eps:
                            best_distance = d
                            best_point = (px, py)
                            best_seg_to_split = None
                
                # 세그먼트 중간 확인
                if seg.vertical_only:
                    seg_x = seg.mx1
                    seg_y_min = min(seg.my1, seg.my2)
                    seg_y_max = max(seg.my1, seg.my2)
                    if seg_y_min - eps <= oy <= seg_y_max + eps:
                        proj_x, proj_y = seg_x, oy
                        d = abs(ox - proj_x)
                        is_mid = abs(proj_y - seg.my1) > eps and abs(proj_y - seg.my2) > eps
                        if d < best_distance and d > eps:
                            best_distance = d
                            best_point = (proj_x, proj_y)
                            best_seg_to_split = seg if is_mid else None
                else:
                    seg_y = seg.my1
                    seg_x_min = min(seg.mx1, seg.mx2)
                    seg_x_max = max(seg.mx1, seg.mx2)
                    if seg_x_min - eps <= ox <= seg_x_max + eps:
                        proj_x, proj_y = ox, seg_y
                        d = abs(oy - proj_y)
                        is_mid = abs(proj_x - seg.mx1) > eps and abs(proj_x - seg.mx2) > eps
                        if d < best_distance and d > eps:
                            best_distance = d
                            best_point = (proj_x, proj_y)
                            best_seg_to_split = seg if is_mid else None
            
            # ★★★ 조건 완화: 1.0m만 절약되어도 적용 ★★★
            savings = current_length - best_distance
            if best_point is None or savings < 1.0:
                continue
            
            pre_opt_segments = copy.deepcopy(self.segments)
            
            try:
                # 좌표 스냅 적용 (중요: 부동소수점 오차 방지)
                bx = snap(best_point[0])
                by = snap(best_point[1])
                
                # 분할
                if best_seg_to_split is not None and best_seg_to_split in self.segments:
                    self.segments.remove(best_seg_to_split)
                    if best_seg_to_split.vertical_only:
                        y_min = min(best_seg_to_split.my1, best_seg_to_split.my2)
                        y_max = max(best_seg_to_split.my1, best_seg_to_split.my2)
                        # 스냅된 좌표 사용
                        self.segments.append(DuctSegment(
                            bx, y_min, bx, by,
                            best_seg_to_split.label_text, best_seg_to_split.duct_w_mm,
                            best_seg_to_split.duct_h_mm, best_seg_to_split.flow, True
                        ))
                        self.segments.append(DuctSegment(
                            bx, by, bx, y_max,
                            best_seg_to_split.label_text, best_seg_to_split.duct_w_mm,
                            best_seg_to_split.duct_h_mm, best_seg_to_split.flow, True
                        ))
                    else:
                        x_min = min(best_seg_to_split.mx1, best_seg_to_split.mx2)
                        x_max = max(best_seg_to_split.mx1, best_seg_to_split.mx2)
                        self.segments.append(DuctSegment(
                            x_min, by, bx, by,
                            best_seg_to_split.label_text, best_seg_to_split.duct_w_mm,
                            best_seg_to_split.duct_h_mm, best_seg_to_split.flow, False
                        ))
                        self.segments.append(DuctSegment(
                            bx, by, x_max, by,
                            best_seg_to_split.label_text, best_seg_to_split.duct_w_mm,
                            best_seg_to_split.duct_h_mm, best_seg_to_split.flow, False
                        ))
                
                # 새 연결
                w, h, label = perform_sizing(outlet_flow, dp, use_fixed, fixed_val, aspect_r)
                if abs(bx - ox) > eps and abs(by - oy) < eps:
                    self.segments.append(DuctSegment(bx, by, ox, oy, label, w, h, outlet_flow, False))
                elif abs(by - oy) > eps and abs(bx - ox) < eps:
                    self.segments.append(DuctSegment(bx, by, ox, oy, label, w, h, outlet_flow, True))
                
                # 기존 제거
                for seg in connected_segs:
                    if seg in self.segments:
                        self.segments.remove(seg)
                
                # 검증
                if self._verify_all_outlets_connected():
                    optimization_count += 1
                else:
                    self.segments = pre_opt_segments
                    
            except Exception:
                self.segments = pre_opt_segments
        
        if optimization_count > 0:
            self._recalculate_segment_flows(dp, use_fixed, fixed_val, aspect_r)
        
        if not self._verify_all_outlets_connected():
            self.segments = backup_segments
            return 0
        
        return optimization_count

    def _recalculate_segment_flows(self, dp: float, use_fixed: bool, fixed_val: float, aspect_r: float):
        if not self.segments or not self.points:
            return
        
        eps = 1e-6
        def node_key(x, y):
            return (round(x / eps) * eps, round(y / eps) * eps)
        
        adj = defaultdict(list)
        for seg in self.segments:
            n1 = node_key(seg.mx1, seg.my1)
            n2 = node_key(seg.mx2, seg.my2)
            adj[n1].append((n2, seg))
            adj[n2].append((n1, seg))
        
        inlet = self.points[0]
        if inlet.kind != "inlet":
            return
        inlet_node = node_key(inlet.mx, inlet.my)
        
        seg_flow_acc = defaultdict(float)
        outlets = [p for p in self.points if p.kind == 'outlet' and p.flow > 0]
        
        for out in outlets:
            start = node_key(out.mx, out.my)
            if start not in adj:
                continue
            
            queue = deque([start])
            parent = {start: None}
            parent_seg = {}
            found = False
            
            while queue:
                cur = queue.popleft()
                if cur == inlet_node:
                    found = True
                    break
                for (nbr, seg) in adj.get(cur, []):
                    if nbr in parent:
                        continue
                    parent[nbr] = cur
                    parent_seg[nbr] = seg
                    queue.append(nbr)
            
            if not found:
                continue
            
            node = inlet_node
            while True:
                prev = parent.get(node)
                if prev is None:
                    break
                seg = parent_seg.get(node)
                if seg is not None:
                    seg_flow_acc[seg] += out.flow
                node = prev
        
        for seg in self.segments:
            f = seg_flow_acc.get(seg, 0.0)
            seg.flow = f
            w, h, label = perform_sizing(f, dp, use_fixed, fixed_val, aspect_r)
            seg.duct_w_mm = w
            seg.duct_h_mm = h
            seg.label_text = label

    def draw_duct_network(self, dp_mmAq_per_m: float, use_fixed: bool, fixed_val: float, aspect_ratio: float):
        """종합 사이징 + 안전한 최적화 (여러 회 반복)"""
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
            if abs(x1 - x2) < 1e-6 and abs(y1 - y2) < 1e-6:
                return 
            w, h, label = perform_sizing(flow, dp_mmAq_per_m, use_fixed, fixed_val, aspect_ratio)
            self.segments.append(DuctSegment(x1, y1, x2, y2, label, w, h, flow, is_vert))

        sum_dx = sum(abs(p.mx - inlet.mx) for p in outlets)
        sum_dy = sum(abs(p.my - inlet.my) for p in outlets)
        is_vertical_main = sum_dy > sum_dx

        groups = {}
        tol = 0.1

        if is_vertical_main:
            for ot in outlets:
                found_key = None
                for k in groups.keys():
                    if abs(k - ot.mx) < tol:
                        found_key = k
                        break
                if found_key is None:
                    groups[ot.mx] = [ot]
                else:
                    groups[found_key].append(ot)
        else:
            for ot in outlets:
                found_key = None
                for k in groups.keys():
                    if abs(k - ot.my) < tol:
                        found_key = k
                        break
                if found_key is None:
                    groups[ot.my] = [ot]
                else:
                    groups[found_key].append(ot)

        main_nodes = []

        if is_vertical_main:
            for grp_x, grp_outlets in groups.items():
                grp_flow = sum(o.flow for o in grp_outlets)
                takeoff_y = min(grp_outlets, key=lambda o: abs(o.my - inlet.my)).my
                main_nodes.append({'pos': takeoff_y, 'flow': grp_flow, 'grp_coord': grp_x, 'outlets': grp_outlets})
            
            if not main_nodes:
                self.redraw_all()
                self._notify_points_changed()
                return
                
            x_main = inlet.mx
            
            up_nodes = [n for n in main_nodes if n['pos'] < inlet.my - 1e-6]
            down_nodes = [n for n in main_nodes if n['pos'] > inlet.my + 1e-6]
            
            if up_nodes:
                up_nodes.sort(key=lambda n: n['pos'], reverse=True)
                curr_y = inlet.my
                current_main_flow = sum(n['flow'] for n in up_nodes)
                create_seg(x_main, curr_y, x_main, up_nodes[0]['pos'], current_main_flow, True)
                for i in range(len(up_nodes) - 1):
                    current_main_flow -= up_nodes[i]['flow']
                    if current_main_flow <= 0: break
                    create_seg(x_main, up_nodes[i]['pos'], x_main, up_nodes[i+1]['pos'], current_main_flow, True)

            if down_nodes:
                down_nodes.sort(key=lambda n: n['pos'])
                curr_y = inlet.my
                current_main_flow = sum(n['flow'] for n in down_nodes)
                create_seg(x_main, curr_y, x_main, down_nodes[0]['pos'], current_main_flow, True)
                for i in range(len(down_nodes) - 1):
                    current_main_flow -= down_nodes[i]['flow']
                    if current_main_flow <= 0: break
                    create_seg(x_main, down_nodes[i]['pos'], x_main, down_nodes[i+1]['pos'], current_main_flow, True)

            for node in main_nodes:
                takeoff_y = node['pos']
                grp_x = node['grp_coord']
                grp_outlets = node['outlets']
                
                if abs(x_main - grp_x) > 1e-6:
                    create_seg(x_main, takeoff_y, grp_x, takeoff_y, node['flow'], False)
                
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
            for grp_y, grp_outlets in groups.items():
                grp_flow = sum(o.flow for o in grp_outlets)
                takeoff_x = min(grp_outlets, key=lambda o: abs(o.mx - inlet.mx)).mx
                main_nodes.append({'pos': takeoff_x, 'flow': grp_flow, 'grp_coord': grp_y, 'outlets': grp_outlets})

            if not main_nodes:
                self.redraw_all()
                self._notify_points_changed()
                return
                
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

            pos_map = {}
            for node in main_nodes:
                pos_map.setdefault(node['pos'], []).append(node)

            for takeoff_x, nodes_at_x in pos_map.items():
                ys = sorted(set([y_main] + [n['grp_coord'] for n in nodes_at_x]))
                if len(ys) <= 1:
                    continue
                for i in range(len(ys) - 1):
                    y0 = ys[i]
                    y1 = ys[i+1]
                    mid = (y0 + y1) / 2.0
                    if mid > y_main:
                        flow = sum(n['flow'] for n in nodes_at_x if n['grp_coord'] >= y1 - 1e-9)
                    else:
                        flow = sum(n['flow'] for n in nodes_at_x if n['grp_coord'] <= y0 + 1e-9)
                    create_seg(takeoff_x, y0, takeoff_x, y1, flow, True)

            for node in main_nodes:
                takeoff_x = node['pos']
                grp_y = node['grp_coord']
                grp_outlets = node['outlets']

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

        # 안전한 최적화 여러 회 반복 (최대 3회)
        total_opt_count = 0
        max_iter = 3
        for _ in range(max_iter):
            opt_count = self._optimize_outlet_connections_safe(
                dp_mmAq_per_m, use_fixed, fixed_val, aspect_ratio
            )
            total_opt_count += opt_count
            if opt_count == 0:
                break
        
        self.redraw_all()
        self._notify_points_changed()
        
        if total_opt_count > 0:
            print(f"[최적화] 총 {total_opt_count}개 outlet의 경로가 단축되었습니다.")

    def undo_last_point(self):
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
            if idx == 0:
                p.flow = Q_in
            else:
                p.flow = Q_each
        self.segments.clear()
        self.redraw_all()
        self._notify_points_changed()


# =========================
# 3. GUI 이벤트 처리
# =========================

def update_sheet_area(pal: Palette):
    global last_sheet_area_m2
    total_area_m2 = 0.0
    for seg in getattr(pal, 'segments', []):
        L = seg.length_m()
        w_m = getattr(seg, 'duct_w_mm', 0) / 1000.0
        h_m = getattr(seg, 'duct_h_mm', 0) / 1000.0
        area = (w_m + h_m) * 2 * L
        total_area_m2 += area
    try:
        last_sheet_area_m2 = total_area_m2
    except Exception:
        pass
    return total_area_m2


def get_thickness_for_longside(long_mm: float, rules_list):
    for (mn, mx, thk) in rules_list:
        if mn is None: mn = 0.0
        if mx is None: mx = float('inf')
        if mn - 1e-6 <= long_mm <= mx + 1e-6:
            return thk
    return None


def compute_and_display_thickness_breakdown(numbered=True):
    global pressure_combo, low_rule_pairs, high_rule_pairs, results_text_widget, palette
    if results_text_widget is None:
        return
    try:
        pressure = pressure_combo.get()
    except Exception:
        pressure = '저압'

    rules = []
    rule_pairs = low_rule_pairs if pressure == '저압' else high_rule_pairs
    
    for item in rule_pairs:
        try:
            if len(item) == 3:
                l_thk, l_min, e_max = item
                thk_txt = l_thk.cget('text') if hasattr(l_thk, 'cget') else str(l_thk)
                try:
                    thk = float(thk_txt.replace('mm','').strip())
                except:
                    continue
                try:
                    min_txt = l_min.cget('text') if hasattr(l_min, 'cget') else str(l_min)
                    min_val = float(min_txt.replace('~','').strip())
                except:
                    min_val = 0.0
                max_txt = e_max.get().strip() if hasattr(e_max, 'get') else str(e_max).strip()
                max_val = float(max_txt) if max_txt else float('inf')
                rules.append((min_val, max_val, thk))
        except:
            continue

    area_by_thk = defaultdict(float)
    count_by_thk = defaultdict(int)

    for seg in getattr(palette, 'segments', []):
        L = seg.length_m()
        w_m = getattr(seg, 'duct_w_mm', 0) / 1000.0
        h_m = getattr(seg, 'duct_h_mm', 0) / 1000.0
        area = (w_m + h_m) * 2 * L
        long_side_mm = max(getattr(seg, 'duct_w_mm', 0), getattr(seg, 'duct_h_mm', 0))
        thk = get_thickness_for_longside(long_side_mm, rules)
        if thk is None:
            thk = 0.8
        area_by_thk[thk] += area
        count_by_thk[thk] += 1

    base = results_text_widget.get("1.0", "end").rstrip()
    if base:
        base += "\n"
    base += "[두께별 덕트 판재 소요량]\n"
    total = 0.0
    for thk in sorted(area_by_thk.keys()):
        a = area_by_thk[thk]
        total += a
        prefix = f"{list(sorted(area_by_thk.keys())).index(thk)+1}. " if numbered else "- "
        base += f"{prefix}{thk:.2f} mm : {a:.2f} m² (구간 수: {count_by_thk[thk]})\n"
    base += f"총 합계: {total:.2f} m²"

    results_text_widget.config(state="normal")
    results_text_widget.delete("1.0", "end")
    results_text_widget.insert("end", base)
    results_text_widget.config(state="disabled")


def get_sizing_params():
    try:
        dp = float(resistance_entry.get())
    except:
        dp = 0.1
    use_fixed = fixed_side_var.get()
    fixed_val = 0.0
    if use_fixed:
        try:
            fixed_val = float(fixed_side_entry.get())
        except:
            fixed_val = 0.0
    try:
        r = float(aspect_ratio_combo.get())
    except:
        r = 2.0
    return dp, use_fixed, fixed_val, r


def format_cubic_meter_entry(event=None):
    global cubic_meter_hour_entry
    try:
        s = cubic_meter_hour_entry.get()
    except:
        return
    if s is None: return
    raw = s.replace(',', '')
    if raw in ('', '-', '.', '-.'): return
    neg = raw.startswith('-')
    raw2 = raw[1:] if neg else raw
    parts = raw2.split('.')
    int_part = parts[0] if parts[0] else '0'
    try:
        intval = int(int_part)
    except:
        return
    int_fmt = f"{intval:,}"
    if neg:
        int_fmt = '-' + int_fmt
    new = int_fmt + ('.' + parts[1] if len(parts) > 1 else '')
    if new != s:
        cubic_meter_hour_entry.delete(0, 'end')
        cubic_meter_hour_entry.insert(0, new)


def calculate():
    try:
        q = float(cubic_meter_hour_entry.get().replace(',', ''))
        dp, use_fixed, fixed_val, r = get_sizing_params()

        D1 = calc_circular_diameter(q, dp)
        D2 = round_step_up(D1, 50)

        w, h, label = perform_sizing(q, dp, use_fixed, fixed_val, r)
        try:
            _, _, _, theo_big, theo_small = size_rect_from_D1(D1, r, 50)
        except:
            theo_big, theo_small = w, h

        text = "[덕트 사이즈 결과]\n"
        text += f"- 원형덕트 (이론치) : {D1:.0f}\n"
        text += f"- 원형덕트(규격화) : {D2}\n"
        text += f"- 사각덕트 (이론치) : {theo_big:.1f} X {theo_small:.1f}\n"
        text += f"- 사각덕트(규격화) : {w} X {h}\n"
        if use_fixed:
            text += "※ 고정 변 모드 적용 중\n"
        text += f"※ 팔레트 격자 1칸 = 0.5 m"
        
        results_text_widget.config(state="normal")
        results_text_widget.delete("1.0", "end")
        results_text_widget.insert("end", text)
        results_text_widget.config(state="disabled")

        try:
            palette.set_inlet_flow(q)
        except:
            pass

    except Exception as e:
        messagebox.showerror("오류", f"입력값을 확인하세요!\n{e}")


def total_sizing():
    dp, use_fixed, fixed_val, r = get_sizing_params()
    if use_fixed and fixed_val <= 0:
        messagebox.showwarning("경고", "고정 변 길이(mm)를 올바르게 입력하세요.")
        return
    palette.draw_duct_network(dp, use_fixed, fixed_val, r)
    try:
        update_sheet_area(palette)
        compute_and_display_thickness_breakdown(numbered=False)
    except:
        pass


def auto_complete_action():
    dp, use_fixed, fixed_val, r = get_sizing_params()
    if use_fixed and fixed_val <= 0:
        messagebox.showwarning("경고", "고정 변 길이(mm)를 올바르게 입력하세요.")
        return
    palette.auto_complete(dp, use_fixed, fixed_val, r)
    try:
        update_sheet_area(palette)
        compute_and_display_thickness_breakdown(numbered=False)
    except:
        pass


def clear_palette():
    palette.clear_all()


def equal_distribution():
    try:
        q = float(cubic_meter_hour_entry.get().replace(',', ''))
        palette.set_inlet_flow(q)
    except:
        messagebox.showerror("입력 오류", "풍량 값을 확인하세요.")
        return
    palette.distribute_equal_flow()


def undo_point():
    palette.undo_last_point()


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
    root.title("스마트 덕트 사이징 프로그램 (v3.2 - 최단경로 최적화 + 포인트 번호)")

    main_frame = tk.Frame(root)
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)
    root.bind("<Control-z>", lambda event: undo_point())
    
    global cubic_meter_hour_entry, resistance_entry, aspect_ratio_combo, fixed_side_var, fixed_side_entry
    global results_text_widget, palette
    global pressure_combo, low_rule_pairs, high_rule_pairs, velocity_check_var

    # 환경 입력
    info_frame = tk.LabelFrame(main_frame, text="환경 입력", width=180)
    info_frame.pack(side="left", fill="y", padx=(0, 10))
    for lbl, dft in [("외기온도 (°C):", "-5.0"), ("실내온도 (°C):", "25.0"), 
                      ("급기온도 (°C):", "18.0"), ("일반 발열량 (W/m²):", "0.00"), 
                      ("장비 발열량 (W/m²):", "0.00")]:
        tk.Label(info_frame, text=lbl).pack(anchor="w")
        e = tk.Entry(info_frame, width=10)
        e.pack(anchor="w", pady=2)
        e.insert(0, dft)

    # 덕트 사이징
    left_frame = tk.LabelFrame(main_frame, text="덕트 사이징", width=120)
    left_frame.pack(side="left", anchor="n", fill="y")

    right_frame = tk.Frame(main_frame, bg="#f5f5f5", bd=1, relief="solid")
    right_frame.configure(width=940)
    right_frame.pack(side="right", fill="both", expand=True)

    notebook = ttk.Notebook(left_frame)
    notebook.pack(fill="both", expand=False)
    
    # 탭1: 사이징/설정
    tab1 = tk.Frame(notebook)
    notebook.add(tab1, text="사이징/설정")

    # 탭2: 장방형 덕트 두께
    tab_thickness = tk.Frame(notebook)
    notebook.add(tab_thickness, text="장방형 덕트 두께")
    tk.Label(tab_thickness, text="덕트 종류 선택:").grid(row=0, column=0, sticky="w", padx=6, pady=(6, 2))
    pressure_combo = ttk.Combobox(tab_thickness, values=["저압", "고압"], state="readonly", width=8)
    pressure_combo.current(0)
    pressure_combo.grid(row=0, column=1, sticky="w", padx=6, pady=(6, 2))

    # 탭3: 검사
    tab_check = tk.Frame(notebook)
    notebook.add(tab_check, text="검사")
    velocity_check_var = tk.BooleanVar(value=False)
    tk.Checkbutton(tab_check, text="유속체크", variable=velocity_check_var, 
                   command=lambda: palette.redraw_all()).grid(row=0, column=0, sticky="w", padx=6, pady=6)

    # 두께 규칙 UI
    tk.Label(tab_thickness, text="저압 규칙:").grid(row=1, column=0, columnspan=3, sticky="w", padx=6)
    
    low_defaults = [(0.5, 450), (0.6, 750), (0.8, 1500), (1.0, 2250), (1.2, None)]
    low_rule_pairs = []
    prev_max = -1
    for i, (thk, max_def) in enumerate(low_defaults):
        r = 2 + i
        l_thk = tk.Label(tab_thickness, text=f"{thk:.2f} mm", font=("Arial", 9))
        l_thk.grid(row=r, column=0, pady=2, sticky="w")
        min_txt = "0~" if i == 0 else f"{int(prev_max)+1}~"
        l_min = tk.Label(tab_thickness, text=min_txt, width=8, font=("Arial", 9))
        l_min.grid(row=r, column=1, padx=4, pady=2, sticky="w")
        e_max = tk.Entry(tab_thickness, width=8)
        e_max.grid(row=r, column=2, pady=2, sticky="w")
        if max_def:
            e_max.insert(0, str(max_def))
            prev_max = max_def
        low_rule_pairs.append((l_thk, l_min, e_max))

    high_row = 2 + len(low_defaults)
    tk.Label(tab_thickness, text="고압 규칙:").grid(row=high_row, column=0, columnspan=3, sticky="w", padx=6)
    
    high_defaults = [(0.8, 450), (1.0, 1200), (1.2, None)]
    high_rule_pairs = []
    prev_max_h = -1
    for i, (thk, max_def) in enumerate(high_defaults):
        r = high_row + 1 + i
        l_thk = tk.Label(tab_thickness, text=f"{thk:.2f} mm", font=("Arial", 9))
        l_thk.grid(row=r, column=0, pady=2, sticky="w")
        min_txt = "0~" if i == 0 else f"{int(prev_max_h)+1}~"
        l_min = tk.Label(tab_thickness, text=min_txt, width=8, font=("Arial", 9))
        l_min.grid(row=r, column=1, padx=4, pady=2, sticky="w")
        e_max = tk.Entry(tab_thickness, width=8)
        e_max.grid(row=r, column=2, pady=2, sticky="w")
        if max_def:
            e_max.insert(0, str(max_def))
            prev_max_h = max_def
        high_rule_pairs.append((l_thk, l_min, e_max))

    btn_row = high_row + 1 + len(high_defaults)
    tk.Button(tab_thickness, text="두께별 소요량 계산", 
              command=compute_and_display_thickness_breakdown).grid(row=btn_row, column=0, columnspan=3, pady=6, sticky="w")

    # 사이징 컨트롤
    ctrl = tab1
    row = 0
    
    tk.Label(ctrl, text="풍량 (m³/h):").grid(row=row, column=0, padx=5, pady=5, sticky="w")
    cubic_meter_hour_entry = tk.Entry(ctrl, width=10)
    cubic_meter_hour_entry.grid(row=row, column=1, padx=5, pady=5, sticky="w")
    cubic_meter_hour_entry.insert(0, "50000")
    cubic_meter_hour_entry.bind('<KeyRelease>', format_cubic_meter_entry)
    row += 1

    tk.Label(ctrl, text="정압값 (mmAq/m):").grid(row=row, column=0, padx=5, pady=5, sticky="w")
    resistance_entry = tk.Entry(ctrl, width=10)
    resistance_entry.grid(row=row, column=1, padx=5, pady=5, sticky="w")
    resistance_entry.insert(0, "0.1")
    row += 1

    tk.Label(ctrl, text="종횡비 (b/a):").grid(row=row, column=0, padx=5, pady=5, sticky="w")
    aspect_ratio_combo = ttk.Combobox(ctrl, values=["1", "2", "3", "4"], state="readonly", width=5)
    aspect_ratio_combo.current(1)
    aspect_ratio_combo.grid(row=row, column=1, padx=5, pady=5, sticky="w")
    row += 1

    fixed_side_var = tk.BooleanVar(value=False)
    tk.Checkbutton(ctrl, text="한 변 고정(mm):", variable=fixed_side_var, 
                   command=toggle_fixed_side).grid(row=row, column=0, padx=5, pady=5, sticky="w")
    fixed_side_entry = tk.Entry(ctrl, width=10, state="disabled")
    fixed_side_entry.grid(row=row, column=1, padx=5, pady=5, sticky="w")
    fixed_side_entry.insert(0, "300")
    row += 1

    tk.Button(ctrl, text="계산하기", command=calculate).grid(row=row, column=0, padx=5, pady=5, sticky="w")
    tk.Button(ctrl, text="균등 풍량 배분", command=equal_distribution).grid(row=row, column=1, padx=5, pady=5, sticky="w")
    row += 1
    
    tk.Button(ctrl, text="종합 사이징", command=total_sizing).grid(row=row, column=0, padx=5, pady=5, sticky="w")
    tk.Button(ctrl, text="전체 지우기", command=clear_palette).grid(row=row, column=1, padx=5, pady=5, sticky="w")
    row += 1

    tk.Button(ctrl, text="펜슬 모드", command=lambda: palette.set_mode_pencil()).grid(row=row, column=0, padx=5, pady=5, sticky="w")
    tk.Button(ctrl, text="자동완성", command=auto_complete_action).grid(row=row, column=1, padx=5, pady=5, sticky="w")
    row += 1

    # 결과창
    results_frame = tk.Frame(ctrl)
    results_frame.grid(row=row, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
    results_text_widget = tk.Text(results_frame, width=36, height=16, wrap="word", bg="white", relief="solid")
    results_text_widget.pack(side="left", fill="both", expand=True)
    scrollbar = tk.Scrollbar(results_frame, orient="vertical", command=results_text_widget.yview)
    scrollbar.pack(side="right", fill="y")
    results_text_widget.configure(yscrollcommand=scrollbar.set, state="disabled")

    # 팔레트
    global palette
    palette = Palette(right_frame)
    palette.velocity_check_var = velocity_check_var
    palette.sheet_changed_callback = update_sheet_area

    root.mainloop()


if __name__ == "__main__":
    create_app()
