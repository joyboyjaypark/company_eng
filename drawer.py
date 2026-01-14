import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, filedialog
from math import sqrt, ceil
import json

# Shapely 관련 import
from shapely.geometry import Polygon, LineString, Point
from shapely.ops import unary_union, polygonize
import re
import os

# HVAC type names
HVAC_NAMES = {
    1: "중앙공조",
    2: "개별공조",
    3: "비공조",
}


class RectShape:
    """하나의 직사각형 도형 + 치수 정보를 관리하는 클래스"""
    def __init__(self, shape_id, coords, rect_id, side_ids, dim_items,
                 editable=True, color="black"):
        self.shape_id = shape_id
        self.coords = coords          # (x1, y1, x2, y2)
        self.rect_id = rect_id        # canvas rectangle id
        self.side_ids = side_ids      # {"top": line_id, ...}
        self.dim_items = dim_items    # {"top": {...}, "left": {...}}
        self.editable = editable
        self.color = color
        self.snap_highlight_sides = set()  # 스냅으로 강조된 변 이름들


class Palette:
    """팔레트 하나(캔버스)와 그 안의 모든 도형/동작을 관리하는 클래스"""

    def __init__(self, parent, app, width=900, height=600):
        self.app = app
        self.canvas = tk.Canvas(parent, bg="white", width=width, height=height)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # 상태
        self.shapes = []
        self.next_shape_id = 1

        self.scale = 20.0  # 1m = 20px
        self.unit = "m"

        # 변 드래그
        self.active_shape = None
        self.active_side_name = None
        self.drag_start_mouse_pos = None
        self.drag_start_coords = None

        # 도형 이동(모서리)
        self.corner_snap_tolerance = 8
        self.corner_highlight_id = None
        self.corner_hover_shape = None
        self.corner_hover_index = None
        self.moving_shape = None
        self.move_start_mouse_pos = None
        self.move_start_shape_coords = None

        # 하이라이트 / 툴팁
        self.highlight_line_id = None
        self.tooltip_id = None

        # 패닝
        self.panning = False
        self.pan_last_pos = None

        # 스냅
        self.snap_tolerance = 8

        # Undo
        self.history = []

        # 코너 삭제용 팝업 메뉴
        self.corner_menu = tk.Menu(self.canvas, tearoff=0)
        self.corner_menu.add_command(label="삭제하기", command=self.delete_corner_shape)
        self.corner_menu_target_shape = None

        # 자동생성 공간 라벨
        # 각 원소: {
        #   "polygon": shapely Polygon,
        #   "name_id":..., "heat_norm_id":..., "heat_equip_id":..., "area_id":...,
        #   "diffuser_ids": [id1, id2, ...] (캔버스 아이템 ID 리스트)
        # }
        self.generated_space_labels = []

        # 격자 관련 상태: 기본으로 보이게 설정
        self.grid_ids = []
        self.show_grid = True
        # debounce handle for panning redraws
        self._grid_redraw_after_id = None

        # 이벤트 바인딩
        self.canvas.bind("<Motion>", self.on_mouse_move)
        self.canvas.bind("<ButtonPress-1>", self.on_left_down)
        self.canvas.bind("<B1-Motion>", self.on_left_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_up)

        self.canvas.bind("<ButtonPress-3>", self.on_right_click)

        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<Button-4>", self.on_mouse_wheel_linux)
        self.canvas.bind("<Button-5>", self.on_mouse_wheel_linux)

        self.canvas.bind("<ButtonPress-2>", self.on_middle_button_down)
        self.canvas.bind("<B2-Motion>", self.on_middle_button_drag)
        self.canvas.bind("<ButtonRelease-2>", self.on_middle_button_up)

        self.canvas.tag_bind("dim_width", "<Button-1>", self.on_dim_width_click)
        self.canvas.tag_bind("dim_height", "<Button-1>", self.on_dim_height_click)

        # 자동생성 텍스트 클릭(편집)
        self.canvas.tag_bind("space_name", "<Button-1>", self.on_space_name_click)
        self.canvas.tag_bind("space_heat_norm", "<Button-1>", self.on_space_heat_norm_click)
        self.canvas.tag_bind("space_heat_equip", "<Button-1>", self.on_space_heat_equip_click)

        # 캔버스 크기 변경 시 그리드 갱신 바인딩
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        try:
            # 초기 레이아웃 후 그리드를 한 번 그립니다
            self.canvas.after(50, self.draw_grid)
        except Exception:
            pass

    # -------- 기본 유틸 --------

    def pixel_to_meter(self, length_px: float) -> float:
        return length_px / self.scale

    def meter_to_pixel(self, length_m: float) -> float:
        return length_m * self.scale

    def _on_canvas_configure(self, event=None):
        if getattr(self, 'show_grid', False):
            try:
                self.draw_grid()
            except Exception:
                pass

    def clear_grid(self):
        for gid in list(getattr(self, 'grid_ids', [])):
            try:
                self.canvas.delete(gid)
            except Exception:
                pass
        self.grid_ids = []

    def toggle_grid(self, show: bool):
        """Show or hide the grid. If showing, draw it for current viewport."""
        self.show_grid = bool(show)
        if self.show_grid:
            try:
                self.draw_grid()
            except Exception:
                pass
        else:
            try:
                self.clear_grid()
            except Exception:
                pass

    def draw_grid(self):
        """Viewport-limited 0.5m grid. Coarsen spacing if too many lines to avoid UI freeze."""
        # remove old grid
        try:
            self.clear_grid()
        except Exception:
            self.grid_ids = []

        # get widget size
        try:
            w = int(self.canvas.winfo_width())
            h = int(self.canvas.winfo_height())
        except Exception:
            w = int(self.canvas['width']) if 'width' in self.canvas.keys() else 800
            h = int(self.canvas['height']) if 'height' in self.canvas.keys() else 600

        # spacing in pixels for 0.5m
        spacing = max(1.0, self.meter_to_pixel(0.5))

        # visible region
        try:
            view_left = self.canvas.canvasx(0)
            view_top = self.canvas.canvasy(0)
            view_right = self.canvas.canvasx(w)
            view_bottom = self.canvas.canvasy(h)
        except Exception:
            view_left, view_top, view_right, view_bottom = 0.0, 0.0, float(w), float(h)

        # small margin to avoid tiny gaps (use larger margin to tolerate pan/rounding)
        MARGIN_MULT = 2
        view_left -= spacing * MARGIN_MULT
        view_top -= spacing * MARGIN_MULT
        view_right += spacing * MARGIN_MULT
        view_bottom += spacing * MARGIN_MULT

        import math
        # Anchor grid to a stable reference so it stays aligned with shapes after zoom/pan.
        # Use first shape's top-left as anchor if available, otherwise use world origin (0,0).
        try:
            if self.shapes:
                anchor_x = float(self.shapes[0].coords[0])
                anchor_y = float(self.shapes[0].coords[1])
            else:
                anchor_x = 0.0
                anchor_y = 0.0
        except Exception:
            anchor_x = 0.0
            anchor_y = 0.0

        # compute remainder offset so anchor_x == (k*spacing + rem_x)
        rem_x = anchor_x - math.floor(anchor_x / spacing) * spacing
        rem_y = anchor_y - math.floor(anchor_y / spacing) * spacing

        kmin = math.floor((view_left - rem_x) / spacing)
        kmax = math.ceil((view_right - rem_x) / spacing)
        hmin = math.floor((view_top - rem_y) / spacing)
        hmax = math.ceil((view_bottom - rem_y) / spacing)

        v_count = max(0, int(kmax - kmin + 1))
        h_count = max(0, int(hmax - hmin + 1))

        # cap total lines to avoid freezing
        MAX_LINES = 1200
        total = v_count + h_count
        while total > MAX_LINES and spacing < max(w, h):
            spacing *= 2
            kmin = math.floor(view_left / spacing)
            kmax = math.ceil(view_right / spacing)
            hmin = math.floor(view_top / spacing)
            hmax = math.ceil(view_bottom / spacing)
            v_count = max(0, int(kmax - kmin + 1))
            h_count = max(0, int(hmax - hmin + 1))
            total = v_count + h_count

        color = "#e6e6e6"
        # draw vertical lines
        for k in range(kmin, kmax + 1):
            x = k * spacing + rem_x
            try:
                lid = self.canvas.create_line(x, view_top, x, view_bottom, fill=color, width=1, tags=("grid",))
                self.grid_ids.append(lid)
            except Exception:
                continue

        # draw horizontal lines
        for k in range(hmin, hmax + 1):
            y = k * spacing + rem_y
            try:
                lid = self.canvas.create_line(view_left, y, view_right, y, fill=color, width=1, tags=("grid",))
                self.grid_ids.append(lid)
            except Exception:
                continue

        try:
            self.canvas.tag_lower("grid")
        except Exception:
            pass

    def push_history(self):
        snapshot = {
            "scale": self.scale,
            "next_shape_id": self.next_shape_id,
            "shapes": [],
            "generated_space_labels": []
        }
        for s in self.shapes:
            snapshot["shapes"].append({
                "shape_id": s.shape_id,
                "coords": tuple(s.coords),
                "editable": s.editable,
                "color": s.color
            })

        # 자동생성 라벨 저장 (디퓨저 위치 포함)
        for lab in self.generated_space_labels:
            name_bbox = self.canvas.bbox(lab["name_id"])
            heat_norm_bbox = self.canvas.bbox(lab["heat_norm_id"])
            heat_equip_bbox = self.canvas.bbox(lab["heat_equip_id"])
            area_bbox = self.canvas.bbox(lab["area_id"])
            
            # 디퓨저 좌표 저장
            diffuser_coords = []
            if "diffuser_ids" in lab:
                for did in lab["diffuser_ids"]:
                    coords = self.canvas.coords(did)
                    if coords:
                        # oval coords (x1, y1, x2, y2) -> center (cx, cy)
                        cx = (coords[0] + coords[2]) / 2
                        cy = (coords[1] + coords[3]) / 2
                        diffuser_coords.append((cx, cy))

            snapshot["generated_space_labels"].append({
                "polygon_coords": list(lab["polygon"].exterior.coords),
                "name_text": self.canvas.itemcget(lab["name_id"], "text"),
                "heat_norm_text": self.canvas.itemcget(lab["heat_norm_id"], "text"),
                "heat_equip_text": self.canvas.itemcget(lab["heat_equip_id"], "text"),
                "area_text": self.canvas.itemcget(lab["area_id"], "text"),
                "name_pos": name_bbox,
                "heat_norm_pos": heat_norm_bbox,
                "heat_equip_pos": heat_equip_bbox,
                "area_pos": area_bbox,
                "diffuser_coords": diffuser_coords
            })

        self.history.append(snapshot)

    def undo(self):
        if not self.history:
            return

        snapshot = self.history.pop()
        self.canvas.delete("all")
        self.shapes.clear()
        self.generated_space_labels.clear()
        self.highlight_line_id = None
        self.tooltip_id = None
        self.corner_highlight_id = None

        self.scale = snapshot["scale"]
        self.next_shape_id = snapshot["next_shape_id"]

        # 도형 복원
        for info in snapshot["shapes"]:
            s = self.create_rect_shape(
                info["coords"][0], info["coords"][1],
                info["coords"][2], info["coords"][3],
                editable=info["editable"],
                color=info["color"],
                push_to_history=False
            )
            s.shape_id = info["shape_id"]

        # 공간 라벨 복원
        for lab in snapshot["generated_space_labels"]:
            poly = Polygon(lab["polygon_coords"])
            if not lab["name_pos"]:
                continue

            x1, y1, x2, y2 = lab["name_pos"]
            name_id = self.canvas.create_text(
                (x1 + x2) / 2, (y1 + y2) / 2,
                text=lab["name_text"], fill="blue", font=("Arial", 11, "bold"),
                tags=("space_name",)
            )

            hx1, hy1, hx2, hy2 = lab["heat_norm_pos"]
            heat_norm_id = self.canvas.create_text(
                (hx1 + hx2) / 2, (hy1 + hy2) / 2,
                text=lab["heat_norm_text"], fill="darkred", font=("Arial", 10),
                tags=("space_heat_norm",)
            )

            ex1, ey1, ex2, ey2 = lab["heat_equip_pos"]
            heat_equip_id = self.canvas.create_text(
                (ex1 + ex2) / 2, (ey1 + ey2) / 2,
                text=lab["heat_equip_text"], fill="darkred", font=("Arial", 10),
                tags=("space_heat_equip",)
            )

            ax1, ay1, ax2, ay2 = lab["area_pos"]
            area_id = self.canvas.create_text(
                (ax1 + ax2) / 2, (ay1 + ay2) / 2,
                text=lab["area_text"], fill="green", font=("Arial", 10)
            )
            
            # 디퓨저 복원
            diffuser_ids = []
            r = 3
            if "diffuser_coords" in lab:
                for (cx, cy) in lab["diffuser_coords"]:
                    did = self.canvas.create_oval(
                        cx - r, cy - r, cx + r, cy + r,
                        fill="green", outline=""
                    )
                    diffuser_ids.append(did)

            self.generated_space_labels.append({
                "polygon": poly,
                "name_id": name_id,
                "heat_norm_id": heat_norm_id,
                "heat_equip_id": heat_equip_id,
                "area_id": area_id,
                "diffuser_ids": diffuser_ids
            })

        # 태그 바인딩 복원
        self.canvas.tag_bind("dim_width", "<Button-1>", self.on_dim_width_click)
        self.canvas.tag_bind("dim_height", "<Button-1>", self.on_dim_height_click)
        self.canvas.tag_bind("space_name", "<Button-1>", self.on_space_name_click)
        self.canvas.tag_bind("space_heat_norm", "<Button-1>", self.on_space_heat_norm_click)
        self.canvas.tag_bind("space_heat_equip", "<Button-1>", self.on_space_heat_equip_click)

        self.active_shape = None
        self.active_side_name = None
        self.app.update_selected_area_label(self)

    # -------- 도형 생성/그리기 --------

    def draw_square_from_area(self, area: float):
        if area <= 0:
            return

        self.push_history()

        side_m = sqrt(area)
        side_px = self.meter_to_pixel(side_m)

        cw = int(self.canvas["width"])
        ch = int(self.canvas["height"])

        x1 = cw / 2 - side_px / 2
        y1 = ch / 2 - side_px / 2
        x2 = cw / 2 + side_px / 2
        y2 = ch / 2 + side_px / 2

        shape = self.create_rect_shape(x1, y1, x2, y2, editable=True, color="black",
                                       push_to_history=False)
        self.set_active_shape(shape)

    def create_rect_shape(self, x1, y1, x2, y2,
                          editable=True, color="black",
                          push_to_history=True):
        if x2 < x1:
            x1, x2 = x2, x1
        if y2 < y1:
            y1, y2 = y2, y1

        shape_id = self.next_shape_id
        self.next_shape_id += 1

        rect_id = self.canvas.create_rectangle(
            x1, y1, x2, y2,
            outline=color,
            width=2,
            dash=() if color == "black" else (3, 2)
        )

        top_id = self.canvas.create_line(x1, y1, x2, y1, fill=color, width=3)
        bottom_id = self.canvas.create_line(x1, y2, x2, y2, fill=color, width=3)
        left_id = self.canvas.create_line(x1, y1, x1, y2, fill=color, width=3)
        right_id = self.canvas.create_line(x2, y1, x2, y2, fill=color, width=3)

        side_ids = {"top": top_id, "bottom": bottom_id,
                    "left": left_id, "right": right_id}
        for side_name, lid in side_ids.items():
            self.canvas.addtag_withtag(f"shape_{shape_id}", lid)
            self.canvas.addtag_withtag(f"side_{shape_id}_{side_name}", lid)

        dim_items = self.draw_dimensions_for_shape(shape_id, x1, y1, x2, y2, color=color)

        shape = RectShape(shape_id, (x1, y1, x2, y2),
                          rect_id, side_ids, dim_items,
                          editable=editable, color=color)

        self.shapes.append(shape)
        self.bring_shape_to_front(shape)
        
        try:
            if self.generated_space_labels:
                pass
        except Exception:
            pass

        return shape

    def bring_shape_to_front(self, shape: RectShape):
        ids = [shape.rect_id]
        ids.extend(shape.side_ids.values())
        for part in shape.dim_items.values():
            ids.extend(part["lines"])
            ids.extend(part["ticks"])
            ids.append(part["text"])
        for item_id in ids:
            if item_id in self.canvas.find_all():
                self.canvas.tag_raise(item_id)

    def draw_dimensions_for_shape(self, shape_id, x1, y1, x2, y2, color="black"):
        dim_items = {}
        offset = 30
        tick_len = 8
        text_offset = 4

        # 가로
        width_px = x2 - x1
        width_m = self.pixel_to_meter(width_px)
        dim_y = y1 - offset

        dim_line_top = self.canvas.create_line(x1, dim_y, x2, dim_y,
                                               fill=color, width=1)
        left_tick_top = self.canvas.create_line(
            x1, dim_y - tick_len / 2, x1, dim_y + tick_len / 2,
            fill=color, width=1)
        right_tick_top = self.canvas.create_line(
            x2, dim_y - tick_len / 2, x2, dim_y + tick_len / 2,
            fill=color, width=1)
        text_x = (x1 + x2) / 2
        text_y = dim_y - text_offset

        width_text_id = self.canvas.create_text(
            text_x, text_y,
            text=f"{width_m:.2f} {self.unit}",
            fill=color,
            font=("Arial", 10),
            tags=("dim_width", f"dim_width_{shape_id}")
        )

        dim_items["top"] = {
            "lines": [dim_line_top],
            "ticks": [left_tick_top, right_tick_top],
            "text": width_text_id
        }

        # 세로
        height_px = y2 - y1
        height_m = self.pixel_to_meter(height_px)
        dim_x = x1 - offset

        dim_line_left = self.canvas.create_line(dim_x, y1, dim_x, y2,
                                                fill=color, width=1)
        top_tick_left = self.canvas.create_line(
            dim_x - tick_len / 2, y1, dim_x + tick_len / 2, y1,
            fill=color, width=1)
        bottom_tick_left = self.canvas.create_line(
            dim_x - tick_len / 2, y2, dim_x + tick_len / 2, y2,
            fill=color, width=1)
        text_x2 = dim_x - text_offset
        text_y2 = (y1 + y2) / 2

        height_text_id = self.canvas.create_text(
            text_x2, text_y2,
            text=f"{height_m:.2f} {self.unit}",
            fill=color,
            font=("Arial", 10),
            angle=90,
            tags=("dim_height", f"dim_height_{shape_id}")
        )

        dim_items["left"] = {
            "lines": [dim_line_left],
            "ticks": [top_tick_left, bottom_tick_left],
            "text": height_text_id
        }

        for item in [dim_line_top, left_tick_top, right_tick_top,
                     dim_line_left, top_tick_left, bottom_tick_left,
                     width_text_id, height_text_id]:
            self.canvas.addtag_withtag(f"shape_{shape_id}", item)

        return dim_items

    # -------- 선택/하이라이트 --------

    def get_shape_by_id(self, shape_id):
        for s in self.shapes:
            if s.shape_id == shape_id:
                return s
        return None

    def set_active_shape(self, shape):
        self.active_shape = shape
        self.app.update_selected_area_label(self)
        if shape:
            self.bring_shape_to_front(shape)

    def find_side_under_mouse(self, x, y, tol=5):
        best_shape = None
        best_side = None
        best_dist2 = None

        for shape in reversed(self.shapes):
            x1, y1, x2, y2 = shape.coords
            candidates = []
            if x1 <= x <= x2:
                candidates.append(("top", (y - y1) ** 2, abs(y - y1)))
                candidates.append(("bottom", (y - y2) ** 2, abs(y - y2)))
            if y1 <= y <= y2:
                candidates.append(("left", (x - x1) ** 2, abs(x - x1)))
                candidates.append(("right", (x - x2) ** 2, abs(x - x2)))

            for side_name, d2, absd in candidates:
                if absd <= tol:
                    if best_dist2 is None or d2 < best_dist2:
                        best_dist2 = d2
                        best_shape = shape
                        best_side = side_name
        return best_shape, best_side

    def highlight_side(self, shape, side_name):
        if self.highlight_line_id and self.highlight_line_id in self.canvas.find_all():
            self.canvas.itemconfigure(self.highlight_line_id, width=3)
        self.highlight_line_id = None

        if not shape or not side_name:
            return

        line_id = shape.side_ids.get(side_name)
        if line_id:
            self.canvas.itemconfigure(line_id, width=4)
            self.highlight_line_id = line_id

    # -------- 모서리 감지 --------

    def clear_corner_highlight(self):
        if self.corner_highlight_id and self.corner_highlight_id in self.canvas.find_all():
            self.canvas.delete(self.corner_highlight_id)
        self.corner_highlight_id = None
        self.corner_hover_shape = None
        self.corner_hover_index = None

    def detect_corner_under_mouse(self, x, y):
        tol = self.corner_snap_tolerance
        best_shape = None
        best_index = None
        best_cx = best_cy = None
        best_d2 = None

        for shape in self.shapes:
            x1, y1, x2, y2 = shape.coords
            corners = [(x1, y1), (x2, y1), (x1, y2), (x2, y2)]
            for idx, (cx, cy) in enumerate(corners):
                dx = x - cx
                dy = y - cy
                d2 = dx * dx + dy * dy
                if abs(dx) <= tol and abs(dy) <= tol:
                    if best_d2 is None or d2 < best_d2:
                        best_d2 = d2
                        best_shape = shape
                        best_index = idx
                        best_cx, best_cy = cx, cy
        return best_shape, best_index, best_cx, best_cy

    # -------- 스냅 하이라이트 --------

    def clear_edge_snap_highlight(self, shape: RectShape):
        for side in shape.snap_highlight_sides:
            lid = shape.side_ids.get(side)
            if lid in self.canvas.find_all():
                self.canvas.itemconfigure(lid, fill=shape.color)
        shape.snap_highlight_sides.clear()

    def highlight_edge_snap(self, shape: RectShape, snapped_sides):
        self.clear_edge_snap_highlight(shape)
        for side in snapped_sides:
            lid = shape.side_ids.get(side)
            if lid in self.canvas.find_all():
                self.canvas.itemconfigure(lid, fill="orange")
                shape.snap_highlight_sides.add(side)

    # -------- 툴팁 --------

    def show_length_tooltip(self, shape, side_name, mx, my):
        x1, y1, x2, y2 = shape.coords
        if side_name in ("top", "bottom"):
            length_px = x2 - x1
        else:
            length_px = y2 - y1
        length_m = self.pixel_to_meter(length_px)
        text = f"{length_m:.2f} {self.unit}"

        if self.tooltip_id and self.tooltip_id in self.canvas.find_all():
            self.canvas.delete(self.tooltip_id)
            self.tooltip_id = None

        offset = 15
        self.tooltip_id = self.canvas.create_text(
            mx + offset, my - offset,
            text=text,
            fill="darkgreen",
            font=("Arial", 10, "bold"),
            anchor="sw"
        )

    def hide_length_tooltip(self):
        if self.tooltip_id and self.tooltip_id in self.canvas.find_all():
            self.canvas.delete(self.tooltip_id)
        self.tooltip_id = None

    # -------- 공유 변 판정 --------

    def find_shared_vertical_edges(self, shape):
        x1, y1, x2, y2 = shape.coords
        shared = {"left": False, "right": False}
        for other in self.shapes:
            if other is shape:
                continue
            ox1, oy1, ox2, oy2 = other.coords
            if abs(ox1 - x1) < 1e-6 or abs(ox2 - x1) < 1e-6:
                overlap = min(y2, oy2) - max(y1, oy1)
                if overlap > 0:
                    shared["left"] = True
            if abs(ox1 - x2) < 1e-6 or abs(ox2 - x2) < 1e-6:
                overlap = min(y2, oy2) - max(y1, oy1)
                if overlap > 0:
                    shared["right"] = True
        return shared

    def find_shared_horizontal_edges(self, shape):
        x1, y1, x2, y2 = shape.coords
        shared = {"top": False, "bottom": False}
        for other in self.shapes:
            if other is shape:
                continue
            ox1, oy1, ox2, oy2 = other.coords
            if abs(oy1 - y1) < 1e-6 or abs(oy2 - y1) < 1e-6:
                overlap = min(x2, ox2) - max(x1, ox1)
                if overlap > 0:
                    shared["top"] = True
            if abs(oy1 - y2) < 1e-6 or abs(oy2 - y2) < 1e-6:
                overlap = min(x2, ox2) - max(x1, ox1)
                if overlap > 0:
                    shared["bottom"] = True
        return shared

    # -------- 코너 팝업 삭제 --------

    def delete_corner_shape(self):
        shape = self.corner_menu_target_shape
        if not shape or not shape.editable:
            return

        self.push_history()

        self.canvas.delete(shape.rect_id)
        for lid in shape.side_ids.values():
            self.canvas.delete(lid)
        for part in shape.dim_items.values():
            for lid in part["lines"] + part["ticks"] + [part["text"]]:
                self.canvas.delete(lid)

        if shape in self.shapes:
            self.shapes.remove(shape)

        if self.active_shape is shape:
            self.active_shape = None
        if self.corner_hover_shape is shape:
            self.clear_corner_highlight()

        self.app.update_selected_area_label(self)
        self.corner_menu_target_shape = None

    # -------- 마우스 이벤트 --------

    def on_mouse_move(self, event):
        if self.moving_shape:
            return

        shape, idx, cx, cy = self.detect_corner_under_mouse(event.x, event.y)
        if shape:
            if not self.corner_highlight_id or self.corner_highlight_id not in self.canvas.find_all():
                r = 4
                self.corner_highlight_id = self.canvas.create_oval(
                    cx - r, cy - r, cx + r, cy + r,
                    fill="red", outline=""
                )
            else:
                r = 4
                self.canvas.coords(self.corner_highlight_id,
                                   cx - r, cy - r, cx + r, cy + r)
            self.corner_hover_shape = shape
            self.corner_hover_index = idx
        else:
            self.clear_corner_highlight()

        if self.active_side_name is None and not self.corner_hover_shape:
            s, side = self.find_side_under_mouse(event.x, event.y, tol=5)
            self.highlight_side(s, side)

    def on_left_down(self, event):
        if self.corner_hover_shape is not None:
            self.push_history()
            self.moving_shape = self.corner_hover_shape
            self.move_start_mouse_pos = (event.x, event.y)
            self.move_start_shape_coords = self.moving_shape.coords
            self.set_active_shape(self.moving_shape)
            return

        shape, side = self.find_side_under_mouse(event.x, event.y, tol=5)
        if shape and side and shape.editable:
            self.push_history()
            self.set_active_shape(shape)
            self.active_side_name = side
            self.drag_start_mouse_pos = (event.x, event.y)
            self.drag_start_coords = shape.coords

    def on_left_drag(self, event):
        # 도형 전체 이동
        if self.moving_shape and self.move_start_mouse_pos and self.move_start_shape_coords:
            dx = event.x - self.move_start_mouse_pos[0]
            dy = event.y - self.move_start_mouse_pos[1]
            x1, y1, x2, y2 = self.move_start_shape_coords
            tentative = (x1 + dx, y1 + dy, x2 + dx, y2 + dy)

            snapped_sides = []
            tentative2, s_left = self.apply_snap_edge(self.moving_shape, "left", tentative)
            if s_left:
                snapped_sides.append("left")
                tentative = tentative2
            tentative2, s_right = self.apply_snap_edge(self.moving_shape, "right", tentative)
            if s_right:
                snapped_sides.append("right")
                tentative = tentative2
            tentative2, s_top = self.apply_snap_edge(self.moving_shape, "top", tentative)
            if s_top:
                snapped_sides.append("top")
                tentative = tentative2
            tentative2, s_bottom = self.apply_snap_edge(self.moving_shape, "bottom", tentative)
            if s_bottom:
                snapped_sides.append("bottom")
                tentative = tentative2

            self.moving_shape.coords = tentative
            self.redraw_shape(self.moving_shape)
            self.highlight_edge_snap(self.moving_shape, snapped_sides)
            self.app.update_selected_area_label(self)
            return

        # 변 드래그
        if not self.active_shape or not self.active_side_name or not self.drag_start_coords:
            return
        if not self.active_shape.editable:
            return

        x1, y1, x2, y2 = self.drag_start_coords
        dx = event.x - self.drag_start_mouse_pos[0]
        dy = event.y - self.drag_start_mouse_pos[1]
        min_size_px = 20
        side = self.active_side_name

        if side == "top":
            new_y1 = y1 + dy
            if new_y1 > y2 - min_size_px:
                new_y1 = y2 - min_size_px
            tentative = (x1, new_y1, x2, y2)
        elif side == "bottom":
            new_y2 = y2 + dy
            if new_y2 < y1 + min_size_px:
                new_y2 = y1 + min_size_px
            tentative = (x1, y1, x2, new_y2)
        elif side == "left":
            new_x1 = x1 + dx
            if new_x1 > x2 - min_size_px:
                new_x1 = x2 - min_size_px
            tentative = (new_x1, y1, x2, y2)
        elif side == "right":
            new_x2 = x2 + dx
            if new_x2 < x1 + min_size_px:
                new_x2 = x1 + min_size_px
            tentative = (x1, y1, new_x2, y2)
        else:
            return

        snapped_coords, snapped = self.apply_snap_edge(self.active_shape, side, tentative)
        self.active_shape.coords = snapped_coords

        self.redraw_shape(self.active_shape)
        self.bring_shape_to_front(self.active_shape)
        if snapped:
            self.highlight_edge_snap(self.active_shape, [side])
        else:
            self.clear_edge_snap_highlight(self.active_shape)

        self.highlight_side(self.active_shape, side)
        self.show_length_tooltip(self.active_shape, side, event.x, event.y)
        self.app.update_selected_area_label(self)

    def on_left_up(self, event):
        if self.moving_shape:
            self.clear_edge_snap_highlight(self.moving_shape)
        self.moving_shape = None
        self.move_start_mouse_pos = None
        self.move_start_shape_coords = None

        if self.active_shape:
            self.clear_edge_snap_highlight(self.active_shape)
        self.active_side_name = None
        self.drag_start_mouse_pos = None
        self.drag_start_coords = None
        self.hide_length_tooltip()

    # -------- 스냅 --------

    def apply_snap_edge(self, shape, side, coords):
        x1, y1, x2, y2 = coords
        snap = self.snap_tolerance

        candidate_positions = []
        for other in self.shapes:
            if other is shape:
                continue
            ox1, oy1, ox2, oy2 = other.coords
            if side in ("top", "bottom"):
                candidate_positions.extend([oy1, oy2])
            else:
                candidate_positions.extend([ox1, ox2])

        if not candidate_positions:
            return coords, False

        snapped = False
        if side in ("top", "bottom"):
            cur_y = y1 if side == "top" else y2
            best_y = cur_y
            best_diff = None
            for py in candidate_positions:
                diff = abs(py - cur_y)
                if diff <= snap and (best_diff is None or diff < best_diff):
                    best_diff = diff
                    best_y = py
            if best_diff is not None:
                snapped = True
                if side == "top":
                    y1 = best_y
                else:
                    y2 = best_y
        else:
            cur_x = x1 if side == "left" else x2
            best_x = cur_x
            best_diff = None
            for px in candidate_positions:
                diff = abs(px - cur_x)
                if diff <= snap and (best_diff is None or diff < best_diff):
                    best_diff = diff
                    best_x = px
            if best_diff is not None:
                snapped = True
                if side == "left":
                    x1 = best_x
                else:
                    x2 = best_x

        return (x1, y1, x2, y2), snapped

    # -------- 다시 그리기 --------

    def redraw_shape(self, shape):
        self.canvas.delete(shape.rect_id)
        for lid in shape.side_ids.values():
            self.canvas.delete(lid)
        for part in shape.dim_items.values():
            for lid in part["lines"] + part["ticks"] + [part["text"]]:
                self.canvas.delete(lid)

        x1, y1, x2, y2 = shape.coords
        color = shape.color

        rect_id = self.canvas.create_rectangle(
            x1, y1, x2, y2,
            outline=color,
            width=2,
            dash=() if color == "black" else (3, 2)
        )
        top_id = self.canvas.create_line(x1, y1, x2, y1, fill=color, width=3)
        bottom_id = self.canvas.create_line(x1, y2, x2, y2, fill=color, width=3)
        left_id = self.canvas.create_line(x1, y1, x1, y2, fill=color, width=3)
        right_id = self.canvas.create_line(x2, y1, x2, y2, fill=color, width=3)
        side_ids = {"top": top_id, "bottom": bottom_id, "left": left_id, "right": right_id}

        for side_name, lid in side_ids.items():
            self.canvas.addtag_withtag(f"shape_{shape.shape_id}", lid)
            self.canvas.addtag_withtag(f"side_{shape.shape_id}_{side_name}", lid)

        dim_items = self.draw_dimensions_for_shape(shape.shape_id, x1, y1, x2, y2, color=color)

        shape.rect_id = rect_id
        shape.side_ids = side_ids
        shape.dim_items = dim_items

        self.bring_shape_to_front(shape)

    # -------- 치수 클릭 (공유벽 고정 규칙 포함) --------

    def on_dim_width_click(self, event):
        closest_id = event.widget.find_closest(event.x, event.y)[0]
        tags = self.canvas.gettags(closest_id)
        shape_id = None
        for t in tags:
            if t.startswith("dim_width_"):
                shape_id = int(t.split("_")[2])
                break
        if shape_id is None:
            return
        shape = self.get_shape_by_id(shape_id)
        if not shape or not shape.editable:
            return

        x1, y1, x2, y2 = shape.coords
        cur_w_m = self.pixel_to_meter(x2 - x1)

        new_w_m = simpledialog.askfloat(
            "가로 길이 변경",
            f"새 가로 길이({self.unit})를 입력하세요 (현재: {cur_w_m:.2f} {self.unit}):",
            minvalue=0.01
        )
        if new_w_m is None:
            return

        shared = self.find_shared_vertical_edges(shape)
        shared_count = sum(1 for k in ("left", "right") if shared[k])

        if shared_count > 1:
            messagebox.showwarning(
                "변경 불가",
                "좌우 변이 모두 다른 도형과 공유되고 있어 가로 길이를 변경할 수 없습니다."
            )
            return

        self.push_history()
        self.set_active_shape(shape)

        new_w_px = self.meter_to_pixel(new_w_m)
        new_x1 = x1
        new_x2 = x2

        if shared_count == 1:
            if shared["left"]:
                new_x1 = x1
                new_x2 = x1 + new_w_px
            else:
                new_x2 = x2
                new_x1 = x2 - new_w_px
        else:
            new_x1 = x1
            new_x2 = x1 + new_w_px

        min_size_px = 20
        if new_x2 - new_x1 < min_size_px:
            if shared_count == 1 and shared["right"]:
                new_x1 = new_x2 - min_size_px
            else:
                new_x2 = new_x1 + min_size_px

        shape.coords = (new_x1, y1, new_x2, y2)
        self.redraw_shape(shape)
        self.app.update_selected_area_label(self)

    def on_dim_height_click(self, event):
        closest_id = event.widget.find_closest(event.x, event.y)[0]
        tags = self.canvas.gettags(closest_id)
        shape_id = None
        for t in tags:
            if t.startswith("dim_height_"):
                shape_id = int(t.split("_")[2])
                break
        if shape_id is None:
            return
        shape = self.get_shape_by_id(shape_id)
        if not shape or not shape.editable:
            return

        x1, y1, x2, y2 = shape.coords
        cur_h_m = self.pixel_to_meter(y2 - y1)

        new_h_m = simpledialog.askfloat(
            "세로 길이 변경",
            f"새 세로 길이({self.unit})를 입력하세요 (현재: {cur_h_m:.2f} {self.unit}):",
            minvalue=0.01
        )
        if new_h_m is None:
            return

        shared = self.find_shared_horizontal_edges(shape)
        shared_count = sum(1 for k in ("top", "bottom") if shared[k])

        if shared_count > 1:
            messagebox.showwarning(
                "변경 불가",
                "위·아래 변이 모두 다른 도형과 공유되고 있어 세로 길이를 변경할 수 없습니다."
            )
            return

        self.push_history()
        self.set_active_shape(shape)

        new_h_px = self.meter_to_pixel(new_h_m)
        new_y1 = y1
        new_y2 = y2

        if shared_count == 1:
            if shared["top"]:
                new_y1 = y1
                new_y2 = y1 + new_h_px
            else:
                new_y2 = y2
                new_y1 = y2 - new_h_px
        else:
            new_y1 = y1
            new_y2 = y1 + new_h_px

        min_size_px = 20
        if new_y2 - new_y1 < min_size_px:
            if shared_count == 1 and shared["bottom"]:
                new_y1 = new_y2 - min_size_px
            else:
                new_y2 = new_y1 + min_size_px

        shape.coords = (x1, new_y1, x2, new_y2)
        self.redraw_shape(shape)
        self.app.update_selected_area_label(self)

    # -------- 공간 텍스트 수정 --------

    def _find_space_label_by_item(self, item_id):
        for lab in self.generated_space_labels:
            if item_id in (
                lab["name_id"],
                lab["heat_norm_id"],
                lab["heat_equip_id"],
                lab["area_id"],
            ):
                return lab
        return None

    def on_space_name_click(self, event):
        item_id = event.widget.find_closest(event.x, event.y)[0]
        lab = self._find_space_label_by_item(item_id)
        if not lab:
            return
        # ensure hvac fields exist so popup initialization can rely on them
        try:
            cur_hvac_def = int(lab.get("hvac_type", 1))
        except Exception:
            cur_hvac_def = 1
        if 'hvac_text' not in lab:
            try:
                lab['hvac_text'] = f"{cur_hvac_def}. {HVAC_NAMES.get(cur_hvac_def, '')}"
            except Exception:
                lab['hvac_text'] = None
        if 'hvac_detail' not in lab:
            lab['hvac_detail'] = None
        # current name and hvac
        old = self.canvas.itemcget(lab["name_id"], "text")
        # extract bare name (remove existing hvac suffix like 'Room 1(1. 중앙공조)'
        # and trailing detail like '_1.PAC(냉방)')
        m = re.match(r'^(.*?)(?:\s*\(\d+\..*?\))?(?:_\d+\..*)?$', old)
        base_name = m.group(1).strip() if m else old

        # popup dialog with entry + combobox
        dlg = tk.Toplevel(self.canvas.master)
        dlg.transient(self.canvas.master)
        dlg.title("공간편집")
        # make popup wider so controls don't overlap
        try:
            dlg.geometry("700x420")
        except Exception:
            pass
        # Reserve a right-side column for the CSV table so left-side controls keep their positions
        try:
            dlg.grid_columnconfigure(0, weight=0)
            dlg.grid_columnconfigure(1, weight=0)
            # reserve a fixed min width for the table column so adding the table won't shift left widgets
            dlg.grid_columnconfigure(2, weight=0, minsize=480)
        except Exception:
            pass
        tk.Label(dlg, text="공간이름:").grid(row=0, column=0, padx=6, pady=6)
        name_entry = tk.Entry(dlg, width=30)
        name_entry.grid(row=0, column=1, padx=6, pady=6, sticky='w')
        name_entry.insert(0, base_name)

        tk.Label(dlg, text="공조방식:").grid(row=1, column=0, padx=6, pady=6)
        from tkinter import ttk
        hvac_var = tk.StringVar()
        # make combobox width match name_entry and align left
        combo = ttk.Combobox(dlg, textvariable=hvac_var, state='readonly', width=30)
        combo['values'] = [f"{k}. {v}" for k, v in HVAC_NAMES.items()]
        # initialize HVAC combobox display from stored lab values
        cur_hvac = lab.get("hvac_type", 1)
        hvac_text = lab.get('hvac_text', None)
        try:
            vals = list(combo['values'])
            if hvac_text and hvac_text in vals:
                combo.current(vals.index(hvac_text))
                hvac_var.set(hvac_text)
                # force visible text after the popup is mapped to avoid readonly rendering quirks
                dlg.after(10, lambda v=hvac_text: combo.set(v))
            else:
                # format from numeric hvac_type
                try:
                    hv_num = int(cur_hvac)
                except Exception:
                    hv_num = 1
                display = f"{hv_num}. {HVAC_NAMES.get(hv_num, '')}"
                if display in vals:
                    combo.current(vals.index(display))
                else:
                    combo.current(0)
                hvac_var.set(display)
                dlg.after(10, lambda v=display: combo.set(v))
        except Exception:
            try:
                combo.current(0)
            except Exception:
                pass
        combo.grid(row=1, column=1, padx=6, pady=6, sticky='w')

        # 공조 상세 콤보박스 추가 (will be enabled only when HVAC == 2)
        tk.Label(dlg, text="공조상세:").grid(row=2, column=0, padx=6, pady=6)
        hvac_detail_var = tk.StringVar()
        # match width with other controls and align left
        detail_combo = ttk.Combobox(dlg, textvariable=hvac_detail_var, state='readonly', width=30)
        # show stripped text (no numeric prefix)
        detail_combo['values'] = [
            "PAC(냉방)",
            "PAC(냉난방)",
            "EHP",
            "항온항습기",
        ]
        # set initial hvac_detail selection from stored lab value if any
        cur_detail = lab.get("hvac_detail", None)
        try:
            vals_d = list(detail_combo['values'])
            if cur_detail is not None and isinstance(cur_detail, int) and 1 <= cur_detail <= len(vals_d):
                di = int(cur_detail) - 1
                detail_combo.current(di)
                try:
                    dval = vals_d[di]
                    hvac_detail_var.set(dval)
                    dlg.after(10, lambda v=dval: detail_combo.set(v))
                except Exception:
                    pass
            else:
                try:
                    # stored hvac_detail_text is already stripped (no prefix)
                    prev = lab.get('hvac_detail_text', None)
                    if prev and prev in vals_d:
                        idx = vals_d.index(prev)
                        detail_combo.current(idx)
                        hvac_detail_var.set(prev)
                        dlg.after(10, lambda v=prev: detail_combo.set(v))
                    else:
                        detail_combo.set("")
                except Exception:
                    detail_combo.set("")
        except Exception:
            detail_combo.set("")
        # enable/disable detail_combo depending on current hvac type
        try:
            if int(cur_hvac) == 2:
                detail_combo.configure(state='readonly')
                if not cur_detail:
                    detail_combo.current(0)
            else:
                detail_combo.configure(state='disabled')
        except Exception:
            detail_combo.configure(state='disabled')
        detail_combo.grid(row=2, column=1, padx=6, pady=6, sticky='w')

        # compute and show total heat (kW) beneath the detail combobox
        total_kw = None
        try:
            area_text = self.canvas.itemcget(lab["area_id"], "text")
            norm_text = self.canvas.itemcget(lab["heat_norm_id"], "text")
            equip_text = self.canvas.itemcget(lab["heat_equip_id"], "text")

            def _extract_first_float(s: str) -> float:
                if not s:
                    return 0.0
                for tok in s.replace(',', ' ').split():
                    try:
                        return float(tok)
                    except Exception:
                        continue
                return 0.0

            # area might be like '100.00 m²' or similar
            area_val = _extract_first_float(area_text)
            norm_v = _extract_first_float(norm_text)
            equip_v = _extract_first_float(equip_text)

            total_kw = area_val * (norm_v + equip_v) / 1000.0
            heat_label = tk.Label(dlg, text=f"총 발열량: {total_kw:.3f} kW")
            heat_label.grid(row=3, column=0, columnspan=2, padx=6, pady=(4, 6), sticky='w')
            status_label = tk.Label(dlg, text="", fg="gray")
            status_label.grid(row=4, column=0, columnspan=2, padx=6, pady=(0,6), sticky='w')
            # prepare CSV table/placeholders (will be updated dynamically)
            csv_shown = None
            tbl = None
            csv_msg = None
            # fix vertical spacing of left-side rows (name, hvac, detail, heat_label)
            fixed_row_mins = {}
            try:
                dlg.update_idletasks()
                # map rows to widgets we want to lock and record initial heights
                row_widget_map = {
                    0: name_entry,
                    1: combo,
                    2: detail_combo,
                    3: heat_label
                }
                for r, w in row_widget_map.items():
                    try:
                        h = max(18, int(w.winfo_reqheight()) + 6)
                        fixed_row_mins[r] = h
                        dlg.grid_rowconfigure(r, minsize=h, weight=0)
                    except Exception:
                        try:
                            dlg.grid_rowconfigure(r, weight=0)
                        except Exception:
                            pass
            except Exception:
                fixed_row_mins = {}

            # table container on the right (keeps the left controls fixed when table appears/disappears)
            table_frame = tk.Frame(dlg, bd=0)
            try:
                table_frame.grid(row=0, column=2, rowspan=7, padx=6, pady=6, sticky='nsew')
            except Exception:
                try:
                    table_frame.grid(row=0, column=2, rowspan=7, padx=6, pady=6)
                except Exception:
                    pass

            # helper to make the Treeview cells (second column) editable inline
            def make_tree_editable(tv):
                # tv: Treeview instance
                def _on_double_click(event):
                    try:
                        row_id = tv.identify_row(event.y)
                        col = tv.identify_column(event.x)
                        # only allow editing the second column (value)
                        if not row_id or col != '#2':
                            return
                        bbox = tv.bbox(row_id, column=col)
                        if not bbox:
                            return
                        x, y, w, h = bbox
                        # get current value
                        vals = list(tv.item(row_id, 'values'))
                        cur = vals[1] if len(vals) > 1 else ''
                        # create Entry overlay
                        entry = tk.Entry(table_frame)
                        entry.insert(0, cur)
                        # place relative to treeview widget
                        # translate bbox x,y to table_frame coordinates
                        try:
                            # tv.winfo_rootx/winfo_rooty not used because placing in same parent simplifies
                            entry.place(x=x, y=y, width=w, height=h)
                        except Exception:
                            entry.place(x=x, y=y, width=w, height=h)
                        entry.focus_set()

                        def _save(e=None):
                            try:
                                new = entry.get()
                                vals2 = list(tv.item(row_id, 'values'))
                                if len(vals2) < 2:
                                    # pad
                                    while len(vals2) < 2:
                                        vals2.append('')
                                vals2[1] = new
                                tv.item(row_id, values=vals2)
                                # update underlying csv_shown if present (keep sync)
                                try:
                                    children = list(tv.get_children())
                                    idx = children.index(row_id)
                                    if csv_shown and 'values' in csv_shown and 0 <= idx < len(csv_shown['values']):
                                        t0 = csv_shown['values'][idx][0] if len(csv_shown['values'][idx]) > 0 else ''
                                        csv_shown['values'][idx] = (t0, new)
                                        # If the edited row is the first row (quantity), recompute target and update table column
                                        if idx == 0:
                                            try:
                                                qty = int(float(new)) if new is not None and new != '' else 1
                                            except Exception:
                                                try:
                                                    qty = int(new)
                                                except Exception:
                                                    qty = 1
                                            try:
                                                # recompute target and choose new column from preferred columns if available
                                                if total_kw is None:
                                                    tval = 0.0
                                                else:
                                                    tval = float(total_kw) / max(1, qty)
                                                # find headers matching sel_detail
                                                sel_detail_local = None
                                                try:
                                                    sel_detail_local = detail_combo.get().strip() if detail_combo.get() else None
                                                except Exception:
                                                    sel_detail_local = None
                                                local_rows = getattr(self.app, 'last_csv_rows', None)
                                                headers_local = local_rows[0] if local_rows else []
                                                preferred = []
                                                if sel_detail_local:
                                                    for ci, h in enumerate(headers_local):
                                                        try:
                                                            if sel_detail_local.lower() in str(h).lower():
                                                                preferred.append(ci)
                                                        except Exception:
                                                            continue
                                                # build candidate list from preferred columns; if none, leave unchanged
                                                cand = []
                                                if preferred:
                                                    for ci in preferred:
                                                        try:
                                                            v = float(local_rows[1][ci]) if local_rows and ci < len(local_rows[1]) else None
                                                            if v is not None:
                                                                cand.append((ci, v))
                                                        except Exception:
                                                            continue
                                                # select column whose second-row value is > tval and closest
                                                chosen = None
                                                if cand:
                                                    greater_local = [c for c in cand if c[1] > tval]
                                                    if greater_local:
                                                        chosen = min(greater_local, key=lambda x: x[1])
                                                    else:
                                                        lesser_local = [c for c in cand if c[1] <= tval]
                                                        if lesser_local:
                                                            chosen = min(lesser_local, key=lambda x: abs(x[1] - tval))
                                                # if chosen found, update csv_shown values to that column
                                                if chosen:
                                                    nci = chosen[0]
                                                    try:
                                                        new_values = [((r[0] if len(r) > 0 else ''), (r[nci] if nci < len(r) else '')) for r in (local_rows if local_rows else [])]
                                                        csv_shown['col_index'] = nci
                                                        csv_shown['header'] = headers_local[nci] if nci < len(headers_local) else csv_shown.get('header', '')
                                                        csv_shown['values'] = new_values
                                                        # ensure quantity row is first
                                                        try:
                                                            csv_shown['values'].insert(0, ("대수(Q'ty)", str(qty)))
                                                        except Exception:
                                                            pass
                                                        # refresh treeview display
                                                        try:
                                                            for it in tv.get_children():
                                                                tv.delete(it)
                                                            for idx2, (t0, v0) in enumerate(csv_shown['values']):
                                                                iid2 = tv.insert('', tk.END, values=(t0, v0))
                                                                if idx2 == 0:
                                                                    try:
                                                                        tv.item(iid2, tags=('qty',))
                                                                    except Exception:
                                                                        pass
                                                            try:
                                                                tv.tag_configure('qty', background='#3399ff', foreground='white')
                                                            except Exception:
                                                                pass
                                                        except Exception:
                                                            pass
                                                    except Exception:
                                                        pass
                                            except Exception:
                                                pass
                                except Exception:
                                    pass
                            except Exception:
                                pass
                            try:
                                entry.destroy()
                            except Exception:
                                pass

                        entry.bind('<Return>', _save)
                        entry.bind('<FocusOut>', _save)
                    except Exception:
                        return

                try:
                    tv.unbind('<Double-1>')
                except Exception:
                    pass
                tv.bind('<Double-1>', _on_double_click)

            def update_csv_table():
                nonlocal csv_shown, tbl, total_kw
                # re-apply fixed row min sizes so left-side vertical spacing does not change
                try:
                    for rr, hh in fixed_row_mins.items():
                        try:
                            dlg.grid_rowconfigure(rr, minsize=hh, weight=0)
                        except Exception:
                            pass
                except Exception:
                    pass
                try:
                    # CSV is stored on the application object (self.app)
                    rows = getattr(self.app, 'last_csv_rows', None)
                    # check current hvac selection (kept for informational use)
                    try:
                        cur_sel = combo.get()
                        hv = int(cur_sel.split('.')[0]) if cur_sel else int(lab.get('hvac_type', 1))
                    except Exception:
                        hv = int(lab.get('hvac_type', 1)) if lab.get('hvac_type', None) is not None else 1

                    # If HVAC is not 개별공조(2), do not show CSV-derived table here
                    try:
                        if int(hv) != 2:
                            if tbl is not None:
                                try:
                                    tbl.destroy()
                                except Exception:
                                    pass
                                tbl = None
                            csv_shown = None
                            # ensure no csv_msg is shown
                            try:
                                if csv_msg is not None:
                                    csv_msg.destroy()
                            except Exception:
                                pass
                            csv_msg = None
                            return
                    except Exception:
                        pass

                    # If CSV rows exist and total_kw is available, proceed
                    if not rows or total_kw is None or len(rows) < 2:
                        # no CSV rows or no total: remove table if present
                        if tbl is not None:
                            try:
                                tbl.destroy()
                            except Exception:
                                pass
                            tbl = None
                        csv_shown = None
                        # if no CSV loaded, show an instruction
                        try:
                            if not rows:
                                if csv_msg is None:
                                    csv_msg = tk.Label(dlg, text="CSV가 로드되지 않았습니다. 툴바의 'CSV로드'로 파일을 먼저 불러오세요.", fg="gray")
                                    csv_msg.grid(row=4, column=0, columnspan=2, padx=6, pady=(2,6), sticky='w')
                            else:
                                if csv_msg is not None:
                                    try:
                                        csv_msg.destroy()
                                    except Exception:
                                        pass
                                    csv_msg = None
                        except Exception:
                            pass
                        return

                    headers = rows[0]
                    second = rows[1]
                    # prefer columns where the CSV first-row header matches the selected 공조상세
                    try:
                        # if HVAC is not 개별공조 (2), ignore detail_combo value even if present
                        if int(hv) != 2:
                            sel_detail = None
                        else:
                            sel_detail = detail_combo.get().strip() if detail_combo.get() else None
                    except Exception:
                        sel_detail = None

                    headers = rows[0]
                    preferred_cols = []
                    if sel_detail:
                        for ci, h in enumerate(headers):
                            try:
                                if sel_detail.lower() in str(h).lower():
                                    preferred_cols.append(ci)
                            except Exception:
                                continue

                    candidates = []
                    # if sel_detail provided but no preferred columns found, do not show table
                    if sel_detail and not preferred_cols:
                        # No header match: show table with empty values and DO NOT run fallback selection.
                        try:
                            first_col_vals = [((r[0] if len(r) > 0 else ''), '') for r in rows]
                        except Exception:
                            first_col_vals = []
                        csv_shown = {
                            'col_index': None,
                            'header': sel_detail if sel_detail else '',
                            'values': first_col_vals
                        }
                        # remove any CSV-not-loaded message
                        try:
                            if csv_msg is not None:
                                try:
                                    csv_msg.destroy()
                                except Exception:
                                    pass
                                csv_msg = None
                        except Exception:
                            pass
                        # create or update treeview immediately with empty values and return
                        try:
                            from tkinter import ttk
                            if tbl is None:
                                tbl = ttk.Treeview(table_frame, columns=("title", "value"), show='headings', height=15)
                                first_col_name = rows[0][0] if rows and len(rows) > 0 and len(rows[0]) > 0 else ""
                                try:
                                    tbl.heading('title', text=first_col_name)
                                except Exception:
                                    tbl.heading('title', text='')
                                tbl.heading('value', text=csv_shown['header'])
                                tbl.column('title', width=180, anchor='w')
                                tbl.column('value', width=300, anchor='w')
                                try:
                                    tbl.pack(fill=tk.BOTH, expand=True)
                                except Exception:
                                    tbl.grid(row=0, column=0, sticky='nsew')
                                try:
                                    make_tree_editable(tbl)
                                except Exception:
                                    pass
                            else:
                                tbl.heading('value', text=csv_shown['header'])
                                for it in tbl.get_children():
                                    tbl.delete(it)
                            for title, val in csv_shown['values']:
                                tbl.insert('', tk.END, values=(title, val))
                            try:
                                status_label.configure(text=f"CSV 로드: {len(rows)}행, 선택열 없음 (공조상세 일치 없음)")
                            except Exception:
                                pass
                        except Exception:
                            pass
                        return

                    # if we have preferred columns from header matching, only consider them
                    if preferred_cols:
                        for ci in preferred_cols:
                            try:
                                v = float(second[ci]) if ci < len(second) else None
                                if v is not None:
                                    candidates.append((ci, v))
                            except Exception:
                                continue
                    else:
                        for ci, val in enumerate(second):
                            try:
                                v = float(val)
                                candidates.append((ci, v))
                            except Exception:
                                continue

                    if not candidates:
                        if tbl is not None:
                            try:
                                tbl.destroy()
                            except Exception:
                                pass
                            tbl = None
                        csv_shown = None
                        return

                    # Special selection when we have preferred columns (header match)
                    import math
                    # if user previously saved a quantity for this lab, prefer it
                    stored_qty = None
                    try:
                        if 'hvac_qty' in lab and lab.get('hvac_qty') is not None:
                            stored_qty = int(lab.get('hvac_qty'))
                    except Exception:
                        stored_qty = None
                    if preferred_cols and sel_detail:
                        # consider only the candidate values (from preferred columns)
                        vals_only = [c[1] for c in candidates]
                        max_val = max(vals_only) if vals_only else 0.0
                        # Assumptions:
                        # - If total_kw < max_val -> divide total by 2, use qty=2.
                        # - If total_kw > max_val -> compute ratio = total_kw / max_val,
                        #   take qty = ceil(ratio) + 2 (integer), then target = total_kw / qty.
                        # These are inferred from the user's description.
                        try:
                            if total_kw is None:
                                total_kw = 0.0
                        except Exception:
                            total_kw = 0.0

                        # if stored_qty exists, use it; else compute per original logic
                        if stored_qty is not None:
                            qty = stored_qty
                            target = total_kw / max(1, qty)
                        else:
                            if max_val > 0 and total_kw < max_val:
                                target = total_kw / 2.0
                                qty = 2
                            else:
                                if max_val > 0:
                                    ratio = total_kw / max_val
                                else:
                                    ratio = total_kw
                                qty = int(math.ceil(ratio)) + 2
                                if qty <= 0:
                                    qty = 2
                                target = total_kw / qty if qty != 0 else total_kw

                        # pick candidate column closest above target, else largest below
                        greater = [c for c in candidates if c[1] >= target]
                        if greater:
                            best = min(greater, key=lambda x: x[1])
                        else:
                            lesser = [c for c in candidates if c[1] < target]
                            if lesser:
                                best = max(lesser, key=lambda x: x[1])
                            else:
                                best = candidates[0]
                        ci, cv = best
                        csv_shown = {
                            'col_index': ci,
                            'header': headers[ci] if ci < len(headers) else f'C{ci+1}',
                            # values: tuples of (first-column title, selected-column value) per row
                            'values': [((r[0] if len(r) > 0 else ''), (r[ci] if ci < len(r) else '')) for r in rows]
                        }
                        # prepend quantity row with computed qty
                        try:
                            # if user has a stored qty prefer that display
                            csv_shown['values'].insert(0, ("대수(Q'ty)", str(qty)))
                        except Exception:
                            pass
                    else:
                        # default selection logic (no header-priority special rules)
                        greater = [c for c in candidates if c[1] >= total_kw]
                        if greater:
                            best = min(greater, key=lambda x: x[1])
                        else:
                            lesser = [c for c in candidates if c[1] < total_kw]
                            if lesser:
                                best = max(lesser, key=lambda x: x[1])
                            else:
                                best = candidates[0]
                        ci, cv = best
                        csv_shown = {
                            'col_index': ci,
                            'header': headers[ci] if ci < len(headers) else f'C{ci+1}',
                            # values: tuples of (first-column title, selected-column value) per row
                            'values': [((r[0] if len(r) > 0 else ''), (r[ci] if ci < len(r) else '')) for r in rows]
                        }
                        # If we found preferred columns by header match, prepend a quantity row as before
                        try:
                            if preferred_cols:
                                try:
                                    stored_qty2 = int(lab.get('hvac_qty')) if lab.get('hvac_qty') is not None else None
                                except Exception:
                                    stored_qty2 = None
                                if stored_qty2 is not None:
                                    csv_shown['values'].insert(0, ("대수(Q'ty)", str(stored_qty2)))
                                else:
                                    csv_shown['values'].insert(0, ("대수(Q'ty)", "1"))
                        except Exception:
                            pass

                    # create or update treeview
                    try:
                        from tkinter import ttk
                        if tbl is None:
                            # place Treeview inside reserved table_frame so left controls don't shift
                            tbl = ttk.Treeview(table_frame, columns=("title", "value"), show='headings', height=15)
                            # first column shows the row title (CSV first column), second shows selected value
                            first_col_name = headers[0] if headers and len(headers) > 0 else ""
                            try:
                                tbl.heading('title', text=first_col_name)
                            except Exception:
                                tbl.heading('title', text='')
                            tbl.heading('value', text=csv_shown['header'])
                            tbl.column('title', width=180, anchor='w')
                            tbl.column('value', width=300, anchor='w')
                            try:
                                tbl.pack(fill=tk.BOTH, expand=True)
                            except Exception:
                                tbl.grid(row=0, column=0, sticky='nsew')
                            # make cells editable
                            try:
                                make_tree_editable(tbl)
                            except Exception:
                                pass
                        else:
                            tbl.heading('value', text=csv_shown['header'])
                            for it in tbl.get_children():
                                tbl.delete(it)

                        # remove any CSV-not-loaded message
                        try:
                            if csv_msg is not None:
                                try:
                                    csv_msg.destroy()
                                except Exception:
                                    pass
                                csv_msg = None
                        except Exception:
                            pass

                        for idx, (title, val) in enumerate(csv_shown['values']):
                            iid = tbl.insert('', tk.END, values=(title, val))
                            # tag the first row (quantity row) to have a blue background
                            if idx == 0:
                                try:
                                    tbl.item(iid, tags=('qty',))
                                except Exception:
                                    pass
                        try:
                            tbl.tag_configure('qty', background='#3399ff', foreground='white')
                        except Exception:
                            pass
                        # ensure popup width can contain the table
                        try:
                            dlg.update_idletasks()
                            req = table_frame.winfo_reqwidth()
                            cur_w = dlg.winfo_width()
                            # reserve ~380px for left controls; expand dlg width if table would be clipped
                            min_total = req + 380
                            if cur_w < min_total:
                                try:
                                    dlg.geometry(f"{min_total}x{dlg.winfo_height()}")
                                except Exception:
                                    pass
                        except Exception:
                            pass

                        # update status label if present
                        try:
                            status_label.configure(text=f"CSV 로드: {len(rows)}행, 선택열: {csv_shown['header']}")
                        except Exception:
                            pass
                    except Exception:
                        # ensure no half-created widget remains
                        try:
                            if tbl is not None:
                                tbl.destroy()
                        except Exception:
                            pass
                        tbl = None
                except Exception:
                    csv_shown = None
                    try:
                        if tbl is not None:
                            tbl.destroy()
                    except Exception:
                        pass
                    tbl = None

            # schedule initial csv table update after the popup widgets settle
            try:
                dlg.after(50, lambda: (update_csv_table()))
            except Exception:
                try:
                    update_csv_table()
                except Exception:
                    pass
        except Exception:
            # fallback: show nothing if parsing fails
            heat_label = tk.Label(dlg, text="총 발열량: - kW")
            heat_label.grid(row=3, column=0, columnspan=2, padx=6, pady=(4, 6), sticky='w')
            csv_shown = None

    # when HVAC selection changes, toggle the detail combobox
        def _on_hvac_change(event=None):
            sel = combo.get()
            try:
                hv = int(sel.split('.')[0])
            except Exception:
                hv = 1
            if hv == 2:
                detail_combo.configure(state='readonly')
                # if no previous detail, default to first
                if not detail_combo.get():
                    detail_combo.current(0)
            else:
                # clear detail selection and disable
                detail_combo.set("")
                detail_combo.configure(state='disabled')
            # update CSV table view when HVAC selection changes
            try:
                update_csv_table()
            except Exception:
                pass

        combo.bind("<<ComboboxSelected>>", _on_hvac_change)
        # also update CSV table when 공조상세 selection changes
        try:
            detail_combo.bind("<<ComboboxSelected>>", lambda e: update_csv_table())
        except Exception:
            pass

        def on_ok():
            new_name = name_entry.get().strip()
            if not new_name:
                return
            sel = combo.get()
            # determine hvac number: prefer parsed combo, fallback to existing lab value
            try:
                if sel:
                    num = int(sel.split('.')[0])
                else:
                    num = int(lab.get('hvac_type', 1))
            except Exception:
                try:
                    num = int(lab.get('hvac_type', 1))
                except Exception:
                    num = 1
            # detail: determine numeric index from combobox (values are stripped text)
            dsel = detail_combo.get()
            prev_detail = lab.get('hvac_detail', None)
            dnum = prev_detail
            try:
                if dsel:
                    vals_d = list(detail_combo['values'])
                    if dsel in vals_d:
                        dnum = vals_d.index(dsel) + 1
                    else:
                        # fallback: try matching after removing possible 'N. ' prefix from candidates
                        stripped = [v[3:].strip() if len(v) > 3 and v[1] == '.' else v for v in vals_d]
                        if dsel in stripped:
                            dnum = stripped.index(dsel) + 1
                        else:
                            dnum = prev_detail
                else:
                    dnum = prev_detail
            except Exception:
                dnum = prev_detail
            # persist: do NOT display hvac or detail on palette; only store values in lab
            full = new_name
            self.push_history()
            self.canvas.itemconfigure(lab["name_id"], text=full)
            lab["hvac_type"] = num
            # store full display text of the selected HVAC (e.g. '2. 개별공조')
            try:
                lab["hvac_text"] = combo.get().strip() if combo.get() else None
            except Exception:
                lab["hvac_text"] = None
            # store hvac_detail numeric and stripped text; clear if hvac != 2
            try:
                if int(num) == 2:
                    if dnum is None:
                        lab["hvac_detail"] = prev_detail if prev_detail is not None else 1
                    else:
                        lab["hvac_detail"] = int(dnum)
                    # store stripped text (no numeric prefix)
                    try:
                        lab["hvac_detail_text"] = detail_combo.get().strip() if detail_combo.get() else lab.get('hvac_detail_text', None)
                    except Exception:
                        lab["hvac_detail_text"] = lab.get('hvac_detail_text', None)
                else:
                    lab["hvac_detail"] = None
                    lab["hvac_detail_text"] = None
            except Exception:
                lab["hvac_detail"] = None
                lab["hvac_detail_text"] = None
            # If csv_shown exists and its first row is the quantity row, persist that quantity into the label
            try:
                if csv_shown and 'values' in csv_shown and len(csv_shown['values']) > 0:
                    first_title, first_val = csv_shown['values'][0]
                    if isinstance(first_title, str) and "대수" in first_title:
                        try:
                            qv = int(float(first_val)) if first_val is not None and str(first_val).strip() != '' else None
                        except Exception:
                            try:
                                qv = int(str(first_val))
                            except Exception:
                                qv = None
                        if qv is not None:
                            lab['hvac_qty'] = qv
            except Exception:
                pass
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        # if csv_shown prepared at popup creation time, place an initial empty table in the reserved frame
        if csv_shown:
            try:
                from tkinter import ttk
                tbl = ttk.Treeview(table_frame, columns=("title","value"), show='headings', height=15)
                try:
                    rows_top = getattr(self.app, 'last_csv_rows', None)
                    first_col_name = rows_top[0][0] if rows_top and len(rows_top) > 0 and len(rows_top[0]) > 0 else ""
                except Exception:
                    first_col_name = ""
                try:
                    tbl.heading('title', text=first_col_name)
                except Exception:
                    tbl.heading('title', text='')
                tbl.heading('value', text=csv_shown['header'])
                tbl.column('title', width=180, anchor='w')
                tbl.column('value', width=300, anchor='w')
                for idx, (title, val) in enumerate(csv_shown['values']):
                    iid = tbl.insert('', tk.END, values=(title, val))
                    if idx == 0:
                        try:
                            tbl.item(iid, tags=('qty',))
                        except Exception:
                            pass
                try:
                    tbl.pack(fill=tk.BOTH, expand=True)
                except Exception:
                    tbl.grid(row=0, column=0, sticky='nsew')
                # make editable and style first row
                try:
                    make_tree_editable(tbl)
                except Exception:
                    pass
                try:
                    tbl.tag_configure('qty', background='#3399ff', foreground='white')
                except Exception:
                    pass
                # resize dialog if table would be clipped
                try:
                    dlg.update_idletasks()
                    req = table_frame.winfo_reqwidth()
                    cur_w = dlg.winfo_width()
                    min_total = req + 380
                    if cur_w < min_total:
                        try:
                            dlg.geometry(f"{min_total}x{dlg.winfo_height()}")
                        except Exception:
                            pass
                except Exception:
                    pass
            except Exception:
                csv_shown = None

        ok_btn = tk.Button(dlg, text="확인", command=on_ok)
        ok_btn.grid(row=6, column=0, padx=6, pady=8)
        cancel_btn = tk.Button(dlg, text="취소", command=on_cancel)
        cancel_btn.grid(row=6, column=1, padx=6, pady=8)
        name_entry.focus_set()

    def on_space_heat_norm_click(self, event):
        item_id = event.widget.find_closest(event.x, event.y)[0]
        lab = self._find_space_label_by_item(item_id)
        if not lab:
            return
        old_text = self.canvas.itemcget(lab["heat_norm_id"], "text")
        try:
            num_str = old_text.split(":")[1].replace("W/m²", "").strip()
            old_val = float(num_str)
        except Exception:
            old_val = 0.0
        new_val = simpledialog.askfloat(
            "일반 발열량 변경",
            "새 일반 발열량 (W/m²)을 입력하세요:",
            initialvalue=old_val,
            minvalue=0.0
        )
        if new_val is None:
            return
        self.push_history()
        self.canvas.itemconfigure(
            lab["heat_norm_id"],
            text=f"Norm: {new_val:.2f} W/m²"
        )

    def on_space_heat_equip_click(self, event):
        item_id = event.widget.find_closest(event.x, event.y)[0]
        lab = self._find_space_label_by_item(item_id)
        if not lab:
            return
        old_text = self.canvas.itemcget(lab["heat_equip_id"], "text")
        try:
            num_str = old_text.split(":")[1].replace("W/m²", "").strip()
            old_val = float(num_str)
        except Exception:
            old_val = 0.0
        new_val = simpledialog.askfloat(
            "장비 발열량 변경",
            "새 장비 발열량 (W/m²)을 입력하세요:",
            initialvalue=old_val,
            minvalue=0.0
        )
        if new_val is None:
            return
        self.push_history()
        self.canvas.itemconfigure(
            lab["heat_equip_id"],
            text=f"Equip: {new_val:.2f} W/m²"
        )

    # -------- 오른쪽 클릭 --------

    def on_right_click(self, event):
        px, py = event.x, event.y
        shape, idx, cx, cy = self.detect_corner_under_mouse(px, py)
        if shape is not None:
            self.corner_menu_target_shape = shape
            try:
                self.corner_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.corner_menu.grab_release()
            return

        if not self.shapes:
            return

        best_corner = None
        best_d2 = None
        for shape in self.shapes:
            x1, y1, x2, y2 = shape.coords
            corners = [(x1, y1), (x2, y1), (x1, y2), (x2, y2)]
            for cx, cy in corners:
                d2 = (cx - px) ** 2 + (cy - py) ** 2
                if best_d2 is None or d2 < best_d2:
                    best_d2 = d2
                    best_corner = (cx, cy)

        if best_corner is None:
            return

        self.push_history()
        cx, cy = best_corner
        new_shape = self.create_rect_shape(cx, cy, px, py, editable=True, color="blue",
                                           push_to_history=False)
        self.set_active_shape(new_shape)

    # -------- Shapely 기반 자동 공간 생성 (텍스트 유지/새로 생성 규칙) --------

    def auto_generate_space_labels(self):
        if not self.shapes:
            messagebox.showinfo("자동생성", "도형이 없습니다.")
            return

        # 1. 모든 사각형의 경계선을 LineString으로 모음
        lines = []
        for s in self.shapes:
            x1, y1, x2, y2 = s.coords
            lines.append(LineString([[x1, y1], [x2, y1]]))
            lines.append(LineString([[x2, y1], [x2, y2]]))
            lines.append(LineString([[x2, y2], [x1, y2]]))
            lines.append(LineString([[x1, y2], [x1, y1]]))

        merged = unary_union(lines)
        polys = list(polygonize(merged))

        if not polys:
            messagebox.showinfo("자동생성", "밀폐된 공간을 찾지 못했습니다.")
            return

        valid_polys = []
        for p in polys:
            area_px2 = p.area
            area_m2 = area_px2 / (self.scale * self.scale)
            if area_m2 > 0.01:
                valid_polys.append((p, area_m2))

        if not valid_polys:
            messagebox.showinfo("자동생성", "유효한 공간이 없습니다.")
            return

        self.push_history()

        # 면적 기준 정렬
        valid_polys.sort(key=lambda x: x[1])

        # 기존 라벨의 텍스트 위치(캔버스 좌표) 및 텍스트 정보 목록
        existing_centers = []
        for lab in self.generated_space_labels:
            poly_old = lab["polygon"]
            try:
                name_coords = self.canvas.coords(lab["name_id"])
                if name_coords:
                    nx, ny = name_coords[0], name_coords[1]
                else:
                    rep_old = poly_old.representative_point()
                    nx, ny = rep_old.x, rep_old.y
            except Exception:
                rep_old = poly_old.representative_point()
                nx, ny = rep_old.x, rep_old.y

            name_text = self.canvas.itemcget(lab["name_id"], "text")
            heat_norm_text = self.canvas.itemcget(lab["heat_norm_id"], "text")
            heat_equip_text = self.canvas.itemcget(lab["heat_equip_id"], "text")
            room_number = None
            if name_text.lower().startswith("room"):
                try:
                    room_number = int(name_text.split()[1])
                except Exception:
                    room_number = None
            
            diffuser_ids = lab.get("diffuser_ids", [])
            existing_centers.append((lab, nx, ny, name_text, heat_norm_text, heat_equip_text, room_number, diffuser_ids))

        used_existing = []   
        new_labels = []

        max_room_index = 0
        for lab in self.generated_space_labels:
            name_text = self.canvas.itemcget(lab["name_id"], "text")
            if name_text.lower().startswith("room"):
                try:
                    idx = int(name_text.split()[1])
                    if idx > max_room_index:
                        max_room_index = idx
                except Exception:
                    pass

        next_room_index = max_room_index + 1 if max_room_index > 0 else 1

        for p, area_m2 in valid_polys:
            cent = p.centroid
            cx, cy = cent.x, cent.y

            matched = None
            matched_name = None
            matched_norm = None
            matched_equip = None
            matched_room_number = None
            matched_diffusers = []

            for lab, nx, ny, name_text, heat_norm_text, heat_equip_text, room_number, diffuser_ids in existing_centers:
                try:
                    if p.contains(Point(nx, ny)):
                        matched = lab
                        matched_name = name_text
                        matched_norm = heat_norm_text
                        matched_equip = heat_equip_text
                        matched_room_number = room_number
                        matched_diffusers = diffuser_ids
                        break
                except Exception:
                    continue

            if matched is not None:
                # 기존 라벨 유지, 면적만 갱신
                name_id = matched["name_id"]
                heat_norm_id = matched["heat_norm_id"]
                heat_equip_id = matched["heat_equip_id"]
                area_id = matched["area_id"]

                self.canvas.itemconfigure(area_id, text=f"{area_m2:.2f} m²")
                if matched_room_number is not None:
                    # preserve hvac_type if present in matched
                    # Do not display hvac on the palette; show only the room number/name
                    self.canvas.itemconfigure(name_id, text=f"Room {matched_room_number}")
                else:
                    # if matched_name has hvac in suffix preserve, else leave as-is
                    if '(' in matched_name and ')' in matched_name:
                        self.canvas.itemconfigure(name_id, text=matched_name)
                    else:
                        # preserve matched_name as-is (do not append hvac)
                        self.canvas.itemconfigure(name_id, text=matched_name)
                self.canvas.itemconfigure(heat_norm_id, text=matched_norm)
                self.canvas.itemconfigure(heat_equip_id, text=matched_equip)

                new_labels.append({
                    "polygon": p,
                    "name_id": name_id,
                    "heat_norm_id": heat_norm_id,
                    "heat_equip_id": heat_equip_id,
                    "area_id": area_id,
                    "diffuser_ids": matched_diffusers, # 기존 디퓨저 점 유지
                    "hvac_type": matched.get('hvac_type', 1) if isinstance(matched, dict) else 1
                })
                used_existing.append(matched)
            else:
                # 새 라벨
                name_text = f"Room {next_room_index}"
                next_room_index += 1

                heat_norm_text = "Norm: 0.00 W/m²"
                heat_equip_text = "Equip: 0.00 W/m²"
                area_text = f"{area_m2:.2f} m²"

                name_id = self.canvas.create_text(
                    cx, cy,
                    text=name_text,
                    fill="blue",
                    font=("Arial", 11, "bold"),
                    tags=("space_name",)
                )
                heat_norm_id = self.canvas.create_text(
                    cx, cy + 14,
                    text=heat_norm_text,
                    fill="darkred",
                    font=("Arial", 10),
                    tags=("space_heat_norm",)
                )
                heat_equip_id = self.canvas.create_text(
                    cx, cy + 28,
                    text=heat_equip_text,
                    fill="darkred",
                    font=("Arial", 10),
                    tags=("space_heat_equip",)
                )
                area_id = self.canvas.create_text(
                    cx, cy + 42,
                    text=area_text,
                    fill="green",
                    font=("Arial", 10)
                )

                # default hvac type = 1 (중앙공조)
                hvac_type = 1
                # display only base name on palette (do not append hvac)
                name_text_with_hvac = name_text
                new_labels.append({
                    "polygon": p,
                    "name_id": name_id,
                    "heat_norm_id": heat_norm_id,
                    "heat_equip_id": heat_equip_id,
                    "area_id": area_id,
                    "diffuser_ids": [],
                    "hvac_type": hvac_type,
                    # ensure hvac_text and hvac_detail are present for later popup uses
                    "hvac_text": f"{hvac_type}. {HVAC_NAMES.get(hvac_type, '')}",
                    "hvac_detail": None
                })
                # set displayed name (base name only)
                self.canvas.itemconfigure(name_id, text=name_text_with_hvac)

        # 기존 라벨 중 사용되지 않은 것 삭제
        for lab in self.generated_space_labels:
            if lab not in used_existing:
                self.canvas.delete(lab["name_id"])
                self.canvas.delete(lab["heat_norm_id"])
                self.canvas.delete(lab["heat_equip_id"])
                self.canvas.delete(lab["area_id"])
                # 디퓨저도 삭제
                if "diffuser_ids" in lab:
                    for did in lab["diffuser_ids"]:
                        self.canvas.delete(did)

        self.generated_space_labels = new_labels

    # -------- 일괄 발열량 적용 --------

    def apply_norm_to_all(self, value: float):
        if not self.generated_space_labels:
            return
        self.push_history()
        for lab in self.generated_space_labels:
            try:
                self.canvas.itemconfigure(lab["heat_norm_id"], text=f"Norm: {value:.2f} W/m²")
            except Exception:
                continue

    def apply_equip_to_all(self, value: float):
        if not self.generated_space_labels:
            return
        self.push_history()
        for lab in self.generated_space_labels:
            try:
                self.canvas.itemconfigure(lab["heat_equip_id"], text=f"Equip: {value:.2f} W/m²")
            except Exception:
                continue

    def compute_and_apply_supply_flow(self):
        if not self.generated_space_labels:
            return 0.0

        try:
            indoor_t = float(self.app.indoor_temp_entry.get())
            supply_t = float(self.app.supply_temp_entry.get())
        except Exception:
            messagebox.showerror("입력 오류", "실내/급기 온도를 올바르게 입력하세요.")
            return 0.0

        delta_t = indoor_t - supply_t
        if delta_t <= 0:
            messagebox.showerror("입력 오류", "실내 온도는 급기 온도보다 높아야 합니다.")
            return 0.0

        total_flow = 0
        for lab in self.generated_space_labels:
            try:
                area_text = self.canvas.itemcget(lab["area_id"], "text")
                area_val = 0.0
                for tok in area_text.split():
                    try:
                        area_val = float(tok)
                        break
                    except Exception:
                        continue

                norm_text = self.canvas.itemcget(lab["heat_norm_id"], "text")
                equip_text = self.canvas.itemcget(lab["heat_equip_id"], "text")
                
                def extract_num(s: str) -> float:
                    for part in s.replace(',', ' ').split():
                        try:
                            return float(part)
                        except Exception:
                            continue
                    return 0.0

                norm_v = extract_num(norm_text)
                equip_v = extract_num(equip_text)

                raw_flow = area_val * (norm_v + equip_v) * 860.0 / 1.2 / 0.24 / 1000.0 / delta_t
                flow_int = int(ceil(raw_flow))
                total_flow += flow_int

                flow_text = f"{flow_int:,}"

                # Try to update existing flow text if present
                updated = False
                if "flow_id" in lab:
                    try:
                        fid = lab["flow_id"]
                        if fid in self.canvas.find_all():
                            self.canvas.itemconfigure(fid, text=f"Flow: {flow_text} m3/hr")
                            updated = True
                        else:
                            # stale id, remove key
                            lab.pop("flow_id", None)
                    except Exception:
                        lab.pop("flow_id", None)

                if not updated:
                    # remove any stray flow texts near the area to avoid duplicates
                    try:
                        x, y = self.canvas.coords(lab["area_id"])
                        # small bbox around area text
                        bx1, by1 = x - 10, y
                        bx2, by2 = x + 10, y + 28
                        for item in self.canvas.find_overlapping(bx1, by1, bx2, by2):
                            try:
                                if self.canvas.type(item) == 'text':
                                    txt = self.canvas.itemcget(item, 'text')
                                    if isinstance(txt, str) and txt.strip().startswith('Flow:'):
                                        self.canvas.delete(item)
                            except Exception:
                                continue
                    except Exception:
                        pass

                    # create new flow text and tag it
                    try:
                        x, y = self.canvas.coords(lab["area_id"])
                        fid = self.canvas.create_text(x, y + 14, text=f"Flow: {flow_text} m3/hr",
                                                      fill="purple", font=("Arial", 10), tags=("flow",))
                        lab["flow_id"] = fid
                    except Exception:
                        pass
            except Exception:
                continue

        return total_flow

    # -------- 디퓨저 자동 배치 로직 --------

    def _decide_grid_rc(self, N: int, width: float, height: float):
        """N개의 점을 width x height 영역에 배치할 때 행(r), 열(c) 결정"""
        if N <= 0:
            return 0, 0
        if N == 1:
            return 1, 1

        best_r, best_c = 1, N
        target_ratio = (width / height) if height > 1e-6 else 1.0
        best_diff = None

        # 행을 1부터 N까지 변화시키며 최적 비율 찾기
        for r in range(1, N + 1):
            c = ceil(N / r)
            if r * c < N:
                continue
            grid_ratio = c / r
            diff = abs(grid_ratio - target_ratio)
            if best_diff is None or diff < best_diff:
                best_diff = diff
                best_r, best_c = r, c

        return best_r, best_c

    def _select_points_greedy_maxmin(self, points, k: int):
        """Select k points from list 'points' using greedy max-min (farthest-first)"""
        if k <= 0 or not points:
            return []
        pts = list(points)
        # start with point farthest from centroid
        cx = sum(p[0] for p in pts) / len(pts)
        cy = sum(p[1] for p in pts) / len(pts)
        dists = [((p[0]-cx)**2 + (p[1]-cy)**2, i) for i, p in enumerate(pts)]
        dists.sort(reverse=True)
        sel = [pts[dists[0][1]]]
        used = {dists[0][1]}
        while len(sel) < k and len(used) < len(pts):
            best_i = None
            best_min = -1
            for i, p in enumerate(pts):
                if i in used:
                    continue
                # distance to nearest selected
                min_d = min((p[0]-q[0])**2 + (p[1]-q[1])**2 for q in sel)
                if min_d > best_min:
                    best_min = min_d
                    best_i = i
            if best_i is None:
                break
            used.add(best_i)
            sel.append(pts[best_i])
        return sel

    def _select_with_min_separation(self, points, k: int, min_px: float):
        """Greedy selection ensuring each new point is at least min_px from existing picks.
        If full k can't be satisfied, the function will try to relax the min_px gradually.
        """
        if k <= 0 or not points:
            return []
        pts = list(points)
        # try with decreasing thresholds until we get k points or threshold reaches 0
        thr = float(min_px)
        while thr >= 0:
            sel = []
            # start with farthest-from-centroid seed
            cx = sum(p[0] for p in pts) / len(pts)
            cy = sum(p[1] for p in pts) / len(pts)
            dists = [((p[0]-cx)**2 + (p[1]-cy)**2, i) for i, p in enumerate(pts)]
            dists.sort(reverse=True)
            used_idx = set()
            sel.append(pts[dists[0][1]])
            used_idx.add(dists[0][1])
            for i, p in enumerate(pts):
                if len(sel) >= k:
                    break
                if i in used_idx:
                    continue
                min_d2 = min((p[0]-q[0])**2 + (p[1]-q[1])**2 for q in sel)
                if min_d2 >= thr*thr:
                    sel.append(p)
                    used_idx.add(i)
            if len(sel) >= k:
                return sel[:k]
            # relax threshold
            thr *= 0.8
            if thr < 1e-6:
                break
        # fallback: use greedy max-min to fill
        fallback = self._select_points_greedy_maxmin(pts, k)
        return fallback

    def _generate_diffuser_points_for_poly(self, poly, N: int):
        if N <= 0:
            return []

        # 1. Fix topology and compute safe interior area
        poly = poly.buffer(0)
        margin_m = 0.5
        margin_px = self.meter_to_pixel(margin_m)
        safe_poly = poly.buffer(-margin_px)
        if safe_poly.is_empty:
            safe_poly = poly

        # 2. Bounding box (pixel coordinates)
        minx, miny, maxx, maxy = poly.bounds
        width = maxx - minx
        height = maxy - miny
        if width <= 0 or height <= 0:
            rep = poly.representative_point()
            return [(rep.x, rep.y)]

        # 3. Grid spacing (0.5m) in pixels
        spacing = max(1.0, self.meter_to_pixel(0.5))

        import math

        # 4. Anchor grid to same reference used by draw_grid (first shape or origin)
        try:
            if self.shapes:
                anchor_x = float(self.shapes[0].coords[0])
                anchor_y = float(self.shapes[0].coords[1])
            else:
                anchor_x = 0.0
                anchor_y = 0.0
        except Exception:
            anchor_x = 0.0
            anchor_y = 0.0

        rem_x = anchor_x - math.floor(anchor_x / spacing) * spacing
        rem_y = anchor_y - math.floor(anchor_y / spacing) * spacing

        # 5. Grid indices covering bbox
        kmin = math.floor((minx - rem_x) / spacing)
        kmax = math.ceil((maxx - rem_x) / spacing)
        hmin = math.floor((miny - rem_y) / spacing)
        hmax = math.ceil((maxy - rem_y) / spacing)

        # 6. Build list of grid intersection points that lie inside safe_poly
        from shapely.geometry import Point as ShapelyPoint
        intersections = []
        for k in range(kmin, kmax + 1):
            x = k * spacing + rem_x
            for j in range(hmin, hmax + 1):
                y = j * spacing + rem_y
                pt = ShapelyPoint(x, y)
                if safe_poly.contains(pt):
                    intersections.append((x, y))

        # If no intersections, fallback to representative point
        if not intersections:
            rep = poly.representative_point()
            return [(rep.x, rep.y)] * min(N, 1)

        # Remove duplicates (round to int pixels)
        uniq = []
        seen = set()
        for (x, y) in intersections:
            key = (round(x, 3), round(y, 3))
            if key in seen:
                continue
            seen.add(key)
            uniq.append((x, y))
        intersections = uniq

        # 7. Decide r x c layout to try to distribute N points in rows/cols
        r, c = self._decide_grid_rc(N, width, height)
        # create ideal cell centers (in bbox coordinates)
        ideal_points = []
        if r > 0 and c > 0:
            dx = width / (c + 1)
            dy = height / (r + 1)
            for irow in range(1, r + 1):
                for icol in range(1, c + 1):
                    px = minx + icol * dx
                    py = miny + irow * dy
                    ideal_points.append((px, py))
        else:
            # fallback to centroid-based selection
            cx = sum(p[0] for p in intersections) / len(intersections)
            cy = sum(p[1] for p in intersections) / len(intersections)
            ideal_points = [(cx, cy)]

        # 8. For each ideal point, choose nearest unused grid intersection (prefer within same cell radius)
        selected = []
        used_idx = set()
        for ip in ideal_points:
            best_i = None
            best_d2 = float('inf')
            # first try to find intersection within the surrounding cell radius (in pixel distance dx/2,dy/2)
            radius_px = max(spacing, min(width, height))
            for idx, pt in enumerate(intersections):
                if idx in used_idx:
                    continue
                d2 = (pt[0] - ip[0]) ** 2 + (pt[1] - ip[1]) ** 2
                if d2 < best_d2:
                    best_d2 = d2
                    best_i = idx
            if best_i is not None:
                used_idx.add(best_i)
                selected.append(intersections[best_i])
            if len(selected) >= N:
                break

        # 9. If still fewer than N, fill remaining by greedy selection with minimum separation
        if len(selected) < N:
            remaining = [p for i, p in enumerate(intersections) if i not in used_idx]
            min_sep_m = 1.0
            min_sep_px = self.meter_to_pixel(min_sep_m)
            more = self._select_with_min_separation(remaining, N - len(selected), min_sep_px)
            selected.extend(more)

        # 10. Final trim and ensure uniqueness
        out = []
        seen2 = set()
        for (x, y) in selected:
            key = (round(x, 3), round(y, 3))
            if key in seen2:
                continue
            seen2.add(key)
            out.append((x, y))
            if len(out) >= N:
                break

        return out

    # Diagnostic: check diffusers in a named room
    def check_diffusers_in_room(self, room_name: str):
        """Return (assigned_count, inside_count, outside_count, list_of_outside_item_ids) for given room_name.
        - assigned_count: number of diffuser item ids stored in this room's lab entry
        - inside_count / outside_count: counts of ALL diffuser items (from any lab) whose centers are inside/outside the room polygon
        This lets us detect when diffusers exist on the canvas but are not assigned to the target lab (or vice versa).
        """
        inside = 0
        outside = 0
        outside_ids = []
        # find matching label
        target_lab = None
        for lab in self.generated_space_labels:
            try:
                name = self.canvas.itemcget(lab["name_id"], "text")
            except Exception:
                name = ""
            if name == room_name:
                target_lab = lab
                break
        if not target_lab:
            return 0, 0, 0, []

        assigned_count = len(target_lab.get("diffuser_ids", []))
        poly = target_lab["polygon"]
        # check diffuser ids from all labs (spatial containment)
        for lab in self.generated_space_labels:
            for did in lab.get("diffuser_ids", []):
                coords = self.canvas.coords(did)
                if not coords or len(coords) < 4:
                    continue
                cx = (coords[0] + coords[2]) / 2.0
                cy = (coords[1] + coords[3]) / 2.0
                try:
                    pt = Point(cx, cy)
                    if poly.contains(pt):
                        inside += 1
                    else:
                        outside += 1
                        outside_ids.append(did)
                except Exception:
                    outside += 1
                    outside_ids.append(did)

        return assigned_count, inside, outside, outside_ids

    def auto_place_diffusers(self, area_per_diffuser: float):
        """각 실의 면적 기준으로 디퓨저 개수 산정 및 배치"""
        if not self.generated_space_labels:
            messagebox.showinfo("정보", "자동생성된 실(공간) 라벨이 없습니다.\n먼저 '자동생성'을 수행하세요.")
            return

        self.push_history()

        # 안전하게 기존의 모든 diffuser 및 diffuser_label 아이템을 삭제
        try:
            for item in list(self.canvas.find_withtag("diffuser")):
                try:
                    self.canvas.delete(item)
                except Exception:
                    pass
        except Exception:
            pass
        try:
            for item in list(self.canvas.find_withtag("diffuser_label")):
                try:
                    self.canvas.delete(item)
                except Exception:
                    pass
        except Exception:
            pass

        # 각 lab 내부의 ID 리스트들도 초기화
        for lab in self.generated_space_labels:
            lab["diffuser_ids"] = []
            lab["diffuser_label_ids"] = []

        # 각 실별 디퓨저 생성
        for lab in self.generated_space_labels:
            # process every room (remove temporary Room 6-only filter)
            poly = lab["polygon"]
            area_text = self.canvas.itemcget(lab["area_id"], "text")
            area_val = 0.0
            for tok in area_text.replace("m²", "").split():
                try:
                    area_val = float(tok)
                    break
                except Exception:
                    continue
            
            if area_val <= 0:
                lab["diffuser_ids"] = []
                continue

            # Only place diffusers for central HVAC (1. 중앙공조)
            hvac = int(lab.get("hvac_type", 1)) if lab.get("hvac_type", None) is not None else 1
            if hvac != 1:
                # ensure no diffusers/labels remain for non-central HVAC rooms
                lab["diffuser_ids"] = []
                lab["diffuser_label_ids"] = []
                continue

            # 개수 결정 로직: 몫(int) -> 홀수면 +1 (짝수화)
            n = int(area_val // area_per_diffuser)
            if n < 1:
                n = 1
            if n % 2 == 1:
                n += 1
            
            # 위치 계산 및 그리기
            pts = self._generate_diffuser_points_for_poly(poly, n)
            diffuser_ids = []
            diffuser_label_ids = []
            radius = 3
            # cluster pts into rows by y coordinate
            spacing_px = max(1.0, self.meter_to_pixel(0.5))
            row_thresh = max(2.0, spacing_px * 0.5)
            pts_sorted = sorted(pts, key=lambda p: (p[1], p[0]))
            rows = []
            for (x, y) in pts_sorted:
                if not rows:
                    rows.append([(x, y)])
                    continue
                last_row = rows[-1]
                # compare to median y of last_row
                ys = [pt[1] for pt in last_row]
                med_y = sum(ys) / len(ys)
                if abs(y - med_y) <= row_thresh:
                    last_row.append((x, y))
                else:
                    rows.append([(x, y)])

            # sort each row by x (left to right)
            for i in range(len(rows)):
                rows[i].sort(key=lambda p: p[0])

            # assign Supply/Return alternating starting with (1,1)=Supply
            for ri, row in enumerate(rows):
                for ci, (x, y) in enumerate(row):
                    # parity: supply if (ri + ci) % 2 == 0
                    is_supply = ((ri + ci) % 2 == 0)
                    color = "green" if is_supply else "skyblue"
                    tag2 = "supply" if is_supply else "return"
                    did = self.canvas.create_oval(
                        x - radius, y - radius, x + radius, y + radius,
                        fill=color, outline="", tags=("diffuser", tag2)
                    )
                    # create label next to the diffuser
                    try:
                        tid = self.canvas.create_text(x + radius + 4, y,
                                                      text=("S" if is_supply else "R"),
                                                      anchor=tk.W,
                                                      fill=color,
                                                      font=("Arial", 8, "bold"),
                                                      tags=("diffuser_label", tag2))
                    except Exception:
                        tid = None
                    diffuser_ids.append(did)
                    if tid:
                        diffuser_label_ids.append(tid)

            lab["diffuser_ids"] = diffuser_ids
            lab["diffuser_label_ids"] = diffuser_label_ids

    # -------- 저장/불러오기용 직렬화 --------

    def to_dict(self):
        """현재 Palette 상태를 JSON 직렬화용 dict로 반환"""
        data = {
            "scale": self.scale,
            "shapes": [],
            "labels": []
        }
        for s in self.shapes:
            data["shapes"].append({
                "coords": list(s.coords),
                "editable": s.editable,
                "color": s.color
            })

        for lab in self.generated_space_labels:
            # 텍스트와 위치
            name_x, name_y = self.canvas.coords(lab["name_id"])
            norm_x, norm_y = self.canvas.coords(lab["heat_norm_id"])
            equip_x, equip_y = self.canvas.coords(lab["heat_equip_id"])
            area_x, area_y = self.canvas.coords(lab["area_id"])
            
            # 디퓨저 위치 저장
            diffuser_coords = []
            if "diffuser_ids" in lab:
                for did in lab["diffuser_ids"]:
                    coords = self.canvas.coords(did)
                    if coords:
                        cx = (coords[0] + coords[2]) / 2
                        cy = (coords[1] + coords[3]) / 2
                        diffuser_coords.append([cx, cy])

            data["labels"].append({
                "polygon_coords": list(lab["polygon"].exterior.coords),
                "name_text": self.canvas.itemcget(lab["name_id"], "text"),
                "heat_norm_text": self.canvas.itemcget(lab["heat_norm_id"], "text"),
                "heat_equip_text": self.canvas.itemcget(lab["heat_equip_id"], "text"),
                "area_text": self.canvas.itemcget(lab["area_id"], "text"),
                "name_pos": [name_x, name_y],
                "heat_norm_pos": [norm_x, norm_y],
                "heat_equip_pos": [equip_x, equip_y],
                "area_pos": [area_x, area_y],
                "diffuser_coords": diffuser_coords
                ,
                "hvac_type": int(lab.get("hvac_type", 1)),
                "hvac_detail": int(lab.get("hvac_detail", 1)) if lab.get("hvac_detail", None) is not None else 0,
                "hvac_text": lab.get("hvac_text", None),
                # persist edited quantity and detail text if present
                "hvac_qty": int(lab.get("hvac_qty")) if lab.get("hvac_qty", None) is not None else None,
                "hvac_detail_text": lab.get("hvac_detail_text", None)
            })
        # save grid visibility
        data["show_grid"] = bool(getattr(self, 'show_grid', False))
        return data

    def load_from_dict(self, data: dict):
        """JSON dict로부터 Palette 상태 복원"""
        self.canvas.delete("all")
        self.shapes.clear()
        self.generated_space_labels.clear()
        self.highlight_line_id = None
        self.tooltip_id = None
        self.corner_highlight_id = None

        self.scale = data.get("scale", 20.0)
        self.next_shape_id = 1

        # 도형 복원
        for info in data.get("shapes", []):
            coords = info.get("coords", [0, 0, 0, 0])
            editable = info.get("editable", True)
            color = info.get("color", "black")
            s = self.create_rect_shape(
                coords[0], coords[1], coords[2], coords[3],
                editable=editable, color=color, push_to_history=False
            )

        # 라벨 복원
        for lab in data.get("labels", []):
            poly = Polygon(lab["polygon_coords"])

            name_x, name_y = lab["name_pos"]
            norm_x, norm_y = lab["heat_norm_pos"]
            equip_x, equip_y = lab["heat_equip_pos"]
            area_x, area_y = lab["area_pos"]

            name_id = self.canvas.create_text(
                name_x, name_y,
                text=lab["name_text"], fill="blue", font=("Arial", 11, "bold"),
                tags=("space_name",)
            )
            heat_norm_id = self.canvas.create_text(
                norm_x, norm_y,
                text=lab["heat_norm_text"], fill="darkred", font=("Arial", 10),
                tags=("space_heat_norm",)
            )
            heat_equip_id = self.canvas.create_text(
                equip_x, equip_y,
                text=lab["heat_equip_text"], fill="darkred", font=("Arial", 10),
                tags=("space_heat_equip",)
            )
            area_id = self.canvas.create_text(
                area_x, area_y,
                text=lab["area_text"], fill="green", font=("Arial", 10)
            )
            
            # 디퓨저 복원
            diffuser_ids = []
            diffuser_coords = lab.get("diffuser_coords", [])
            r = 3
            for (cx, cy) in diffuser_coords:
                did = self.canvas.create_oval(
                    cx - r, cy - r, cx + r, cy + r,
                    fill="green", outline=""
                )
                diffuser_ids.append(did)

            # hvac_detail stored as 0 when missing; convert back to None
            stored_detail = lab.get("hvac_detail", 0)
            if stored_detail == 0:
                stored_detail = None
            self.generated_space_labels.append({
                "polygon": poly,
                "name_id": name_id,
                "heat_norm_id": heat_norm_id,
                "heat_equip_id": heat_equip_id,
                "area_id": area_id,
                "diffuser_ids": diffuser_ids,
                "hvac_type": int(lab.get("hvac_type", 1)),
                "hvac_detail": int(stored_detail) if stored_detail is not None else None,
                "hvac_text": lab.get("hvac_text", None),
                # restore persisted quantity and detail text if present
                "hvac_qty": int(lab.get("hvac_qty")) if lab.get("hvac_qty", None) is not None else None,
                "hvac_detail_text": lab.get("hvac_detail_text", None)
            })

        # 태그 바인딩 복원
        self.canvas.tag_bind("dim_width", "<Button-1>", self.on_dim_width_click)
        self.canvas.tag_bind("dim_height", "<Button-1>", self.on_dim_height_click)
        self.canvas.tag_bind("space_name", "<Button-1>", self.on_space_name_click)
        self.canvas.tag_bind("space_heat_norm", "<Button-1>", self.on_space_heat_norm_click)
        self.canvas.tag_bind("space_heat_equip", "<Button-1>", self.on_space_heat_equip_click)

        self.active_shape = None
        self.active_side_name = None
        self.app.update_selected_area_label(self)
        # restore grid visibility
        self.show_grid = bool(data.get("show_grid", False))
        if getattr(self, 'show_grid', False):
            try:
                self.draw_grid()
            except Exception:
                pass

    # -------- 줌 / 팬 --------

    def on_mouse_wheel(self, event):
        zoom_in = event.delta > 0
        self.apply_zoom(zoom_in, event.x, event.y)

    def on_mouse_wheel_linux(self, event):
        zoom_in = (event.num == 4)
        self.apply_zoom(zoom_in, event.x, event.y)

    def apply_zoom(self, zoom_in, cx, cy):
        factor = 1.1 if zoom_in else 1 / 1.1
        new_scale = self.scale * factor
        if new_scale < 2.0 or new_scale > 200.0:
            return

        self.push_history()

        self.canvas.scale("all", cx, cy, factor, factor)
        for shape in self.shapes:
            x1, y1, x2, y2 = shape.coords
            x1 = cx + (x1 - cx) * factor
            y1 = cy + (y1 - cy) * factor
            x2 = cx + (x2 - cx) * factor
            y2 = cy + (y2 - cy) * factor
            shape.coords = (x1, y1, x2, y2)

        self.scale = new_scale
        self.app.update_selected_area_label(self)
        # redraw grid to match new scale
        try:
            self.clear_grid()
            if getattr(self, 'show_grid', False):
                self.draw_grid()
        except Exception:
            pass

    def on_middle_button_down(self, event):
        self.push_history()
        self.panning = True
        self.pan_last_pos = (event.x, event.y)

    def on_middle_button_drag(self, event):
        if not self.panning or self.pan_last_pos is None:
            return
        last_x, last_y = self.pan_last_pos
        dx = event.x - last_x
        dy = event.y - last_y

        self.canvas.move("all", dx, dy)
        for shape in self.shapes:
            x1, y1, x2, y2 = shape.coords
            shape.coords = (x1 + dx, y1 + dy, x2 + dx, y2 + dy)
        self.pan_last_pos = (event.x, event.y)

    def on_middle_button_up(self, event):
        self.panning = False
        self.pan_last_pos = None
        # after panning, redraw grid for the new viewport
        try:
            if getattr(self, 'show_grid', False):
                # clear any old grid items and draw fresh
                self.clear_grid()
                self.draw_grid()
        except Exception:
            pass


# ================= 상위 App =================

# Backward compatibility: previous code used the name RectCanvas
RectCanvas = Palette

class ResizableRectApp:
    def __init__(self, root):
        self.root = root
        self.root.title("도형 편집기 (디퓨저 배치 기능 추가됨)")

        # left-side control panel (wider so controls are not clipped)
        left_panel = tk.Frame(self.root, width=280)
        left_panel.pack(side=tk.LEFT, fill=tk.Y)
        left_panel.pack_propagate(False)

        # Left-side notebook hosting Room Design as a tab
        self.left_notebook = ttk.Notebook(left_panel)
        self.left_notebook.pack(fill=tk.BOTH, expand=False, padx=6, pady=6)

        # Room Design tab/frame inside left_notebook
        tab_frame = tk.Frame(self.left_notebook)
        self.left_notebook.add(tab_frame, text="Room Design")

        # Duct Design tab with simple runtime rename controls
        self.duct_tab = tk.Frame(self.left_notebook)
        self.left_notebook.add(self.duct_tab, text="Duct")

        duct_rename_frame = tk.Frame(self.duct_tab)
        duct_rename_frame.pack(side=tk.TOP, fill=tk.X, pady=(4, 4), padx=6)
        tk.Label(duct_rename_frame, text="탭 이름:").pack(side=tk.LEFT)
        self.duct_tab_name_entry = tk.Entry(duct_rename_frame, width=14)
        self.duct_tab_name_entry.insert(0, "Duct")
        self.duct_tab_name_entry.pack(side=tk.LEFT, padx=(4, 6))
        # rename button with validation
        duct_rename_btn = tk.Button(duct_rename_frame, text="이름 변경", command=self._rename_duct_tab)
        duct_rename_btn.pack(side=tk.LEFT, padx=(2, 4))
        # restore default name button
        def _restore_default():
            self.duct_tab_name_entry.delete(0, tk.END)
            self.duct_tab_name_entry.insert(0, "Duct")
            try:
                idx = self.left_notebook.index(self.duct_tab)
                self.left_notebook.tab(idx, text="Duct")
            except Exception:
                pass

        restore_btn = tk.Button(duct_rename_frame, text="기본 복원", command=_restore_default)
        restore_btn.pack(side=tk.LEFT)

        # HVAC systems list for Duct tab
        hvac_frame = tk.Frame(self.duct_tab)
        hvac_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=False, padx=6, pady=(6, 4))
        tk.Label(hvac_frame, text="공조 시스템 목록:").pack(anchor='w')
        listbox_frame = tk.Frame(hvac_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=False)
        self.hvac_listbox = tk.Listbox(listbox_frame, height=6)
        self.hvac_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.hvac_scroll = tk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.hvac_listbox.yview)
        self.hvac_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.hvac_listbox.config(yscrollcommand=self.hvac_scroll.set)
        # right-click popup menu for deleting an HVAC item
        self._hvac_menu = tk.Menu(self.hvac_listbox, tearoff=0)
        self._hvac_menu.add_command(label="삭제", command=self._delete_selected_hvac)
        # platform-independent right-click binding
        self.hvac_listbox.bind("<Button-3>", self._on_hvac_right_click)
        # sample entries
        for item in ["AHU-1", "AHU-2", "FCU-1", "VAV-1"]:
            self.hvac_listbox.insert(tk.END, item)

        # controls: entry + add + apply
        hvac_ctrl = tk.Frame(hvac_frame)
        hvac_ctrl.pack(fill=tk.X, pady=(6, 2))
        tk.Label(hvac_ctrl, text="새 시스템:").pack(side=tk.LEFT)
        self.hvac_new_entry = tk.Entry(hvac_ctrl, width=16)
        self.hvac_new_entry.pack(side=tk.LEFT, padx=(4, 6))
        add_btn = tk.Button(hvac_ctrl, text="추가", command=lambda: self._add_hvac())
        add_btn.pack(side=tk.LEFT, padx=(0, 6))
        apply_btn = tk.Button(hvac_ctrl, text="적용", command=lambda: self._apply_hvac())
        apply_btn.pack(side=tk.LEFT)

        # area controls at top of Room Design
        area_ctrl = tk.Frame(tab_frame)
        area_ctrl.pack(side=tk.TOP, fill=tk.X, pady=(2, 4))
        tk.Label(area_ctrl, text="면적 (m²):").pack(side=tk.LEFT)
        self.area_entry = tk.Entry(area_ctrl, width=10)
        self.area_entry.pack(side=tk.LEFT, padx=5)
        draw_btn = tk.Button(area_ctrl, text="정사각형 그리기", command=self.draw_square_from_area_current)
        draw_btn.pack(side=tk.LEFT, padx=5)
        self.area_entry.bind("<Return>", lambda e: self.draw_square_from_area_current())

        # small top area in the room tab to host the Auto-generate button above the inputs
        top_ctrl = tk.Frame(tab_frame)
        top_ctrl.pack(side=tk.TOP, pady=(6, 4))
        ag_btn = tk.Button(top_ctrl, text="자동생성", width=18, command=self.auto_generate_current)
        ag_btn.pack()

        # control area inside Room Design
        control_frame = tk.Frame(tab_frame)
        control_frame.pack(fill=tk.BOTH, expand=False, padx=8, pady=4)

        # Temperature inputs
        tk.Label(control_frame, text="외기(°C):").grid(row=0, column=0, sticky="w")
        self.outdoor_temp_entry = tk.Entry(control_frame, width=8)
        self.outdoor_temp_entry.grid(row=0, column=1, padx=(6, 0), pady=2)
        self.outdoor_temp_entry.insert(0, "-5.0")

        tk.Label(control_frame, text="실내(°C):").grid(row=1, column=0, sticky="w")
        self.indoor_temp_entry = tk.Entry(control_frame, width=8)
        self.indoor_temp_entry.grid(row=1, column=1, padx=(6, 0), pady=2)
        self.indoor_temp_entry.insert(0, "25.0")

        tk.Label(control_frame, text="급기(°C):").grid(row=2, column=0, sticky="w")
        self.supply_temp_entry = tk.Entry(control_frame, width=8)
        self.supply_temp_entry.grid(row=2, column=1, padx=(6, 0), pady=2)
        self.supply_temp_entry.insert(0, "18.0")

        # Heat norm/equip with apply buttons and Enter bindings
        tk.Label(control_frame, text="일반 발열량\n(W/m²):").grid(row=3, column=0, sticky="w")
        self.heat_norm_entry = tk.Entry(control_frame, width=8)
        self.heat_norm_entry.grid(row=3, column=1, padx=(6, 0), pady=2)
        self.heat_norm_entry.insert(0, "0.00")
        norm_apply_btn = tk.Button(control_frame, text="적용", width=6, command=lambda: self._on_apply_norm())
        norm_apply_btn.grid(row=3, column=2, padx=(6, 0), pady=2)
        self.heat_norm_entry.bind("<Return>", lambda e: self._on_apply_norm())

        tk.Label(control_frame, text="장비 발열량\n(W/m²):").grid(row=4, column=0, sticky="w")
        self.heat_equip_entry = tk.Entry(control_frame, width=8)
        self.heat_equip_entry.grid(row=4, column=1, padx=(6, 0), pady=2)
        self.heat_equip_entry.insert(0, "0.00")
        equip_apply_btn = tk.Button(control_frame, text="적용", width=6, command=lambda: self._on_apply_equip())
        equip_apply_btn.grid(row=4, column=2, padx=(6, 0), pady=2)
        self.heat_equip_entry.bind("<Return>", lambda e: self._on_apply_equip())

        # 급기 풍량 산정 버튼
        supply_calc_btn = tk.Button(control_frame, text="급기 풍량 산정", width=12,
            command=lambda: self._on_calc_supply_flow())
        supply_calc_btn.grid(row=5, column=0, columnspan=3, pady=(8, 2))

        # 결과 표시용 텍스트 박스
        tk.Label(control_frame, text="총 급기 풍량 (m3/hr):").grid(row=6, column=0, columnspan=3, sticky="w", pady=(6, 0))
        self.supply_result_text = tk.Text(control_frame, height=4, width=24)
        self.supply_result_text.grid(row=7, column=0, columnspan=3, pady=(2, 0))

        # --- 디퓨저 관련 UI 추가 ---
        tk.Label(control_frame, text="디퓨저\n담당면적(m²):").grid(row=8, column=0, sticky="w", pady=(8, 2))
        self.diffuser_area_entry = tk.Entry(control_frame, width=8)
        self.diffuser_area_entry.grid(row=8, column=1, padx=(6, 0), pady=2)
        self.diffuser_area_entry.insert(0, "10.0")

        diffuser_btn = tk.Button(control_frame, text="디퓨저 자동 배치", width=14,
                 command=lambda: self._on_place_diffusers())
        diffuser_btn.grid(row=9, column=0, columnspan=2, pady=(4, 2))
        # Reset diffusers button next to auto-place
        reset_btn = tk.Button(control_frame, text="디퓨져 초기화", width=10,
                  command=lambda: self._on_reset_diffusers())
        reset_btn.grid(row=9, column=2, pady=(4, 2))

        # Diagnostic button: count diffusers inside/outside a room
        check_btn = tk.Button(control_frame, text="디퓨저 점검", width=14,
                  command=lambda: self._on_check_diffusers())
        check_btn.grid(row=10, column=0, columnspan=3, pady=(2, 6))

        # top_frame for remaining main toolbar buttons
        top_frame = tk.Frame(self.root)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

        self.area_label_var = tk.StringVar()
        self.area_label_var.set("선택 도형 면적: - m²")
        area_label = tk.Label(top_frame, textvariable=self.area_label_var, fg="blue")
        area_label.pack(side=tk.LEFT, padx=20)

        undo_btn = tk.Button(top_frame, text="되돌리기 (Ctrl+Z)", command=self.undo_current)
        undo_btn.pack(side=tk.LEFT, padx=10)

        add_tab_btn = tk.Button(top_frame, text="팔레트 추가", command=self.add_new_tab)
        add_tab_btn.pack(side=tk.LEFT, padx=10)

        delete_tab_btn = tk.Button(top_frame, text="팔레트 삭제", command=self.delete_current_tab)
        delete_tab_btn.pack(side=tk.LEFT, padx=5)

        clear_palette_btn = tk.Button(top_frame, text="팔레트 지우기", command=self.clear_current_palette)
        clear_palette_btn.pack(side=tk.LEFT, padx=5)

        save_btn = tk.Button(top_frame, text="저장하기", command=self.save_current)
        save_btn.pack(side=tk.LEFT, padx=5)

        load_btn = tk.Button(top_frame, text="불러오기", command=self.load_current)
        load_btn.pack(side=tk.LEFT, padx=5)
        # CSV preview/load button: opens a CSV and shows it in a new window as a table
        csv_btn = tk.Button(top_frame, text="CSV로드 (C)", command=self.load_csv_preview)
        csv_btn.pack(side=tk.LEFT, padx=5)
        equip_btn = tk.Button(top_frame, text="장비일람표 추출", command=self.extract_equipment_list)
        equip_btn.pack(side=tk.LEFT, padx=5)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.palettes = []
        self.add_new_tab()

        self.root.bind_all("<Control-z>", lambda e: self.undo_current())

    # ---------- 버튼 콜백 ----------
    def _on_apply_norm(self):
        try:
            v = float(self.heat_norm_entry.get())
        except Exception:
            messagebox.showerror("입력 오류", "일반 발열량에 숫자를 입력하세요.")
            return
        rc = self.get_current_palette()
        if rc:
            rc.apply_norm_to_all(v)

    def _on_apply_equip(self):
        try:
            v = float(self.heat_equip_entry.get())
        except Exception:
            messagebox.showerror("입력 오류", "장비 발열량에 숫자를 입력하세요.")
            return
        rc = self.get_current_palette()
        if rc:
            rc.apply_equip_to_all(v)
        
    def _on_calc_supply_flow(self):
        rc = self.get_current_palette()
        if not rc:
            messagebox.showinfo("정보", "활성화된 팔레트가 없습니다.")
            return

        total = rc.compute_and_apply_supply_flow()

        try:
            self.supply_result_text.delete("1.0", tk.END)
            self.supply_result_text.insert(tk.END, f"Total supply flow: {total:.1f} m3/hr")
        except Exception:
            messagebox.showinfo("결과", f"총 급기 풍량: {total:.1f} m3/hr")
            
    def _on_place_diffusers(self):
        rc = self.get_current_palette()
        if not rc:
            messagebox.showinfo("정보", "활성화된 팔레트가 없습니다.")
            return
        try:
            a = float(self.diffuser_area_entry.get())
            if a <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("입력 오류", "디퓨저 담당면적에 양의 숫자를 입력하세요.")
            return
        rc.auto_place_diffusers(a)

    def _rename_duct_tab(self):
        """Rename the Duct tab at runtime using the entry value."""
        try:
            new_name = (self.duct_tab_name_entry.get() or "").strip()
            if not new_name:
                messagebox.showinfo("입력 필요", "새 탭 이름을 입력하세요.")
                return
            # enforce a reasonable max length
            max_len = 20
            if len(new_name) > max_len:
                new_name = new_name[:max_len]
                messagebox.showinfo("이름 잘림", f"탭 이름은 {max_len}자 이하로 잘립니다.")
            # find index of the duct tab and set its text
            idx = None
            try:
                idx = self.left_notebook.index(self.duct_tab)
            except Exception:
                # fallback: search by widget
                for i in range(self.left_notebook.index("end")):
                    if self.left_notebook.nametowidget(self.left_notebook.tabs()[i]) is self.duct_tab:
                        idx = i
                        break
            if idx is not None:
                self.left_notebook.tab(idx, text=new_name)
                # update entry to normalized name (in case trimming occurred)
                try:
                    self.duct_tab_name_entry.delete(0, tk.END)
                    self.duct_tab_name_entry.insert(0, new_name)
                except Exception:
                    pass
        except Exception as e:
            messagebox.showerror("오류", f"탭 이름 변경 중 오류: {e}")

    def _add_hvac(self):
        """Add new HVAC system from the entry into the listbox."""
        try:
            name = (self.hvac_new_entry.get() or "").strip()
            if not name:
                messagebox.showinfo("입력 필요", "추가할 시스템 이름을 입력하세요.")
                return
            # prevent duplicates
            existing = self.hvac_listbox.get(0, tk.END)
            if name in existing:
                messagebox.showinfo("중복", "이미 존재하는 시스템입니다.")
                return
            self.hvac_listbox.insert(tk.END, name)
            self.hvac_new_entry.delete(0, tk.END)
        except Exception as e:
            messagebox.showerror("오류", f"시스템 추가 중 오류: {e}")

    def _apply_hvac(self):
        """Apply the selected HVAC system (placeholder action)."""
        try:
            sel = self.hvac_listbox.curselection()
            if not sel:
                messagebox.showinfo("선택 필요", "적용할 시스템을 선택하세요.")
                return
            name = self.hvac_listbox.get(sel[0])
            # placeholder: show a message; integrate with actual logic as needed
            messagebox.showinfo("적용", f"선택된 시스템: {name}")
        except Exception as e:
            messagebox.showerror("오류", f"적용 중 오류: {e}")

    def _on_hvac_right_click(self, event):
        """Show a small popup menu to delete the item under the cursor."""
        try:
            # index of item under pointer
            idx = self.hvac_listbox.nearest(event.y)
            if idx is None:
                return
            # select the item so user sees which will be affected
            self.hvac_listbox.selection_clear(0, tk.END)
            self.hvac_listbox.selection_set(idx)
            # popup the menu at the pointer location
            try:
                self._hvac_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self._hvac_menu.grab_release()
        except Exception as e:
            # silent fail with optional debug
            print(f"HVAC right-click menu error: {e}")

    def _delete_selected_hvac(self):
        """Delete the currently selected HVAC item after user confirmation."""
        try:
            sel = self.hvac_listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            name = self.hvac_listbox.get(idx)
            if messagebox.askyesno("삭제 확인", f"'{name}' 항목을 삭제하시겠습니까?"):
                self.hvac_listbox.delete(idx)
        except Exception as e:
            messagebox.showerror("오류", f"삭제 중 오류: {e}")

    def _on_reset_diffusers(self):
        rc = self.get_current_palette()
        if not rc:
            messagebox.showinfo("정보", "활성화된 팔레트가 없습니다.")
            return
    
    def load_csv_preview(self):
        """Open a CSV file and show its contents in a new window as a simple table.

        Uses several encoding fallbacks (utf-8, cp949, latin-1) to handle common CSV encodings.
        """
        from tkinter import filedialog
        import csv

        file_path = filedialog.askopenfilename(title="CSV 파일 선택", filetypes=[("CSV files", "*.csv"), ("All files", "*")])
        if not file_path:
            return

        encodings = ["utf-8", "cp949", "latin-1"]
        rows = []
        used_enc = None
        for enc in encodings:
            try:
                with open(file_path, 'r', encoding=enc, errors='strict') as f:
                    reader = csv.reader(f)
                    rows = [r for r in reader]
                used_enc = enc
                break
            except Exception:
                # try next encoding
                continue

        if used_enc is None:
            # last resort: open with latin-1 permissive
            try:
                with open(file_path, 'r', encoding='latin-1', errors='replace') as f:
                    reader = csv.reader(f)
                    rows = [r for r in reader]
                used_enc = 'latin-1'
            except Exception as e:
                messagebox.showerror("CSV 로드 오류", f"파일을 읽을 수 없습니다:\n{e}")
                return

        # create preview window
        # save loaded CSV on the app for later use by popups
        try:
            self.last_csv_rows = rows
            self.last_csv_file_path = file_path
            self.last_csv_encoding = used_enc
        except Exception:
            pass
        
        win = tk.Toplevel(self.root)
        win.title(f"CSV 미리보기 - {os.path.basename(file_path)} ({used_enc})")
        win.geometry("800x400")

        from tkinter import ttk
        tree = ttk.Treeview(win, show='headings')

        # determine column count
        max_cols = max((len(r) for r in rows), default=0)
        cols = [f"C{i+1}" for i in range(max_cols)]
        tree['columns'] = cols
        for i, c in enumerate(cols):
            tree.heading(c, text=c)
            tree.column(c, width=120, anchor='w')

        # insert rows
        for r in rows:
            # pad shorter rows
            row = list(r) + [""] * (max_cols - len(r))
            tree.insert('', tk.END, values=row)

        tree.pack(fill=tk.BOTH, expand=True)

        # add a simple close button
        btn = tk.Button(win, text="닫기", command=win.destroy)
        btn.pack(side=tk.BOTTOM, pady=6)
 
    def extract_equipment_list(self):
        """Collect per-room displayed values and export to .xlsx (or CSV fallback).

        Uses the currently selected palette (tab). If CSV rows were loaded earlier via
        'CSV로드', those rows are used to select table values per-room similarly to the
        popup logic.
        """
        rows = getattr(self, 'last_csv_rows', None)
        if not rows:
            messagebox.showinfo("장비일람표 없음", "먼저 CSV를 로드하세요 (툴바의 'CSV로드').")
            return

        rc = self.get_current_palette()
        if not rc:
            messagebox.showinfo('장비일람표 없음', '활성화된 팔레트를 선택하세요.')
            return

        base_header = [
            "Room Name",
            "HVAC Type",
            "HVAC Detail",
            "Area (m2)",
            "Norm (W/m2)",
            "Equip (W/m2)",
            "Matched CSV Column",
            "Matched Value",
        ]

        export_rows = [base_header]

        def _get_text(lab, key):
            try:
                return rc.canvas.itemcget(lab[key], 'text')
            except Exception:
                return ""

        # helper to extract first numeric token
        def _extract_num_token(s: str):
            if not s:
                return ""
            for tok in s.replace(',', ' ').split():
                try:
                    float(tok)
                    return tok
                except Exception:
                    continue
            return ""

        # iterate labels from the active palette
        for lab in getattr(rc, 'generated_space_labels', []):
            try:
                name = _get_text(lab, 'name_id')
                area_text = _get_text(lab, 'area_id')
                norm_text = _get_text(lab, 'heat_norm_id')
                equip_text = _get_text(lab, 'heat_equip_id')

                area_val = _extract_num_token(area_text)
                norm_val = _extract_num_token(norm_text)
                equip_val = _extract_num_token(equip_text)

                hvac_type = lab.get('hvac_type', '')
                hvac_detail_text = lab.get('hvac_detail_text', '') or lab.get('hvac_detail', '') or ''

                matched_col = ''
                matched_val = ''
                table_rows_for_export = []

                # reproduce popup selection logic minimally
                try:
                    rows_local = rows
                    if rows_local and len(rows_local) >= 2 and str(hvac_type) == '2' and hvac_detail_text:
                        headers = rows_local[0]
                        second = rows_local[1]
                        sel_detail = hvac_detail_text

                        preferred_cols = []
                        for ci, h in enumerate(headers):
                            try:
                                if sel_detail.lower() in str(h).lower():
                                    preferred_cols.append(ci)
                            except Exception:
                                continue

                        candidates = []
                        if preferred_cols:
                            for ci in preferred_cols:
                                try:
                                    v = float(second[ci]) if ci < len(second) else None
                                    if v is not None:
                                        candidates.append((ci, v))
                                except Exception:
                                    continue
                        else:
                            for ci, val in enumerate(second):
                                try:
                                    v = float(val)
                                    candidates.append((ci, v))
                                except Exception:
                                    continue

                        csv_shown_vals = []
                        if sel_detail and not preferred_cols:
                            csv_shown_vals = [((r[0] if len(r) > 0 else ''), '') for r in rows_local]
                        else:
                            if candidates:
                                import math
                                # compute total_kw safely
                                try:
                                    total_kw_local = float(area_val) * (float(norm_val) if norm_val else 0.0 + float(equip_val) if equip_val else 0.0) / 1000.0
                                except Exception:
                                    total_kw_local = 0.0

                                if preferred_cols and sel_detail:
                                    vals_only = [c[1] for c in candidates]
                                    max_val = max(vals_only) if vals_only else 0.0

                                    # If the user previously saved a quantity for this lab, prefer it
                                    stored_qty = None
                                    try:
                                        if 'hvac_qty' in lab and lab.get('hvac_qty') is not None:
                                            stored_qty = int(lab.get('hvac_qty'))
                                    except Exception:
                                        stored_qty = None

                                    # compute qty/target using stored quantity when available
                                    if stored_qty is not None and stored_qty > 0:
                                        qty = stored_qty
                                        target = total_kw_local / max(1, qty)
                                    else:
                                        if max_val > 0 and total_kw_local < max_val:
                                            target = total_kw_local / 2.0
                                            qty = 2
                                        else:
                                            if max_val > 0:
                                                ratio = total_kw_local / max_val
                                            else:
                                                ratio = total_kw_local
                                            qty = int(math.ceil(ratio)) + 2
                                            if qty <= 0:
                                                qty = 2
                                            target = total_kw_local / qty if qty != 0 else total_kw_local

                                    greater = [c for c in candidates if c[1] >= target]
                                    if greater:
                                        best = min(greater, key=lambda x: x[1])
                                    else:
                                        lesser = [c for c in candidates if c[1] < target]
                                        if lesser:
                                            best = max(lesser, key=lambda x: x[1])
                                        else:
                                            best = candidates[0]
                                    ci, cv = best
                                    matched_col = headers[ci] if ci < len(headers) else f'C{ci+1}'
                                    csv_shown_vals = [((r[0] if len(r) > 0 else ''), (r[ci] if ci < len(r) else '')) for r in rows_local]
                                    try:
                                        # ensure displayed quantity reflects stored value if present
                                        csv_shown_vals.insert(0, ("대수(Q'ty)", str(qty)))
                                    except Exception:
                                        pass
                                    for data_row in rows_local[1:]:
                                        if ci < len(data_row) and data_row[ci] is not None and str(data_row[ci]).strip() != "":
                                            matched_val = data_row[ci]
                                            break
                                else:
                                            # consider stored quantity when available for default selection as well
                                            stored_qty = None
                                            try:
                                                if 'hvac_qty' in lab and lab.get('hvac_qty') is not None:
                                                    stored_qty = int(lab.get('hvac_qty'))
                                            except Exception:
                                                stored_qty = None

                                            if stored_qty is not None and stored_qty > 0:
                                                # compute target per-unit value and pick candidate closest >= target else closest below
                                                target = total_kw_local / max(1, stored_qty)
                                                greater = [c for c in candidates if c[1] >= target]
                                                if greater:
                                                    best = min(greater, key=lambda x: x[1])
                                                else:
                                                    lesser = [c for c in candidates if c[1] < target]
                                                    if lesser:
                                                        best = max(lesser, key=lambda x: abs(x[1] - target))
                                                    else:
                                                        best = candidates[0]
                                            else:
                                                greater = [c for c in candidates if c[1] >= total_kw_local]
                                                if greater:
                                                    best = min(greater, key=lambda x: x[1])
                                                else:
                                                    lesser = [c for c in candidates if c[1] < total_kw_local]
                                                    if lesser:
                                                        best = max(lesser, key=lambda x: x[1])
                                                    else:
                                                        best = candidates[0]
                                            ci, cv = best
                                            matched_col = headers[ci] if ci < len(headers) else f'C{ci+1}'
                                            csv_shown_vals = [((r[0] if len(r) > 0 else ''), (r[ci] if ci < len(r) else '')) for r in rows_local]
                                            # if stored_qty exists, insert it as the first quantity row so exported table reflects edited qty
                                            try:
                                                if stored_qty is not None:
                                                    csv_shown_vals.insert(0, ("대수(Q'ty)", str(stored_qty)))
                                                else:
                                                    csv_shown_vals.insert(0, ("대수(Q'ty)", "1"))
                                            except Exception:
                                                pass
                                            for data_row in rows_local[1:]:
                                                if ci < len(data_row) and data_row[ci] is not None and str(data_row[ci]).strip() != "":
                                                    matched_val = data_row[ci]
                                                    break

                        # prepare table rows list
                        for t0, v0 in csv_shown_vals:
                            table_rows_for_export.append((str(t0).strip(), str(v0).strip()))
                except Exception:
                    matched_col = ''
                    matched_val = ''
                    table_rows_for_export = []

                export_rows.append([
                    name,
                    str(hvac_type),
                    hvac_detail_text,
                    area_val,
                    norm_val,
                    equip_val,
                    matched_col,
                    matched_val,
                    table_rows_for_export
                ])
            except Exception:
                # skip problematic label but continue
                continue

        # Build final header with Table columns expanded to maximum seen
        max_table_pairs = 0
        for r in export_rows[1:]:
            trows = r[8] if len(r) > 8 else []
            if trows and isinstance(trows, list):
                max_table_pairs = max(max_table_pairs, len(trows))

        # If we have only the base header and no room rows, show diagnostics and stop
        if len(export_rows) <= 1:
            try:
                cnt = len(getattr(rc, 'generated_space_labels', []))
                rows_diag = []
                for lab in getattr(rc, 'generated_space_labels', []):
                    try:
                        keys = list(lab.keys())
                    except Exception:
                        keys = []
                    try:
                        n = rc.canvas.itemcget(lab.get('name_id', -1), 'text')
                    except Exception:
                        n = '<err>'
                    try:
                        a = rc.canvas.itemcget(lab.get('area_id', -1), 'text')
                    except Exception:
                        a = '<err>'
                    try:
                        hv = lab.get('hvac_type', '<no>')
                    except Exception:
                        hv = '<err>'
                    try:
                        hd = lab.get('hvac_detail_text', lab.get('hvac_detail', ''))
                    except Exception:
                        hd = '<err>'
                    rows_diag.append(f"keys={keys[:10]} name={n!s} area={a!s} hvac={hv!s} detail={hd!s}")

                diag_text = f"내보낼 실별 데이터가 없습니다. 생성된 라벨 수: {cnt}\n\n" + "\n".join(rows_diag[:20])
                if len(diag_text) > 2500:
                    diag_text = diag_text[:2500] + "\n..."
                messagebox.showinfo('데이터 없음 - 진단', diag_text)
            except Exception:
                messagebox.showinfo('데이터 없음', '내보낼 실별 데이터가 없습니다. 캔버스의 라벨을 확인하세요.')
            return

        final_header = base_header[:]
        for i in range(1, max_table_pairs + 1):
            final_header.append(f"Table_{i}_Title")
            final_header.append(f"Table_{i}_Value")

        # Ask user where to save .xlsx
        from tkinter import filedialog
        fp = filedialog.asksaveasfilename(parent=self.root, defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx'), ('All files', '*.*')], title='장비일람표 저장 (.xlsx)')
        if not fp:
            return

        # Try to save as .xlsx; fallback to CSV if openpyxl not available
        try:
            try:
                import openpyxl
                from openpyxl import Workbook
                openpyxl_path = getattr(openpyxl, '__file__', None)
            except Exception:
                # fallback to CSV
                import csv, sys
                csv_fp = fp
                if csv_fp.lower().endswith('.xlsx'):
                    csv_fp = csv_fp[:-5] + '.csv'
                # write CSV with BOM so Excel (Windows) recognizes UTF-8 Korean text
                with open(csv_fp, 'w', newline='', encoding='utf-8-sig') as cf:
                    writer = csv.writer(cf)
                    writer.writerow(final_header)
                    for r in export_rows[1:]:
                        base = r[:8]
                        tlist = r[8] if len(r) > 8 else []
                        row_out = list(base)
                        for (t0, v0) in tlist:
                            row_out.append(t0)
                            row_out.append(v0)
                        while len(row_out) < len(final_header):
                            row_out.append("")
                        writer.writerow(row_out)
                messagebox.showinfo('저장 완료 (CSV)', f"openpyxl이 없어 CSV로 저장했습니다: {csv_fp}\nPython: {sys.executable}")
                return

            wb = Workbook()
            ws = wb.active
            ws.title = '장비일람표'
            ws.append(final_header)
            for r in export_rows[1:]:
                base = r[:8]
                tlist = r[8] if len(r) > 8 else []
                row_out = list(base)
                for (t0, v0) in tlist:
                    row_out.append(t0)
                    row_out.append(v0)
                while len(row_out) < len(final_header):
                    row_out.append("")
                ws.append(row_out)
            wb.save(fp)
            try:
                openpyxl_info = openpyxl_path
            except Exception:
                openpyxl_info = None
            messagebox.showinfo('저장 완료', f'장비일람표를 저장했습니다: {fp}\nPython: {sys.executable}\nopenpyxl: {openpyxl_info}')
        except Exception as e:
            messagebox.showerror('저장 오류', f'파일 저장 중 오류가 발생했습니다:\n{e}')
 
    def get_current_palette(self) -> Palette | None:
        if not self.notebook.tabs():
            return None
        idx = self.notebook.index(self.notebook.select())
        if 0 <= idx < len(self.palettes):
            return self.palettes[idx]
        return None

    def add_new_tab(self):
        tab = tk.Frame(self.notebook)
        self.notebook.add(tab, text=f"팔레트 {len(self.palettes)+1}")
        self.notebook.select(len(self.palettes))
        rc = Palette(tab, app=self)
        self.palettes.append(rc)

    def delete_current_tab(self):
        if not self.palettes:
            return

        current_index = self.notebook.index(self.notebook.select())
        if len(self.palettes) == 1:
            messagebox.showinfo(
                "삭제 불가",
                "마지막 팔레트는 삭제할 수 없습니다.\n새 팔레트를 추가한 후 삭제해 주세요."
            )
            return

        answer = messagebox.askyesno(
            "팔레트 삭제 확인",
            "현재 선택된 팔레트를 정말로 삭제하시겠습니까?\n"
            "이 작업은 되돌릴 수 없습니다."
        )
        if not answer:
            return

        rc_to_delete = self.palettes[current_index]
        rc_to_delete.shapes.clear()
        rc_to_delete.generated_space_labels.clear()
        rc_to_delete.canvas.delete("all")
        rc_to_delete.highlight_line_id = None
        rc_to_delete.tooltip_id = None
        rc_to_delete.corner_highlight_id = None
        rc_to_delete.active_shape = None
        rc_to_delete.active_side_name = None

        tabs = self.notebook.tabs()
        if 0 <= current_index < len(tabs):
            tab_id = tabs[current_index]
            self.notebook.forget(tab_id)
        if 0 <= current_index < len(self.palettes):
            del self.palettes[current_index]

    def clear_current_palette(self):
        rc = self.get_current_palette()
        if not rc:
            return

        answer = messagebox.askyesno("팔레트 초기화 확인", "현재 팔레트의 모든 내용을 삭제하고 처음부터 다시 그리시겠습니까?")
        if not answer:
            return

        rc.shapes.clear()
        # delete all items except those tagged as 'grid' so the grid remains visible
        try:
            all_items = list(rc.canvas.find_all())
            for item in all_items:
                try:
                    tags = rc.canvas.gettags(item)
                    if 'grid' in tags:
                        continue
                    rc.canvas.delete(item)
                except Exception:
                    continue
        except Exception:
            try:
                rc.canvas.delete("all")
            except Exception:
                pass
        rc.generated_space_labels.clear()
        rc.highlight_line_id = None
        rc.tooltip_id = None
        rc.corner_highlight_id = None
        rc.active_shape = None
        rc.active_side_name = None

        rc.canvas.tag_bind("dim_width", "<Button-1>", rc.on_dim_width_click)
        rc.canvas.tag_bind("dim_height", "<Button-1>", rc.on_dim_height_click)
        rc.canvas.tag_bind("space_name", "<Button-1>", rc.on_space_name_click)
        rc.canvas.tag_bind("space_heat_norm", "<Button-1>", rc.on_space_heat_norm_click)
        rc.canvas.tag_bind("space_heat_equip", "<Button-1>", rc.on_space_heat_equip_click)

        self.update_selected_area_label(rc)

    def draw_square_from_area_current(self):
        rc = self.get_current_palette()
        if not rc:
            return
        s = self.area_entry.get().strip()
        try:
            area = float(s)
        except ValueError:
            return
        rc.draw_square_from_area(area)

    def undo_current(self):
        rc = self.get_current_palette()
        if rc:
            rc.undo()

    def auto_generate_current(self):
        rc = self.get_current_palette()
        if rc:
            rc.auto_generate_space_labels()

    def save_current(self):
        rc = self.get_current_palette()
        if not rc:
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON 파일", "*.json"), ("모든 파일", "*.*")]
        )
        if not file_path:
            return
        data = rc.to_dict()
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("저장 완료", f"현재 팔레트를 저장했습니다.\n{file_path}")
        except Exception as e:
            messagebox.showerror("저장 오류", f"파일 저장 중 오류가 발생했습니다.\n{e}")

    def load_current(self):
        rc = self.get_current_palette()
        if not rc:
            return
        file_path = filedialog.askopenfilename(
            defaultextension=".json",
            filetypes=[("JSON 파일", "*.json"), ("모든 파일", "*.*")]
        )
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            rc.load_from_dict(data)
            messagebox.showinfo("불러오기 완료", f"팔레트를 불러왔습니다.\n{file_path}")
        except Exception as e:
            messagebox.showerror("불러오기 오류", f"파일 불러오기 중 오류가 발생했습니다.\n{e}")

    def update_selected_area_label(self, rc: Palette | None):
        if not rc or not rc.active_shape:
            self.area_label_var.set("선택 도형 면적: - m²")
            return
        x1, y1, x2, y2 = rc.active_shape.coords
        w = rc.pixel_to_meter(x2 - x1)
        h = rc.pixel_to_meter(y2 - y1)
        self.area_label_var.set(f"선택 도형 면적: {w*h:.3f} m²")

    def _on_check_diffusers(self):
        rc = self.get_current_palette()
        if not rc:
            messagebox.showinfo("정보", "활성화된 팔레트가 없습니다.")
            return
        room_name = simpledialog.askstring("룸 선택", "검사할 룸 이름을 입력하세요:", initialvalue="Room 6")
        if not room_name:
            return
        assigned_count, inside_count, outside_count, outside_ids = rc.check_diffusers_in_room(room_name)
        msg = (f"{room_name}: 할당된 디퓨저 {assigned_count}개\n"
               f"공간 기준 내부 {inside_count}개, 외부 {outside_count}개")
        messagebox.showinfo("디퓨저 점검 결과", msg)
        # highlight outside points in red briefly
        for did in outside_ids:
            try:
                rc.canvas.itemconfig(did, fill="red")
            except Exception:
                pass


if __name__ == "__main__":
    root = tk.Tk()
    app = ResizableRectApp(root)
    root.mainloop()
