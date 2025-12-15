# -*- coding: utf-8 -*-
"""
tkinter + tkinterweb + folium 예제
- '지도에서 클릭으로 거리 계산' 버튼 클릭
- Tk 창 안에 Folium 지도 표시
- 지도에서 두 지점을 '우클릭'하면 두 점 사이 직선거리(km) 팝업
"""

import os
import math
import traceback
import tkinter as tk
from tkinter import ttk, messagebox

import folium
from tkinterweb import HtmlFrame   # pip install tkinterweb


# ---------------------------
# 거리 계산 (하버사인)
# ---------------------------

def haversine_distance(lat1, lon1, lat2, lon2):
    R = 6371.0
    from math import radians, sin, cos, atan2, sqrt
    phi1, phi2 = radians(lat1), radians(lat2)
    d_phi = radians(lat2 - lat1)
    d_lambda = radians(lon2 - lon1)

    a = sin(d_phi / 2) ** 2 + cos(phi1) * cos(phi2) * sin(d_lambda / 2) ** 2
    c = 2 * atan2(sqrt(a), sqrt(1 - a))
    return R * c


# ---------------------------
# Folium 클릭형 지도 생성
# ---------------------------

def create_clickable_map(output_file: str = "map_distance_click.html") -> str:
    """
    클릭으로 두 지점을 선택해서 직선거리를 표시하는 Folium 지도 생성.
    외부 JS 파일 사용 안 함 → HTML 안에 직접 <script> 삽입.
    """
    center = [36.0, 127.5]
    m = folium.Map(location=center, zoom_start=7)
    map_name = m.get_name()  # ex) "map_123456..."

    # JS 전체를 하나의 문자열로 작성 (외부 js 없음)
    js_code = f"""
    (function() {{
        var map = window['{map_name}'];
        if (!map) {{
            console.error('Leaflet map object not found');
            return;
        }}

        var points = [];
        var layers = [];

        function toRad(x) {{ return x * Math.PI / 180; }}
        function haversine(lat1, lon1, lat2, lon2) {{
            var R = 6371.0;
            var dLat = toRad(lat2 - lat1);
            var dLon = toRad(lon2 - lon1);
            var a = Math.sin(dLat/2) * Math.sin(dLat/2) +
                    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
                    Math.sin(dLon/2) * Math.sin(dLon/2);
            var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
            return R * c;
        }}

        function clearLayers() {{
            for (var i = 0; i < layers.length; i++) {{
                try {{ map.removeLayer(layers[i]); }} catch(e) {{}}
            }}
            layers = [];
        }}

        // 지도 영역에서 기본 우클릭 메뉴 방지
        map.getContainer().addEventListener('contextmenu', function(ev) {{
            ev.preventDefault();
        }});

        map.on('contextmenu', function(e) {{
            if (points.length >= 2) {{
                points = [];
                clearLayers();
            }}

            var lat = e.latlng.lat;
            var lng = e.latlng.lng;
            points.push([lat, lng]);

            var color = (points.length === 1) ? 'red' : 'blue';
            var marker = L.circleMarker([lat, lng], {{
                radius: 6,
                color: color,
                fill: true,
                fillColor: color
            }}).addTo(map).bindPopup('Point ' + points.length);
            layers.push(marker);

            if (points.length === 2) {{
                var p0 = points[0];
                var p1 = points[1];

                var line = L.polyline([p0, p1], {{
                    color: 'green',
                    weight: 4,
                    opacity: 0.8
                }}).addTo(map);
                layers.push(line);

                var d = haversine(p0[0], p0[1], p1[0], p1[1]);
                var midLat = (p0[0] + p1[0]) / 2.0;
                var midLng = (p0[1] + p1[1]) / 2.0;

                var html = '<div style="font-size:14px;font-weight:bold;' +
                           'background:rgba(255,255,255,0.95);padding:6px;' +
                           'border-radius:4px;">' +
                           '직선거리: ' + d.toFixed(3) + ' km' +
                           '</div>';

                var popup = L.popup({{ maxWidth: 300 }})
                    .setLatLng([midLat, midLng])
                    .setContent(html)
                    .openOn(map);
                layers.push(popup);
            }}
        }});

        console.log('Contextmenu distance tool attached.');
    }})();
    """

    # Folium HTML 안에 스크립트 직접 삽입
    m.get_root().html.add_child(folium.Element(f"<script>{js_code}</script>"))
    m.save(output_file)
    return output_file


# ---------------------------
# Tkinter GUI
# ---------------------------

class DistanceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("대한민국 주소 간 직선거리 계산기 (GUI)")
        self.geometry("900x700")
        self.resizable(True, True)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        self._init_ui()

    def _init_ui(self):
        title = ttk.Label(
            self,
            text="대한민국 두 주소 사이의 직선거리 계산 & 지도 생성",
            font=("맑은 고딕", 12, "bold"),
        )
        title.pack(pady=8)

        info = ttk.Label(
            self,
            text="지도에서 두 지점을 '우클릭'하여 직선거리를 계산합니다.",
            font=("맑은 고딕", 10),
            foreground="gray",
        )
        info.pack(pady=4)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=6)

        btn = ttk.Button(
            btn_frame,
            text="지도에서 클릭으로 거리 계산",
            command=self.on_open_click_map,
        )
        btn.grid(row=0, column=0, padx=6)

        btn_quit = ttk.Button(btn_frame, text="종료", command=self.destroy)
        btn_quit.grid(row=0, column=1, padx=6)

        self.lbl_result = ttk.Label(
            self,
            text="결과: 아직 계산되지 않았습니다.",
            font=("맑은 고딕", 11),
        )
        self.lbl_result.pack(pady=8)

        self.map_frame = ttk.Frame(self)
        self.map_frame.pack(fill="both", expand=True, padx=6, pady=6)

        self.status_var = tk.StringVar(value="준비 완료")
        status = ttk.Label(
            self,
            textvariable=self.status_var,
            font=("맑은 고딕", 9),
            relief="sunken",
            anchor="w",
        )
        status.pack(side="bottom", fill="x")

    def on_open_click_map(self):
        try:
            self.status_var.set("지도를 생성하는 중입니다...")
            self.update_idletasks()

            html_file = create_clickable_map("map_distance_click.html")

            # 이전 내용 제거 후 HtmlFrame에 로드
            for c in self.map_frame.winfo_children():
                c.destroy()

            hf = HtmlFrame(self.map_frame, horizontal_scrollbar="auto")
            hf.pack(fill="both", expand=True)

            # 절대 경로 또는 file:// 경로 사용
            abs_path = os.path.abspath(html_file)
            hf.load_website(abs_path)

            self.status_var.set(
                "완료! 창 안의 지도에서 두 지점을 '우클릭'하면 직선거리를 볼 수 있습니다."
            )
            self.lbl_result.config(text=f"지도 파일: {abs_path}")

        except Exception:
            self.status_var.set("오류 발생")
            with open("map_distance_error.log", "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
            messagebox.showerror(
                "오류",
                "지도를 생성/표시하는 중 오류가 발생했습니다.\n"
                "map_distance_error.log 파일을 확인해 주세요.",
            )


def main():
    app = DistanceApp()
    app.mainloop()


if __name__ == "__main__":
    main()
