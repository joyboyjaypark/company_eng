import sys
import math
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QComboBox, QMessageBox
)
from PyQt5.QtCore import Qt

class DuctSizingTool(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("덕트 사이징 Tool (최대 풍속 & 정압 체크)")
        self.setMinimumWidth(480)

        # ----- 입력부 -----
        # 1) 풍량
        flow_label = QLabel("풍량 (m³/h):")
        self.flow_input = QLineEdit()
        self.flow_input.setPlaceholderText("예: 2000")

        # 2) 최대 허용 풍속
        vel_label = QLabel("최대 허용 풍속 (m/s):")
        self.vel_input = QLineEdit()
        self.vel_input.setPlaceholderText("예: 5")

        # 3) 최대 허용 정압 (Pa)
        #    실제 설계에서는 전체 시스템 정압, 팬 정압 한계 등을 의미
        pa_label = QLabel("최대 허용 정압 (Pa):")
        self.pa_input = QLineEdit()
        self.pa_input.setPlaceholderText("예: 800")

        # 4) 덕트 형식
        type_label = QLabel("덕트 형식:")
        self.type_combo = QComboBox()
        self.type_combo.addItems(["원형", "직사각형"])
        self.type_combo.currentIndexChanged.connect(self.on_type_changed)

        # 5) 직사각형 가로:세로 비 (W:H)
        ratio_label = QLabel("가로:세로 비 (W:H):")
        self.ratio_input = QLineEdit()
        self.ratio_input.setPlaceholderText("예: 2:1")
        self.ratio_input.setEnabled(False)  # 기본: 원형이므로 비활성화

        # 계산 버튼
        self.calc_button = QPushButton("계산하기")
        self.calc_button.clicked.connect(self.calculate)

        # ----- 결과부 -----
        self.area_label = QLabel("단면적 A: - m²")
        self.size_label = QLabel("덕트 크기: -")
        self.velocity_label = QLabel("실제 속도: - m/s")
        self.pressure_label = QLabel("정압 체크: -")

        # 굵은 글씨
        bold_font = self.area_label.font()
        bold_font.setBold(True)
        self.area_label.setFont(bold_font)
        self.size_label.setFont(bold_font)
        self.velocity_label.setFont(bold_font)
        self.pressure_label.setFont(bold_font)

        # ----- 레이아웃 구성 -----
        main_layout = QVBoxLayout()

        # 풍량
        flow_layout = QHBoxLayout()
        flow_layout.addWidget(flow_label)
        flow_layout.addWidget(self.flow_input)
        main_layout.addLayout(flow_layout)

        # 최대 풍속
        vel_layout = QHBoxLayout()
        vel_layout.addWidget(vel_label)
        vel_layout.addWidget(self.vel_input)
        main_layout.addLayout(vel_layout)

        # 최대 정압
        pa_layout = QHBoxLayout()
        pa_layout.addWidget(pa_label)
        pa_layout.addWidget(self.pa_input)
        main_layout.addLayout(pa_layout)

        # 덕트 형식
        type_layout = QHBoxLayout()
        type_layout.addWidget(type_label)
        type_layout.addWidget(self.type_combo)
        main_layout.addLayout(type_layout)

        # 가로:세로 비
        ratio_layout = QHBoxLayout()
        ratio_layout.addWidget(ratio_label)
        ratio_layout.addWidget(self.ratio_input)
        main_layout.addLayout(ratio_layout)

        # 버튼
        main_layout.addWidget(self.calc_button, alignment=Qt.AlignRight)

        # 결과 구역
        main_layout.addWidget(QLabel("결과"))
        main_layout.addWidget(self.area_label)
        main_layout.addWidget(self.size_label)
        main_layout.addWidget(self.velocity_label)
        main_layout.addWidget(self.pressure_label)

        self.setLayout(main_layout)

    def on_type_changed(self, index):
        # 0: 원형, 1: 직사각형
        is_rectangular = (index == 1)
        self.ratio_input.setEnabled(is_rectangular)

    def parse_ratio(self, text):
        """
        '2:1' 형식 문자열을 받아서 float 비율(2.0) 반환
        """
        try:
            parts = text.split(':')
            if len(parts) != 2:
                return None
            w = float(parts[0])
            h = float(parts[1])
            if h == 0:
                return None
            return w / h
        except ValueError:
            return None

    def calculate(self):
        # 1) 입력값 파싱
        try:
            Q_h = float(self.flow_input.text())      # m³/h
            V_max = float(self.vel_input.text())     # m/s (최대 허용 풍속)
            P_max = float(self.pa_input.text())      # Pa  (최대 허용 정압)
            if Q_h <= 0 or V_max <= 0 or P_max <= 0:
                raise ValueError
        except ValueError:
            QMessageBox.warning(
                self, "입력 오류",
                "풍량, 최대 허용 풍속, 최대 허용 정압을 올바른 양수로 입력하세요."
            )
            return

        duct_type = self.type_combo.currentText()
        ratio = None
        if duct_type == "직사각형":
            ratio = self.parse_ratio(self.ratio_input.text())
            if ratio is None or ratio <= 0:
                QMessageBox.warning(
                    self, "입력 오류",
                    "가로:세로 비를 '2:1' 형식의 올바른 숫자로 입력하세요."
                )
                return

        # 2) 풍량 변환 (m³/h → m³/s)
        Q_s = Q_h / 3600.0  # m³/s

        # 3) 단면적 계산
        #    A = Q / V_max  → 최대 풍속을 넘지 않도록 하는 최소 단면적
        A = Q_s / V_max  # m²

        # 4) 형식별 치수 계산
        if duct_type == "원형":
            D = math.sqrt(4.0 * A / math.pi)  # m
            D_mm = D * 1000.0
            size_text = f"원형 직경 D ≈ {D_mm:,.0f} mm"
        else:
            # 직사각형: A = W * H,  W = ratio * H
            H = math.sqrt(A / ratio)   # m
            W = ratio * H              # m
            H_mm = H * 1000.0
            W_mm = W * 1000.0
            size_text = f"직사각형 W×H ≈ {W_mm:,.0f} mm × {H_mm:,.0f} mm"

        # 5) 실제 속도(이론상 V_max와 같음, 검산용)
        V_real = Q_s / A

        # 6) 정압 체크 (단순 예시)
        #    실제 정압은 덕트 길이, 마찰계수, 국부손실 등을 모두 합해야 하지만,
        #    여기서는 사용자가 입력한 P_max만 기준으로,
        #    "이 덕트에서 예상 정압이 P_max를 넘지 않는지"라는 개념적 체크만 표시.
        #
        #    보다 현실적인 계산을 위해서는:
        #    - 덕트 길이 L (m)
        #    - 단위 길이당 손실 dp/dx (Pa/m)
        #    - 엘보, 댐퍼 등의 국부 손실
        #    를 합산해서 P_total을 구한 후 P_total <= P_max 여부를 확인해야 함.
        #
        #    여기서는 아직 dp/dx, L을 받지 않았으므로,
        #    "설정된 최대 정압 P_max 이내에서 설계해야 함"이라는 메시지만 제공.
        pressure_text = (
            f"※ 주의: 실제 정압 계산(덕트 길이, Pa/m, 국부손실 등)을 통해 "
            f"총 정압이 {P_max:.0f} Pa를 넘지 않는지 별도 검토 필요."
        )

        # 7) 결과 표시
        self.area_label.setText(f"단면적 A ≈ {A:.4f} m²")
        self.size_label.setText(f"덕트 크기: {size_text}")
        self.velocity_label.setText(f"실제 속도 ≈ {V_real:.2f} m/s")
        self.pressure_label.setText(pressure_text)

def main():
    app = QApplication(sys.argv)
    window = DuctSizingTool()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
