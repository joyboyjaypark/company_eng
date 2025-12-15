import sys
import math
import time

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QPen, QColor
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QGraphicsView, QGraphicsScene,
    QGraphicsEllipseItem, QGraphicsTextItem, QMenu, QMessageBox,
    QProgressBar
)


class PointItem(QGraphicsEllipseItem):
    """
    개별 점(원)을 표현하는 그래픽 아이템.
    인덱스, 시작/종료 여부 등을 저장해 둔다.
    """
    def __init__(self, x, y, radius=5):
        super().__init__(-radius, -radius, 2 * radius, 2 * radius)
        self.setPos(x, y)
        self.radius = radius
        self.is_start = False
        self.is_end = False
        self.index = -1  # 외부에서 설정
        self.text_item: QGraphicsTextItem | None = None

        self.setBrush(QColor("black"))
        self.setPen(Qt.NoPen)
        self.setFlag(QGraphicsEllipseItem.ItemIsSelectable, True)

    def update_appearance(self):
        """
        시작점/종료점 여부에 따라 색과 라벨 텍스트 변경
        """
        if self.is_start:
            color = QColor("green")
            label = f"S({self.index})"
        elif self.is_end:
            color = QColor("red")
            label = f"E({self.index})"
        else:
            color = QColor("black")
            label = str(self.index)

        self.setBrush(color)
        if self.text_item is not None:
            self.text_item.setDefaultTextColor(color)
            self.text_item.setPlainText(label)


class GraphicsView(QGraphicsView):
    """
    사용자 입력(좌클릭으로 점 추가, 우클릭 컨텍스트 메뉴)을 처리하는 뷰.
    """
    def __init__(self, scene: QGraphicsScene, app_ref: "MainWindow"):
        super().__init__(scene)
        self.app_ref = app_ref  # 메인 윈도우 객체 참조

        # 우클릭 메뉴
        self.context_menu = QMenu(self)
        self.action_set_start = QAction("시작점으로 지정", self)
        self.action_set_end = QAction("종료점으로 지정", self)
        self.action_delete = QAction("이 점 삭제", self)

        self.context_menu.addAction(self.action_set_start)
        self.context_menu.addAction(self.action_set_end)
        self.context_menu.addSeparator()
        self.context_menu.addAction(self.action_delete)

        self.clicked_item: PointItem | None = None

        self.action_set_start.triggered.connect(self.set_start_point)
        self.action_set_end.triggered.connect(self.set_end_point)
        self.action_delete.triggered.connect(self.delete_point)

    def mousePressEvent(self, event):
        # Qt6 스타일: position()/globalPosition() 사용
        if event.button() == Qt.LeftButton:
            scene_pos = self.mapToScene(event.position().toPoint())
            self.app_ref.add_point(scene_pos.x(), scene_pos.y())
        elif event.button() == Qt.RightButton:
            scene_pos = self.mapToScene(event.position().toPoint())
            items = self.scene().items(scene_pos)
            point_item = None
            for it in items:
                if isinstance(it, PointItem):
                    point_item = it
                    break
            if point_item is not None:
                self.clicked_item = point_item
                self.context_menu.popup(event.globalPosition().toPoint())
        else:
            super().mousePressEvent(event)

    def set_start_point(self):
        if self.clicked_item is not None:
            self.app_ref.set_start_point(self.clicked_item)

    def set_end_point(self):
        if self.clicked_item is not None:
            self.app_ref.set_end_point(self.clicked_item)

    def delete_point(self):
        if self.clicked_item is not None:
            self.app_ref.delete_point(self.clicked_item)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PySide6 90도 경로 최적화 (휴리스틱: 최근접 + 2-opt)")

        # 그래픽 장면/뷰
        self.scene = QGraphicsScene(self)
        self.scene.setSceneRect(0, 0, 800, 600)

        self.view = GraphicsView(self.scene, self)
        self.view.setRenderHints(self.view.renderHints())

        # 상단 컨트롤 바
        top_widget = QWidget()
        top_layout = QHBoxLayout(top_widget)

        self.btn_optimize = QPushButton("최적화")
        self.btn_clear = QPushButton("초기화")
        self.label_info = QLabel("왼쪽 클릭: 점 추가 / 오른쪽 클릭: 시작·종료 지정")
        self.label_distance = QLabel("총 길이: -")

        # 진행률 바 (2-opt 반복 정도를 시각화하는 용도)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedWidth(150)

        top_layout.addWidget(self.btn_optimize)
        top_layout.addWidget(self.btn_clear)
        top_layout.addWidget(self.label_info)
        top_layout.addStretch()
        top_layout.addWidget(QLabel("개선 진행:"))
        top_layout.addWidget(self.progress_bar)
        top_layout.addWidget(self.label_distance)

        self.btn_optimize.clicked.connect(self.optimize_path)
        self.btn_clear.clicked.connect(self.clear_all)

        # 중앙 레이아웃
        central_widget = QWidget()
        central_layout = QVBoxLayout(central_widget)
        central_layout.addWidget(top_widget)
        central_layout.addWidget(self.view)

        self.setCentralWidget(central_widget)

        # 데이터 구조: PointItem 리스트
        self.points: list[PointItem] = []
        self.start_index: int | None = None
        self.end_index: int | None = None

        # 그려진 경로(직선) 아이템들
        self.path_lines: list = []

    # --------------------------
    # 유틸 함수
    # --------------------------
    def manhattan(self, a: PointItem, b: PointItem) -> float:
        # 이 줄은 정상입니다. QGraphicsItem.pos() 사용 (경고 아님)
        return abs(a.pos().x() - b.pos().x()) + abs(a.pos().y() - b.pos().y())

    def path_length(self, order: list[int]) -> float:
        total = 0.0
        for i in range(len(order) - 1):
            total += self.manhattan(self.points[order[i]], self.points[order[i + 1]])
        return total

    def add_point(self, x: float, y: float):
        """
        새로운 점 추가 (왼쪽 클릭)
        """
        item = PointItem(x, y, radius=5)
        item.index = len(self.points)

        # 라벨 텍스트
        text_item = QGraphicsTextItem(str(item.index))
        text_item.setDefaultTextColor(QColor("black"))
        text_item.setPos(x + 10, y - 15)
        item.text_item = text_item

        self.scene.addItem(item)
        self.scene.addItem(text_item)

        self.points.append(item)

        # 경로는 무효화
        self.clear_path()
        self.label_distance.setText("총 길이: -")
        self.progress_bar.setValue(0)

    def redraw_indices(self):
        """
        점 인덱스를 다시 매기고, 라벨/색을 갱신
        (삭제 후 호출)
        """
        for i, pt in enumerate(self.points):
            pt.index = i
            pt.update_appearance()
            if pt.text_item is not None:
                pt.text_item.setPos(pt.pos().x() + 10, pt.pos().y() - 15)

    def set_start_point(self, point_item: PointItem):
        """
        우클릭 메뉴에서 시작점 지정
        """
        # 기존 시작점 해제
        if self.start_index is not None and 0 <= self.start_index < len(self.points):
            self.points[self.start_index].is_start = False

        self.start_index = point_item.index
        point_item.is_start = True

        # 시작점이 종료점과 같으면 종료점 해제
        if self.end_index == self.start_index:
            point_item.is_end = False
            self.end_index = None

        # 외형 갱신
        for pt in self.points:
            pt.update_appearance()

        self.clear_path()
        self.label_distance.setText("총 길이: -")
        self.progress_bar.setValue(0)

    def set_end_point(self, point_item: PointItem):
        """
        우클릭 메뉴에서 종료점 지정
        """
        # 기존 종료점 해제
        if self.end_index is not None and 0 <= self.end_index < len(self.points):
            self.points[self.end_index].is_end = False

        self.end_index = point_item.index
        point_item.is_end = True

        # 종료점이 시작점과 같으면 시작점 해제
        if self.start_index == self.end_index:
            point_item.is_start = False
            self.start_index = None

        # 외형 갱신
        for pt in self.points:
            pt.update_appearance()

        self.clear_path()
        self.label_distance.setText("총 길이: -")
        self.progress_bar.setValue(0)

    def delete_point(self, point_item: PointItem):
        """
        우클릭 메뉴에서 점 삭제
        """
        idx = point_item.index

        # 씬에서 제거
        self.scene.removeItem(point_item)
        if point_item.text_item is not None:
            self.scene.removeItem(point_item.text_item)

        # 리스트에서 제거
        if 0 <= idx < len(self.points):
            del self.points[idx]

        # 시작/종료 인덱스 조정
        if self.start_index is not None:
            if self.start_index == idx:
                self.start_index = None
            elif self.start_index > idx:
                self.start_index -= 1

        if self.end_index is not None:
            if self.end_index == idx:
                self.end_index = None
            elif self.end_index > idx:
                self.end_index -= 1

        self.redraw_indices()
        self.clear_path()
        self.label_distance.setText("총 길이: -")
        self.progress_bar.setValue(0)

    def clear_path(self):
        """
        그려진 경로(라인)만 삭제
        """
        for line in self.path_lines:
            self.scene.removeItem(line)
        self.path_lines = []

    def clear_all(self):
        """
        모든 점/경로 삭제 및 상태 초기화
        """
        self.clear_path()
        for pt in self.points:
            self.scene.removeItem(pt)
            if pt.text_item is not None:
                self.scene.removeItem(pt.text_item)
        self.points = []
        self.start_index = None
        self.end_index = None
        self.label_distance.setText("총 길이: -")
        self.progress_bar.setValue(0)

    # --------------------------
    # 휴리스틱: 최근접 이웃 + 2-opt
    # --------------------------
    def build_initial_path_nearest_neighbor(self) -> list[int]:
        """
        시작점에서 출발해,
        아직 방문하지 않은 점들 중 맨해튼 거리가 가장 가까운 점을
        차례로 방문하는 초기 경로를 만든다.
        마지막에 종료점을 붙인다.
        """
        n = len(self.points)
        if n < 2 or self.start_index is None or self.end_index is None:
            return []

        all_indices = list(range(n))
        middle = [i for i in all_indices if i not in (self.start_index, self.end_index)]

        path = [self.start_index]
        current = self.start_index

        unvisited = set(middle)

        while unvisited:
            cur_pt = self.points[current]
            # 가장 가까운 점 선택
            next_idx = min(
                unvisited,
                key=lambda j: self.manhattan(cur_pt, self.points[j])
            )
            path.append(next_idx)
            unvisited.remove(next_idx)
            current = next_idx

        path.append(self.end_index)
        return path

    def two_opt(self, path: list[int], max_iter: int = 200) -> list[int]:
        """
        2-opt 개선 알고리즘.
        path를 조금씩 뒤집어가며 경로 길이를 줄이는 지역 탐색.
        """
        if len(path) <= 3:
            return path

        best = path[:]
        best_len = self.path_length(best)
        n = len(best)

        # 2-opt 반복 횟수를 진행률에 반영
        for it in range(max_iter):
            improved = False
            for i in range(1, n - 2):
                for j in range(i + 1, n - 1):
                    # 간선 (i-1,i) 와 (j,j+1)를 끊고, 중간 구간을 뒤집어 연결
                    a, b = best[i - 1], best[i]
                    c, d = best[j], best[j + 1]

                    before = (self.manhattan(self.points[a], self.points[b]) +
                              self.manhattan(self.points[c], self.points[d]))
                    after = (self.manhattan(self.points[a], self.points[c]) +
                             self.manhattan(self.points[b], self.points[d]))

                    if after + 1e-9 < before:
                        new_path = best[:i] + best[i:j + 1][::-1] + best[j + 1:]
                        best = new_path
                        best_len = self.path_length(best)
                        improved = True
                        break
                if improved:
                    break

            # 진행률 (단순히 반복 비율)
            progress = int((it + 1) / max_iter * 100)
            self.progress_bar.setValue(progress)
            QApplication.processEvents()

            if not improved:
                break

        return best

    # --------------------------
    # 최적화 버튼: 휴리스틱 실행
    # --------------------------
    def optimize_path(self):
        if len(self.points) < 2:
            QMessageBox.warning(self, "경고", "최소 2개 이상의 점이 필요합니다.")
            return
        if self.start_index is None or self.end_index is None:
            QMessageBox.warning(self, "경고", "시작점과 종료점을 모두 지정하세요.")
            return
        if self.start_index == self.end_index:
            QMessageBox.warning(self, "경고", "시작점과 종료점을 서로 다르게 지정하세요.")
            return

        # 초기 경로 생성 (최근접 이웃)
        initial_path = self.build_initial_path_nearest_neighbor()
        if not initial_path:
            QMessageBox.critical(self, "오류", "경로를 생성하지 못했습니다.")
            return

        # 2-opt 개선
        self.progress_bar.setValue(0)
        QApplication.processEvents()
        start_time = time.time()

        improved_path = self.two_opt(initial_path, max_iter=200)

        elapsed = time.time() - start_time

        # 경로 그리기
        self.clear_path()
        self.draw_orthogonal_path(improved_path)

        total_len = self.path_length(improved_path)
        self.label_distance.setText(f"총 길이: {total_len:.0f} (휴리스틱, {elapsed:.2f}초)")
        self.progress_bar.setValue(100)

    # --------------------------
    # 90도(수평/수직) 경로 그리기
    # --------------------------
    def draw_orthogonal_path(self, order: list[int]):
        """
        order: 점 인덱스 순서 (예: [start, ..., end])
        각 구간을 두 번의 직선(수평, 수직)으로 연결:
        (x1, y1) -> (x2, y1) -> (x2, y2)
        """
        pen = QPen(QColor("blue"))
        pen.setWidth(2)

        for i in range(len(order) - 1):
            a = self.points[order[i]]
            b = self.points[order[i + 1]]

            x1, y1 = a.pos().x(), a.pos().y()
            x2, y2 = b.pos().x(), b.pos().y()

            # 먼저 x 맞추고, 그 다음 y 맞추는 형태 (직각 경로)
            line1 = self.scene.addLine(x1, y1, x2, y1, pen)
            line2 = self.scene.addLine(x2, y1, x2, y2, pen)

            self.path_lines.append(line1)
            self.path_lines.append(line2)


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.resize(900, 700)
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
