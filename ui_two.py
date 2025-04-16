import os

# import shutil
import sys

from PySide6.QtCore import Qt, QUrl, Slot, QTimer
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QMessageBox,
    QHBoxLayout,
    QLineEdit,
)
from qt_material import apply_stylesheet

from utils import download, get_school_names
from draw import draw_two


class MainWindow(QWidget):
    def __init__(self, target_dir):
        super().__init__()
        self.target_dir = target_dir

        self.draw_count = 0

        self.school1_name = ""
        self.school2_name = ""
        self.school1_count = 15
        self.school2_count = 9

        self.run_id = 0

        self.renderer()

    def renderer(self):
        self.resize(800, 600)
        self.setWindowTitle("台州市人人测抽签程序")

        title = QLabel("台州市人人测抽取程序（2025年，两所学校）")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 24pt;")

        info1 = QLabel("科目：语数英")
        info1.setAlignment(Qt.AlignCenter)
        info1.setStyleSheet("font-size: 20pt")

        btn1 = QPushButton("上传学校的语数英成绩 EXCEL")
        btn1.clicked.connect(self.handle_get_schools)
        btn1.setFixedWidth = 100

        self.counts = QWidget()
        self.counts.setVisible(False)
        counts_hbox = QHBoxLayout(self.counts)

        self.school1_name_label = QLabel()
        self.school1_name_label.setStyleSheet("font-size: 20pt")
        self.school1_count_edit = QLineEdit()
        self.school1_count_edit.setStyleSheet("font-size: 20pt")
        self.school1_count_edit.setText("15")
        self.school1_count_edit.textChanged.connect(self.handle_school1_count)

        self.school2_name_label = QLabel()
        self.school2_name_label.setStyleSheet("font-size: 20pt")
        self.school2_count_edit = QLineEdit()
        self.school2_count_edit.setStyleSheet("font-size: 20pt")
        self.school2_count_edit.setText("9")
        self.school2_count_edit.textChanged.connect(self.handle_school2_count)

        counts_hbox.addStretch(1)
        counts_hbox.addWidget(self.school1_name_label)
        counts_hbox.addWidget(self.school1_count_edit)
        counts_hbox.addWidget(self.school2_name_label)
        counts_hbox.addWidget(self.school2_count_edit)
        counts_hbox.addStretch(1)

        self.btn2 = QPushButton("抽取")
        self.btn2.clicked.connect(self.handle_draw_yushuying)
        self.btn2.setFixedWidth = 100
        self.btn2.setVisible(False)

        self.message = QLabel("")
        self.message.setAlignment(Qt.AlignCenter)
        self.message.setStyleSheet("font-size: 24pt;")

        self.layout = QVBoxLayout(self)

        self.layout.addWidget(title)
        self.layout.addWidget(info1)
        self.layout.addWidget(btn1)

        self.layout.addWidget(self.counts)

        self.layout.addWidget(self.btn2)

        self.layout.addWidget(self.message)
        self.layout.addStretch(1)

    @Slot()
    def handle_school1_count(self):
        self.school1_count = int(self.school1_count_edit.text())
        self.school2_count = 24 - self.school1_count
        self.school2_count_edit.setText(str(self.school2_count))

    @Slot()
    def handle_school2_count(self):
        self.school2_count = int(self.school2_count_edit.text())
        self.school1_count = 24 - self.school2_count
        self.school1_count_edit.setText(str(self.school1_count))

    @Slot()
    def handle_get_schools(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "选择文件", "", "Excel Files (*.xlsx *.xls)"
        )
        if not file:
            return
        self.file = file

        school1, school2 = get_school_names(self.file)
        self.school1_name_label.setText(school1)
        self.school2_name_label.setText(school2)
        self.school1_name = school1
        self.school2_name = school2
        self.counts.setVisible(True)
        self.btn2.setVisible(True)

    @Slot()
    def handle_draw_yushuying(self):
        if not self.file:
            return
        self.run_id += 1
        self.message.setText(f"尝试第 {self.run_id} 次抽取...")
        try:
            df = draw_two(
                self.file,
                (self.school1_name, self.school1_count),
                (self.school2_name, self.school2_count),
            )
            download(df, self.target_dir, self.file)
            self.message.setText("抽取成功！")
            self.run_id = 0
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.target_dir))
        except ValueError:
            QTimer.singleShot(100, self.handle_draw_yushuying)


if __name__ == "__main__":
    app = QApplication([])

    cwd = os.getcwd()
    target_dir = os.path.join(cwd, "已抽取")
    # if os.path.exists(target_dir):
    #     reply = QMessageBox.warning(
    #         None,
    #         "友情提醒",
    #         "检测到当前目录已经存在【已抽取】文件夹，是否删除它？\n\n选【Yes】自动删除\n选【No】请手动改名后再打开软件",
    #         QMessageBox.Yes | QMessageBox.No,
    #     )
    #     if reply == QMessageBox.Yes:
    #         shutil.rmtree(target_dir)
    #     else:
    #         exit()

    if not os.path.exists(target_dir):
        os.mkdir(target_dir)

    window = MainWindow(target_dir)
    apply_stylesheet(app, theme="dark_teal.xml")
    window.show()
    sys.exit(app.exec())
