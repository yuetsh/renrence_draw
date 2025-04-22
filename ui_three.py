import os
import sys

import pandas as pd

from PySide6.QtCore import Qt, QUrl, Slot
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QHBoxLayout,
    QLineEdit,
)
from qt_material import apply_stylesheet

from utils import download
from draw import draw_one, draw_zhuanyeke


class MainWindow(QWidget):
    def __init__(self, target_dir):
        super().__init__()
        self.target_dir = target_dir
        self.draw_count = 0
        self.renderer()

    def renderer(self):
        self.resize(800, 600)
        self.setWindowTitle("台州市人人测抽签程序")

        title = QLabel("台州市人人测抽取程序（2025年，专业课）")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 24pt;")

        self.zhuanyeke1 = QPushButton("第一步：上传【语数英】已经抽到的学生名单 EXECL")
        self.zhuanyeke1.clicked.connect(self.upload_yushuying)
        zhuanyeke2 = QPushButton("第二步：上传各校专业课成绩 EXCEL")
        zhuanyeke2.clicked.connect(self.handle_draw_zhuanyeke)

        self.message = QLabel("")
        self.message.setAlignment(Qt.AlignCenter)
        self.message.setStyleSheet("font-size: 24pt;")

        self.layout = QVBoxLayout(self)

        self.layout.addWidget(title)
        self.layout.addWidget(self.zhuanyeke1)
        self.layout.addWidget(zhuanyeke2)

        self.layout.addWidget(self.message)
        self.layout.addStretch(1)

    @Slot()
    def upload_yushuying(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "选择文件", "", "Excel Files (*.xlsx *.xls)"
        )
        if not file:
            return
        self.yushuying_users = pd.read_excel(file, sheet_name=0)
        self.zhuanyeke1.setDisabled(True)

    @Slot()
    def handle_draw_zhuanyeke(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择文件", "", "Excel Files (*.xlsx *.xls)"
        )
        if not len(files):
            return
        self.message.setText("抽取中...")
        for filepath in files:
            df = draw_zhuanyeke(self.yushuying_users, filepath)
            download(df, self.target_dir, filepath)
        self.message.setText("抽取成功！")
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.target_dir))


if __name__ == "__main__":
    app = QApplication([])

    cwd = os.getcwd()
    target_dir = os.path.join(cwd, "已抽取")

    if not os.path.exists(target_dir):
        os.mkdir(target_dir)

    window = MainWindow(target_dir)
    apply_stylesheet(app, theme="dark_teal.xml")
    window.show()
    sys.exit(app.exec())
