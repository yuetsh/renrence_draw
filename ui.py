import os
import shutil
import sys

from PySide6.QtCore import Qt, QUrl, Slot
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)
from qt_material import apply_stylesheet

from download import download
from draw import draw


class MainWindow(QWidget):
    def __init__(self, target_dir):
        super().__init__()
        self.target_dir = target_dir
        self.renderer()

    def renderer(self):
        self.resize(800, 600)
        self.setWindowTitle("台州市人人测抽签程序")

        title = QLabel("台州市人人测抽取程序")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 24pt;")

        self.message = QLabel("")
        self.message.setAlignment(Qt.AlignCenter)
        self.message.setStyleSheet("font-size: 24pt;")

        self.upload_btn = QPushButton("上传excel文件")
        self.upload_btn.clicked.connect(self.handler)

        self.layout = QVBoxLayout(self)
        self.layout.addWidget(title)
        self.layout.addWidget(self.upload_btn)
        self.layout.addWidget(self.message)

    @Slot()
    def handler(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择文件", "", "Excel Files (*.xlsx *.xls)"
        )
        if not len(files):
            return
        self.message.setText("抽取中...")
        for filepath in files:
            df = draw(filepath)
            download(df, self.target_dir, filepath)
        self.message.setText("抽取成功！")
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.target_dir))


if __name__ == "__main__":
    cwd = os.getcwd()
    target_dir = os.path.join(cwd, "已抽取")
    if os.path.exists(target_dir):
        shutil.rmtree(target_dir)
    os.mkdir(target_dir)

    app = QApplication([])
    window = MainWindow(target_dir)
    apply_stylesheet(app, theme="dark_teal.xml")
    window.show()
    sys.exit(app.exec())
