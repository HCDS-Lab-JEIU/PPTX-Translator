from PySide6.QtCore import Signal
from PySide6.QtWidgets import QWidget,QVBoxLayout,QHBoxLayout,QLabel,QProgressBar,QPushButton


class ProgressPanel(QWidget):
    pause_requested=Signal(); stop_requested=Signal()
    def __init__(self):
        super().__init__(); self.setObjectName('progressPanel')
        box=QVBoxLayout(self); box.setContentsMargins(20,12,20,12)
        row=QHBoxLayout(); self.title=QLabel('전체 진행 · 대기 중'); self.count=QLabel('0 / 0 파일')
        row.addWidget(self.title); row.addStretch(); row.addWidget(self.count); box.addLayout(row)
        self.bar=QProgressBar(); self.bar.setFixedHeight(9); self.bar.setTextVisible(False); box.addWidget(self.bar)
        row=QHBoxLayout(); self.current=QLabel('현재 파일 · —    현재 단계 · —'); self.current.setWordWrap(True)
        row.addWidget(self.current,1)
        self.pause=QPushButton('일시정지'); self.stop=QPushButton('중지')
        self.pause.clicked.connect(self.pause_requested); self.stop.clicked.connect(self.stop_requested)
        row.addWidget(self.pause); row.addWidget(self.stop); box.addLayout(row)
        self.set_running(False)
    def set_running(self,running,paused=False):
        self.pause.setEnabled(running); self.stop.setEnabled(running)
        self.pause.setText('재개' if paused else '일시정지')
