from datetime import datetime
from PySide6.QtCore import Signal
from PySide6.QtWidgets import QFrame,QVBoxLayout,QHBoxLayout,QLabel,QPushButton,QProgressBar


class QuotaPanel(QFrame):
    refresh_requested=Signal()
    def __init__(self):
        super().__init__(); self.setObjectName('quotaPanel')
        box=QVBoxLayout(self); box.setContentsMargins(12,12,12,12)
        row=QHBoxLayout(); row.addWidget(QLabel('남은 한도')); row.addStretch()
        self.refresh=QPushButton('새로고침'); self.refresh.clicked.connect(self.refresh_requested); row.addWidget(self.refresh); box.addLayout(row)
        self.content=QVBoxLayout(); box.addLayout(self.content)
    def render(self,state):
        while self.content.count():
            item=self.content.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        title=QLabel(state.quota_status + (' · 모의 데이터' if state.demo else '')); title.setWordWrap(True); self.content.addWidget(title)
        for bucket in state.buckets:
            value=f'{bucket.remaining:g}% 남음' if bucket.remaining is not None else '잔량 미제공'
            label=QLabel(f'{bucket.name}\n{value}'); label.setWordWrap(True); self.content.addWidget(label)
            if bucket.remaining is not None:
                meter=QProgressBar(); meter.setRange(0,100); meter.setValue(round(bucket.remaining)); meter.setTextVisible(False); meter.setFixedHeight(7); self.content.addWidget(meter)
            if bucket.resets_at is not None:
                label=QLabel('초기화 · '+datetime.fromtimestamp(bucket.resets_at).astimezone().strftime('%m/%d %H:%M %Z')); label.setObjectName('muted'); self.content.addWidget(label)
        note=state.quota_error or '같은 연결을 선택한 노드와 공유 · 계정 한도'
        if state.updated_at: note+='\n갱신 · '+state.updated_at.astimezone().strftime('%H:%M:%S')
        label=QLabel(note); label.setObjectName('muted'); label.setWordWrap(True); self.content.addWidget(label)
