from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QWidget, QFrame, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QPlainTextEdit, QSplitter,QScrollArea
from ..models import NODES, STATES, PROVIDERS


class NodeCard(QPushButton):
    def __init__(self,key):
        super().__init__()
        self.key=key; self.setObjectName('nodeCard'); self.setMinimumHeight(82); self.setMinimumWidth(330)
        self.setCheckable(True)
        box=QVBoxLayout(self); box.setContentsMargins(17,12,17,12)
        row=QHBoxLayout()
        title=QLabel(NODES[key][0]); title.setObjectName('nodeTitle')
        self.state=QLabel(STATES['ready']); self.state.setObjectName('nodeState')
        for child in (title,self.state): child.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        row.addWidget(title); row.addStretch(); row.addWidget(self.state); box.addLayout(row)
        self.summary=QLabel(NODES[key][1]); self.summary.setObjectName('muted'); self.summary.setWordWrap(True)
        self.summary.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents); box.addWidget(self.summary)


class PipelinePanel(QWidget):
    selected = Signal(str)
    def __init__(self):
        super().__init__()
        box=QVBoxLayout(self); box.setContentsMargins(20,18,20,14)
        row=QHBoxLayout(); title=QLabel('파이프라인'); title.setObjectName('sectionTitle'); row.addWidget(title)
        row.addStretch(); hint=QLabel('노드를 선택해 설정'); hint.setObjectName('muted'); row.addWidget(hint); box.addLayout(row)
        self.splitter=QSplitter(Qt.Orientation.Vertical); box.addWidget(self.splitter,1)
        canvas=QWidget(); canvas.setMinimumHeight(470); center=QVBoxLayout(canvas); center.setContentsMargins(15,20,15,20)
        center.addStretch()
        self.nodes={}
        for index,key in enumerate(NODES):
            card=NodeCard(key); card.setMaximumWidth(500); self.nodes[key]=card
            card.clicked.connect(lambda checked=False,k=key:self.selected.emit(k))
            center.addWidget(card,0,Qt.AlignmentFlag.AlignHCenter)
            if index<3:
                arrow=QLabel('↓'); arrow.setFixedHeight(18); arrow.setAlignment(Qt.AlignmentFlag.AlignCenter); arrow.setObjectName('muted'); center.addWidget(arrow)
        center.addStretch()
        canvas_scroll=QScrollArea(); canvas_scroll.setWidgetResizable(True); canvas_scroll.setWidget(canvas); canvas_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.splitter.addWidget(canvas_scroll)
        logs=QWidget(); logbox=QVBoxLayout(logs); logbox.setContentsMargins(0,8,0,0)
        row=QHBoxLayout(); title=QLabel('실행 로그'); title.setObjectName('sectionTitle'); row.addWidget(title); row.addStretch()
        self.toggle=QPushButton('접기'); self.toggle.setMaximumWidth(55); row.addWidget(self.toggle); logbox.addLayout(row)
        self.log=QPlainTextEdit(); self.log.setReadOnly(True); self.log.setMaximumBlockCount(1000); self.log.setMinimumHeight(70)
        self.log.setPlaceholderText('실행하면 파일·단계·처리 결과가 표시됩니다.'); logbox.addWidget(self.log)
        self.toggle.clicked.connect(self.toggle_log); self.splitter.addWidget(logs); self.splitter.setSizes([540,150])
    def toggle_log(self):
        self.log.setVisible(self.log.isHidden()); self.toggle.setText('펼치기' if self.log.isHidden() else '접기')
    def render(self,settings,selected,states=None):
        for key,card in self.nodes.items():
            card.setChecked(key==selected)
            card.state.setText(STATES.get((states or {}).get(key,'ready'),'○ 준비'))
            if key in ('translate','adjust'):
                cfg=getattr(settings,key)
                card.summary.setText('건너뜀' if not cfg.enabled else f'{PROVIDERS[cfg.provider]} · {cfg.model or "모델 미선택"}' + (f' · {cfg.effort}' if cfg.effort else ''))
            if key=='save': card.summary.setText(settings.save.filename)
