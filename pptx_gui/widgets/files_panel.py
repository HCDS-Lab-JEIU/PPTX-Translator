from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTabWidget, QTreeWidget, QTreeWidgetItem
from ..models import STATES
from ..review_summary import review_message


class FileTree(QTreeWidget):
    paths_dropped = Signal(list)
    def __init__(self):
        super().__init__()
        self.setColumnCount(1)
        self.setHeaderHidden(True)
        self.setRootIsDecorated(False)
        self.setTextElideMode(Qt.TextElideMode.ElideMiddle)
        self.setAcceptDrops(True)
        self.setAlternatingRowColors(False)
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls(): event.acceptProposedAction()
    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls(): event.acceptProposedAction()
    def dropEvent(self, event):
        self.paths_dropped.emit([u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()])
        event.acceptProposedAction()


class FilesPanel(QWidget):
    add_requested = Signal()
    remove_requested = Signal(str)
    move_requested = Signal(str,int)
    rerun_requested = Signal(str)
    open_requested = Signal(str,bool)
    dropped = Signal(list)
    selected = Signal(str)
    def __init__(self):
        super().__init__()
        self.setObjectName('filesPanel')
        self.setMinimumWidth(200)
        box = QVBoxLayout(self); box.setContentsMargins(14,18,14,14); box.setSpacing(12)
        title=QLabel('프레젠테이션'); title.setObjectName('sectionTitle'); box.addWidget(title)
        self.tabs=QTabWidget(); box.addWidget(self.tabs,1)
        self.queue,self.done=FileTree(),FileTree()
        self.tabs.addTab(self.queue,'대기 0'); self.tabs.addTab(self.done,'완료 0')
        for tree in (self.queue,self.done):
            tree.paths_dropped.connect(self.dropped)
            tree.itemSelectionChanged.connect(self._selected)
        self.hint=QLabel('PPTX 파일을 여기에 놓거나\n파일 추가를 선택하세요.'); self.hint.setWordWrap(True); self.hint.setObjectName('muted')
        box.addWidget(self.hint)
        add=QPushButton('＋ 파일 추가'); add.clicked.connect(self.add_requested); box.addWidget(add)
        row=QHBoxLayout()
        self.remove=QPushButton('제거'); up=QPushButton('↑'); down=QPushButton('↓')
        self.remove.clicked.connect(lambda:self.remove_requested.emit(self.selected_id()))
        up.clicked.connect(lambda:self.move_requested.emit(self.selected_id(),-1))
        down.clicked.connect(lambda:self.move_requested.emit(self.selected_id(),1))
        for b in (self.remove,up,down): row.addWidget(b)
        box.addLayout(row)
        row=QHBoxLayout()
        for label, folder in [('결과 열기',False),('폴더 열기',True)]:
            button=QPushButton(label); button.clicked.connect(lambda checked=False,f=folder:self.open_requested.emit(self.selected_id(),f)); row.addWidget(button)
        box.addLayout(row)
        retry=QPushButton('선택 파일 재실행'); retry.clicked.connect(lambda:self.rerun_requested.emit(self.selected_id())); box.addWidget(retry)
    def selected_id(self):
        tree=self.queue if self.tabs.currentIndex()==0 else self.done
        item=tree.currentItem()
        return item.data(0,Qt.ItemDataRole.UserRole) if item else ''
    def _selected(self): self.selected.emit(self.selected_id())
    def render(self,jobs):
        selected=self.selected_id()
        for tree in (self.queue,self.done): tree.blockSignals(True); tree.clear()
        waiting=done=0
        for job in jobs:
            finished=job.status in ('done','needs_review')
            tree=self.done if finished else self.queue
            done+=int(finished); waiting+=int(not finished)
            slides=f'{job.slides} slides' if job.slides is not None else '분석 대기'
            item=QTreeWidgetItem([f'{job.source.name}\n{slides} · {STATES.get(job.status,job.status)}'])
            item.setData(0,Qt.ItemDataRole.UserRole,job.id); item.setToolTip(0,str(job.source))
            if finished and job.review_required_pages:
                summary=review_message(job.review_required_pages)
                item.setText(0,item.text(0)+'\n'+summary)
                item.setToolTip(0,str(job.source)+'\n'+summary)
            tree.addTopLevelItem(item)
            if job.id==selected: tree.setCurrentItem(item)
        for tree in (self.queue,self.done): tree.blockSignals(False)
        self.tabs.setTabText(0,f'대기 {waiting}'); self.tabs.setTabText(1,f'완료 {done}')
        self.hint.setVisible(not jobs)
