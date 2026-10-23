from pathlib import Path
from PySide6.QtCore import Signal,Qt
from PySide6.QtWidgets import (QWidget,QVBoxLayout,QHBoxLayout,QLabel,QScrollArea,QComboBox,
    QCheckBox,QLineEdit,QPlainTextEdit,QPushButton,QFileDialog,QGroupBox,QSpinBox,QSizePolicy)
from ..models import NODES,PROVIDERS,LANGUAGES,validate_model
from .quota_panel import QuotaPanel


class Inspector(QScrollArea):
    changed=Signal()
    connection_requested=Signal()
    refresh_requested=Signal(str)
    def __init__(self,settings,connections):
        super().__init__()
        self.settings,self.connections,self.node,self.source=settings,connections,'translate',None
        self.setWidgetResizable(True); self.setMinimumWidth(300); self.setObjectName('inspector')
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.body=QWidget(); self.setWidget(self.body)
        self.body.setSizePolicy(QSizePolicy.Policy.Ignored,QSizePolicy.Policy.Preferred)
        self.box=QVBoxLayout(self.body); self.box.setContentsMargins(20,18,20,20); self.box.setSpacing(10)
        self.render('translate')

    def label(self,text,muted=False):
        label=QLabel(text); label.setWordWrap(True)
        if muted: label.setObjectName('muted')
        self.box.addWidget(label); return label

    def render(self,node=None,source=None):
        if node: self.node=node
        if source is not None: self.source=source
        while self.box.count():
            item=self.box.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self.label('노드 속성',True)
        title=self.label(NODES[self.node][0]); title.setObjectName('sectionTitle')
        self.label(NODES[self.node][1],True)
        if self.node=='extract':
            self.label('원본 파일을 읽고 페이지·도형별 텍스트와 이미지를 추출합니다. 머리말·꼬리말·페이지 번호는 제외합니다.')
            self.label('수식은 네이티브 XML을 보존합니다. 이미지 속 문자는 OCR하지 않습니다.',True)
        elif self.node=='save': self.save_form()
        else: self.ai_form()
        self.box.addStretch()

    def ai_form(self):
        node=self.node; cfg=getattr(self.settings,node)
        enabled=QCheckBox('이 단계 실행'); enabled.setChecked(cfg.enabled)
        enabled.toggled.connect(lambda v:self.set_value('enabled',v)); self.box.addWidget(enabled)
        self.label('실행 도구',True)
        provider=QComboBox()
        for key,label in PROVIDERS.items():
            provider.addItem(label,key)
        provider.setCurrentIndex(provider.findData(cfg.provider)); self.box.addWidget(provider)
        provider.currentIndexChanged.connect(lambda i:self.provider_changed(provider.itemData(i)))
        state=self.connections.state(cfg.provider)
        models=[m for m in state.models if node!='adjust' or m.vision]
        from copy import copy
        compatible=copy(state); compatible.models=models
        selected=validate_model(cfg,compatible)
        self.label(('연결됨 · ' if state.connected else '미연결 · ')+state.account,True)
        manage=QPushButton('AI 연결 관리에서 계정 설정'); manage.clicked.connect(self.connection_requested); self.box.addWidget(manage)
        quota=QuotaPanel(); quota.render(state); quota.refresh_requested.connect(lambda:self.refresh_requested.emit(cfg.provider)); self.box.addWidget(quota)
        self.label('모델',True)
        model=QComboBox()
        if models:
            for entry in models: model.addItem(entry.label,entry.id)
            model.setCurrentIndex(model.findData(cfg.model))
        else:
            message='모델 조회 중' if cfg.provider in self.connections.tasks else '연결 후 사용 가능한 모델 없음'
            model.addItem(message,''); model.setEnabled(False)
        self.box.addWidget(model)
        model.currentIndexChanged.connect(lambda i:self.model_changed(model.itemData(i)))
        if state.error: self.label(state.error,True)
        if node=='adjust' and cfg.provider=='gemini':
            self.label('공식 Gemini CLI의 이미지 검토는 미검증입니다. 조정은 Codex를 사용하세요.',True)
        self.label('추론 수준',True)
        effort=QComboBox()
        if selected and selected.efforts:
            for value in selected.efforts: effort.addItem(value,value)
            effort.setCurrentIndex(effort.findData(cfg.effort))
            effort.currentIndexChanged.connect(lambda i:self.set_value('effort',effort.itemData(i)))
        else:
            effort.addItem('이 도구·모델에서 조회되지 않음'); effort.setEnabled(False)
        self.box.addWidget(effort)
        if cfg.provider=='antigravity':
            self.label('High/Medium/Low는 조회된 모델 이름에서 선택합니다. 별도 추론 플래그는 적용하지 않습니다.',True)
        if node=='translate':
            for attr,label in [('source_language','원본 언어'),('target_language','대상 언어')]:
                self.label(label,True); combo=QComboBox()
                for key,value in LANGUAGES.items(): combo.addItem(value,key)
                combo.setCurrentIndex(combo.findData(getattr(cfg,attr)))
                combo.currentIndexChanged.connect(lambda i,c=combo,a=attr:self.set_value(a,c.itemData(i)))
                self.box.addWidget(combo)
            preserve=QCheckBox('코드·수식·숫자 보존'); preserve.setChecked(True); preserve.setEnabled(False)
            preserve.setToolTip('원본 보존을 위해 항상 적용합니다.'); self.box.addWidget(preserve)
            self.label('용어집 · 한 줄에 한 용어',True)
            glossary=QPlainTextEdit(cfg.glossary); glossary.setPlaceholderText('전류 = current'); glossary.setMaximumHeight(80)
            glossary.textChanged.connect(lambda:self.set_value('glossary',glossary.toPlainText())); self.box.addWidget(glossary)
        else:
            self.label('조정 목적',True); self.label('텍스트 넘침·그림 겹침·잘림 검토')
            for text in ('원래 의미·수식 유지','기존 글자 크기 우선 유지'):
                check=QCheckBox(text); check.setChecked(True); check.setEnabled(False); self.box.addWidget(check)
            self.label('용어·문체 재작성은 현재 미지원입니다. 조정 단계는 텍스트 내용을 바꾸지 않습니다.',True)
            self.label('최대 수정 횟수',True)
            passes=QSpinBox(); passes.setRange(1,5); passes.setValue(cfg.max_passes)
            passes.valueChanged.connect(lambda v:self.set_value('max_passes',v)); self.box.addWidget(passes)
        advanced=QGroupBox('고급 설정 · 추가 프롬프트'); advanced.setCheckable(True); advanced.setChecked(bool(cfg.prompt))
        inner=QVBoxLayout(advanced); prompt=QPlainTextEdit(cfg.prompt); prompt.setMaximumHeight(110)
        prompt.setPlaceholderText('예: 강의용 표현을 사용하세요.'); prompt.textChanged.connect(lambda:self.set_value('prompt',prompt.toPlainText()))
        inner.addWidget(prompt); prompt.setVisible(advanced.isChecked()); advanced.toggled.connect(prompt.setVisible); self.box.addWidget(advanced)
        self.label('작업 중 변경은 다음 파일부터 적용됩니다. 한도 소진으로 대기 중일 때는 연결 변경 후 재개할 수 있습니다.',True)

    def set_value(self,key,value):
        setattr(getattr(self.settings,self.node),key,value); self.changed.emit()
    def provider_changed(self,value):
        cfg=getattr(self.settings,self.node); cfg.provider=value; cfg.model=''; cfg.effort=''
        self.render(); self.changed.emit()
    def model_changed(self,value):
        cfg=getattr(self.settings,self.node); cfg.model=value
        self.render(); self.changed.emit()

    def save_form(self):
        cfg=self.settings.save
        self.label('파일 저장 위치',True)
        folder=QLineEdit(cfg.directory); self.box.addWidget(folder)
        browse=QPushButton('폴더 선택…'); self.box.addWidget(browse)
        def choose():
            value=QFileDialog.getExistingDirectory(self,'PPT 저장 폴더',folder.text())
            if value: folder.setText(value)
        browse.clicked.connect(choose)
        self.label('파일 이름',True)
        name=QLineEdit(cfg.filename); self.box.addWidget(name)
        self.label('기본: 원본 PPT 이름_최종언어.pptx\n{stem}: 원본 이름 · {language}: 대상 언어',True)
        preview=self.label(''); preview.setTextInteractionFlags(preview.textInteractionFlags())
        def update():
            cfg.directory=folder.text(); cfg.filename=name.text()
            try:
                path=cfg.target(self.source or Path('강의자료.pptx'),self.settings.translate.target_language)
                preview.setText('저장 예시\n'+str(path))
            except ValueError as exc: preview.setText(str(exc))
            self.changed.emit()
        folder.textChanged.connect(update); name.textChanged.connect(update); update()
        self.label('원본 및 기존 결과 파일은 덮어쓰지 않습니다. 실행 전 출력 경로를 확인합니다.',True)
