import subprocess
from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QDialog,QVBoxLayout,QLabel,QGroupBox,QLineEdit,QHBoxLayout,QPushButton,QFileDialog,QMessageBox
from ..models import PROVIDERS
from ..adapters.gemini import find_gemini
from gemini_cli import find_executable


class ConnectionsDialog(QDialog):
    def __init__(self,manager,auth_locked=False,parent=None):
        super().__init__(parent); self.manager=manager; self.auth_locked=auth_locked
        self.setWindowTitle('AI 연결 관리'); self.resize(680,650)
        box=QVBoxLayout(self)
        note=QLabel('각 PC의 현재 OS 사용자 계정으로 실행합니다.\n비밀번호·API 키를 앱에서 수집하지 않으며 인증은 공식 도구가 관리합니다.')
        note.setWordWrap(True); box.addWidget(note)
        self.rows={}
        for key in PROVIDERS:
            group=QGroupBox(PROVIDERS[key]); inner=QVBoxLayout(group)
            status=QLabel(); status.setWordWrap(True); inner.addWidget(status)
            path=QLineEdit(manager.adapters[key].executable or '')
            path.setPlaceholderText('자동 탐색 · 실제 실행 파일 경로'); inner.addWidget(path)
            row=QHBoxLayout(); inner.addLayout(row)
            browse=QPushButton('경로 선택'); browse.clicked.connect(lambda checked=False,k=key:self.choose(k)); row.addWidget(browse)
            refresh=QPushButton('설치·연결 조회'); refresh.clicked.connect(lambda checked=False,k=key:self.refresh(k)); row.addWidget(refresh)
            test=QPushButton('연결 테스트'); test.clicked.connect(lambda checked=False,k=key:self.refresh(k,'test')); row.addWidget(test)
            row=QHBoxLayout(); inner.addLayout(row)
            login=QPushButton('공식 로그인'); logout=QPushButton('로그아웃')
            login.clicked.connect(lambda checked=False,k=key:self.auth(k,'login'))
            logout.clicked.connect(lambda checked=False,k=key:self.auth(k,'logout'))
            row.addWidget(login); row.addWidget(logout)
            self.rows[key]=(status,path,[browse,refresh,test,login,logout]); box.addWidget(group)
        close=QPushButton('닫기'); close.clicked.connect(self.accept); box.addWidget(close)
        manager.changed.connect(self.render); self.render()

    def choose(self,key):
        path,_=QFileDialog.getOpenFileName(self,'실제 CLI 실행 파일 선택')
        if path:
            self.rows[key][1].setText(path); self.rows[key][1].setModified(True)
    def refresh(self,key,action='refresh'):
        field=self.rows[key][1]
        if field.isModified():
            self.manager.adapters[key].executable=field.text().strip() or None
            field.setModified(False)
        self.manager.refresh(key,action)
    def auth(self,key,action):
        if self.manager.demo:
            QMessageBox.information(self,'모의 데이터','데모 모드에서는 실제 로그인·로그아웃을 실행하지 않습니다.'); return
        if key in ('gemini','antigravity'):
            try:
                command=(find_gemini(self.rows[key][1].text().strip() or None) if key=='gemini' else
                         [find_executable(self.rows[key][1].text().strip() or None)])
                subprocess.Popen(command,creationflags=getattr(subprocess,'CREATE_NEW_CONSOLE',0))
                instruction='로그인 안내를 따르세요. 로그아웃은 CLI의 /logout을 사용하세요.' if key=='antigravity' else '/auth로 로그인하세요.'
                QMessageBox.information(self,'공식 CLI 인증',
                    '열린 '+PROVIDERS[key]+'에서 '+instruction+
                    '\n완료 후 이 화면에서 연결 테스트를 실행하세요. 앱은 인증 저장소를 직접 수정하지 않습니다.')
            except Exception:
                QMessageBox.warning(self,'CLI 실행 실패',PROVIDERS[key]+' 설치와 실행 경로를 확인하세요.')
            return
        if action=='logout' and QMessageBox.question(self,'Codex 로그아웃','이 PC의 Codex 공유 로그인에서 로그아웃합니다. 계속할까요?')!=QMessageBox.StandardButton.Yes: return
        self.refresh(key,action)
    def render(self,*args):
        for key,(label,path,buttons) in self.rows.items():
            state=self.manager.state(key)
            text=('연결됨 · '+state.account if state.connected else '미연결')
            if state.demo: text+=' · 모의 데이터'
            text+='\n'+(state.version or '버전 미확인')+' · '+(state.executable or '실행 파일 미발견')
            if state.error: text+='\n'+state.error
            label.setText(text)
            for index,button in enumerate(buttons):
                button.setEnabled(key not in self.manager.tasks and (index<3 or not self.auth_locked))
            path.setEnabled(key not in self.manager.tasks)
            if state.executable and not path.hasFocus() and not path.isModified() and not self.manager.adapters[key].executable:
                path.setText(state.executable); path.setModified(False)
                path.setToolTip('자동 탐색한 경로 · 비우면 다시 자동 탐색합니다.')
            if key in ('gemini','antigravity'):
                buttons[3].setText('공식 CLI에서 인증')
                buttons[4].setEnabled(False); buttons[4].setToolTip('공식 CLI 내부 인증 메뉴에서 수행 · 직접 로그아웃 연동 미지원')
