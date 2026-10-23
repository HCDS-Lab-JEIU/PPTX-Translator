STYLE = '''
QMainWindow, QWidget { background: #f4f6f8; color: #202a36; font-size: 13px; }
QLabel { background: transparent; }
QWidget#filesPanel, QScrollArea#inspector, QScrollArea#inspector > QWidget > QWidget,
QWidget#progressPanel, QToolBar { background: white; }
QWidget#filesPanel { border-right: 1px solid #dde3ea; }
QScrollArea#inspector { border: 0; border-left: 1px solid #dde3ea; }
QToolBar { border: 0; border-bottom: 1px solid #dde3ea; spacing: 10px; padding: 14px; }
QStatusBar { background: white; border-top: 1px solid #dde3ea; color: #647080; }
QLabel#appTitle { font-size: 18px; font-weight: 600; }
QLabel#sectionTitle, QLabel#nodeTitle { font-size: 15px; font-weight: 600; }
QLabel#muted { color: #647080; font-size: 12px; }
QLabel#nodeState { color: #176783; font-size: 12px; }
QPushButton { background: white; border: 1px solid #dde3ea; border-radius: 7px; padding: 7px 11px; }
QPushButton:hover { background: #edf6fa; border-color: #83afbf; }
QPushButton:disabled { background: #f5f6f7; color: #929aa5; }
QPushButton#primary { background: #176783; color: white; border-color: #176783; font-weight: 600; }
QPushButton#primary:hover { background: #12566e; }
QPushButton#nodeCard { border-radius: 9px; text-align: left; background: white; }
QPushButton#nodeCard:checked { border: 2px solid #176783; background: #e9f4f8; }
QPushButton#nodeCard:checked QLabel { background: transparent; }
QComboBox, QLineEdit, QPlainTextEdit, QSpinBox { background: white; border: 1px solid #dde3ea; border-radius: 6px; padding: 7px; selection-background-color: #d0eaf4; }
QComboBox:disabled { color: #929aa5; background: #f4f6f8; }
QTreeWidget { background: white; border: 0; outline: none; }
QTreeWidget::item { padding: 12px 3px; border-bottom: 1px solid #edf0f3; }
QTreeWidget::item:selected { background: #e9f4f8; color: #176783; }
QTabWidget::pane { border: 0; background: white; }
QTabBar::tab { background: white; padding: 8px 12px; border-bottom: 2px solid #dde3ea; }
QTabBar::tab:selected { border-bottom: 2px solid #176783; color: #176783; }
QFrame#quotaPanel { background: #e9f4f8; border-radius: 8px; }
QProgressBar { background: #dde3ea; border: 0; border-radius: 3px; }
QProgressBar::chunk { background: #176783; border-radius: 3px; }
QSplitter::handle { background: #e5e9ee; }
QSplitter::handle:horizontal { width: 3px; }
QSplitter::handle:vertical { height: 4px; }
QGroupBox { border: 1px solid #dde3ea; border-radius: 7px; margin-top: 12px; padding-top: 15px; }
QGroupBox::title { subcontrol-origin: margin; left: 9px; }
QCheckBox { background: transparent; spacing: 8px; }
QCheckBox::indicator { width: 15px; height: 15px; }
'''
