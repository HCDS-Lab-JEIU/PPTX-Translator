"""Main composition root for the native PPT translator desktop application."""
import argparse
import os
from pathlib import Path
import sys


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--demo',action='store_true',help='명시적인 모의 데이터 · AI 호출/실제 PPT 저장 없음')
    parser.add_argument('--screenshot',type=Path,help='검증용 화면 PNG 저장')
    parser.add_argument('--quit-after',type=int,help='검증용 자동 종료 시간(ms)')
    args=parser.parse_args(argv)
    try:
        from PySide6.QtCore import QTimer
        from PySide6.QtGui import QFont,QFontDatabase
        from PySide6.QtWidgets import QApplication
        from pptx_gui.main_window import MainWindow
        from pptx_gui.style import STYLE
    except ImportError as exc:
        print(f'GUI 의존성을 설치하세요: python -m pip install -r requirements-gui.txt\n{exc}',file=sys.stderr)
        return 1
    app=QApplication.instance() or QApplication(sys.argv[:1])
    family='Malgun Gothic'
    # Offscreen Qt on Windows may not discover system fonts automatically.
    for filename in ('malgun.ttf','segoeui.ttf','seguisym.ttf'):
        font=Path(os.environ.get('SystemRoot', 'C:/Windows'))/'Fonts'/filename
        if font.is_file():
            identifier=QFontDatabase.addApplicationFont(str(font))
            families=QFontDatabase.applicationFontFamilies(identifier)
            if filename=='malgun.ttf' and families: family=families[0]
    app.setApplicationName('PPT Translator'); app.setStyle('Fusion'); app.setFont(QFont(family,10)); app.setStyleSheet(STYLE)
    window=MainWindow(demo=args.demo); window.show()
    if args.screenshot: window.resize(1440,960)
    if args.screenshot:
        def capture():
            args.screenshot.parent.mkdir(parents=True,exist_ok=True)
            window.grab().save(str(args.screenshot))
        QTimer.singleShot(900,capture)
    if args.quit_after: QTimer.singleShot(args.quit_after,window.close)
    return app.exec()


if __name__=='__main__': raise SystemExit(main())
