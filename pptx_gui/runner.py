from PySide6.QtCore import QThread, Signal
from .engine import DesktopEngine, RunControl, Stopped
from .adapters.base import safe_error


class FileRunner(QThread):
    event=Signal(object)
    def __init__(self,jobs,settings,executables,demo=False,parent=None):
        super().__init__(parent)
        self.jobs=jobs
        self.control=RunControl()
        self.engine=DesktopEngine(settings,executables,self.event.emit,self.control,demo)
    def run(self):
        for job in self.jobs:
            if self.control.stop_event.is_set(): break
            try:
                result=self.engine.run_file(job)
                self.event.emit({'type':'file_done','job':job.id,'result':result})
            except Stopped:
                self.event.emit({'type':'file_stopped','job':job.id}); break
            except Exception as exc:
                self.event.emit({'type':'file_failed','job':job.id,'message':safe_error(exc)})
        self.event.emit({'type':'batch_finished'})
