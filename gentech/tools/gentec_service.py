import json
import os
import subprocess
import sys
from pathlib import Path

import servicemanager
import win32event
import win32service
import win32serviceutil


PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = PROJECT_ROOT / "data" / "service_config.json"
LOG_DIR = PROJECT_ROOT / "data" / "service_logs"


def _load_port(default_port: int = 5000) -> int:
    if not CONFIG_PATH.exists():
        return default_port
    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
        return int(payload.get("port", default_port))
    except Exception:
        return default_port


class GentecBillingService(win32serviceutil.ServiceFramework):
    _svc_name_ = "GentecBillingService"
    _svc_display_name_ = "Gentec Billing Service"
    _svc_description_ = "Runs Gentec Billing app in background using Waitress."

    def __init__(self, args):
        super().__init__(args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.process = None
        self.stdout_handle = None
        self.stderr_handle = None

    def _start_worker(self):
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        self.stdout_handle = (LOG_DIR / "service.out.log").open("a", encoding="utf-8")
        self.stderr_handle = (LOG_DIR / "service.err.log").open("a", encoding="utf-8")

        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        port = _load_port()
        command = [
            sys.executable,
            "-m",
            "waitress",
            "--host=0.0.0.0",
            f"--port={port}",
            "app:app",
        ]
        self.process = subprocess.Popen(
            command,
            cwd=str(PROJECT_ROOT),
            stdout=self.stdout_handle,
            stderr=self.stderr_handle,
            env=env,
        )

    def _stop_worker(self):
        if self.process and self.process.poll() is None:
            self.process.terminate()
            try:
                self.process.wait(timeout=20)
            except Exception:
                self.process.kill()
        self.process = None
        if self.stdout_handle:
            self.stdout_handle.close()
            self.stdout_handle = None
        if self.stderr_handle:
            self.stderr_handle.close()
            self.stderr_handle = None

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self._stop_worker()
        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self):
        servicemanager.LogInfoMsg("GentecBillingService starting")
        self._start_worker()

        while True:
            result = win32event.WaitForSingleObject(self.stop_event, 1000)
            if result == win32event.WAIT_OBJECT_0:
                break
            if self.process and self.process.poll() is not None:
                servicemanager.LogErrorMsg("GentecBillingService worker exited unexpectedly")
                break

        self._stop_worker()
        servicemanager.LogInfoMsg("GentecBillingService stopped")


if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(GentecBillingService)
