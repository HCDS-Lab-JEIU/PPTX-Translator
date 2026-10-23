"""Codex integration over app-server stdio (the former MCP server was removed)."""

import json
import math
import os
from pathlib import Path
from queue import Queue, Empty
import shutil
import subprocess
from threading import Thread
import time
from uuid import uuid4

from gemini_cli import CLIResult, EmptyModelResponseError
from cli_discovery import codex_candidates, native


class CodexRPCError(RuntimeError):
    def __init__(self, message, result=None):
        super().__init__(message)
        self.result = result


def find_codex(executable=None):
    configured = executable or os.environ.get("CODEX_EXECUTABLE")
    if configured:
        found = shutil.which(str(configured))
        candidate = Path(found or configured).expanduser()
    else:
        candidate = next((p for p in codex_candidates() if native(p)),None)
        if candidate is None:
            raise FileNotFoundError("Codex 자동 탐색 실패: PATH·기본 설치 폴더·VS Code 확장을 확인하세요.")
    if not candidate.is_file():
        raise FileNotFoundError("Codex 실행 파일을 찾을 수 없습니다. --codex-executable로 codex.exe 경로를 지정하세요.")
    if candidate.suffix.lower() in {".cmd", ".bat", ".ps1"}:
        raise ValueError("셸 래퍼 대신 실제 codex 실행 파일을 지정하세요.")
    return str(candidate.resolve())


class StdioRPC:
    """One bounded stdio session. Drain stderr independently to avoid deadlocks."""
    def __init__(self, command, *, cwd, timeout):
        self.command, self.timeout = tuple(command), timeout
        self.events, self.stderr = [], []
        self.queue = Queue()
        self.sequence = 0
        self.event_callback = None
        self.deadline = time.monotonic() + timeout
        self.process = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace", cwd=cwd,
            shell=False, creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)
        def stdout():
            try:
                for line in self.process.stdout:
                    try:
                        self.queue.put(json.loads(line))
                    except ValueError:
                        self.queue.put(CodexRPCError("Codex RPC가 JSON이 아닌 응답을 반환했습니다."))
            finally:
                self.queue.put(None)
        def stderr():
            for line in self.process.stderr:
                self.stderr.append(line)
        self.readers = [Thread(target=stdout, daemon=True), Thread(target=stderr, daemon=True)]
        for reader in self.readers:
            reader.start()

    def send(self, message):
        try:
            self.process.stdin.write(json.dumps(message, ensure_ascii=False) + "\n")
            self.process.stdin.flush()
        except (OSError, ValueError) as exc:
            raise CodexRPCError("Codex RPC 연결이 닫혔습니다.") from exc

    def receive(self):
        try:
            event = self.queue.get(timeout=max(0, self.deadline - time.monotonic()))
        except Empty as exc:
            raise CodexRPCError(f"Codex 응답 시간이 {self.timeout}초를 초과했습니다.") from exc
        if event is None:
            raise CodexRPCError("Codex App Server가 응답 도중 종료됐습니다. " + "".join(self.stderr)[-2000:])
        if isinstance(event, Exception):
            raise event
        self.events.append(event)
        if self.event_callback is not None:
            self.event_callback(event)
        # Headless reviews must not approve unexpected operations or wait for UI input.
        if "method" in event and "id" in event:
            self.send({"id": event["id"], "error": {"code": -32601,
                "message": "This read-only review client does not support interactive requests."}})
            raise CodexRPCError(f"검토 중 지원하지 않는 상호작용 요청: {event['method']}")
        return event

    def request(self, method, params):
        self.sequence += 1
        identifier = self.sequence
        self.send({"id": identifier, "method": method, "params": params})
        while True:
            event = self.receive()
            if event.get("id") == identifier and "method" not in event:
                if "error" in event:
                    raise CodexRPCError(f"Codex {method} 실패: {event['error']}")
                return event["result"]

    def close(self):
        try:
            self.process.stdin.close()
            self.process.wait(timeout=3)
        except (OSError, subprocess.TimeoutExpired):
            self.process.kill()
            self.process.wait(timeout=3)
        finally:
            for reader in self.readers:
                reader.join(timeout=1)
            self.process.stdout.close()
            self.process.stderr.close()


class CodexAppServerProvider:
    """AIProvider using native localImage inputs and existing Codex authentication."""
    def __init__(self, *, executable=None, model=None, timeout=300, cwd=None,
                 effort="low", trace_dir=None, rpc_factory=StdioRPC, on_event=None):
        if not math.isfinite(timeout) or timeout <= 0:
            raise ValueError("timeout은 양수여야 합니다.")
        if effort is not None and effort not in {"low", "medium", "high", "xhigh", "max"}:
            raise ValueError("지원하지 않는 Codex reasoning effort입니다.")
        self.executable = find_codex(executable)
        self.model, self.effort, self.timeout = model, effort, timeout
        self.cwd = Path(cwd or Path.cwd()).resolve()
        self.trace_dir = Path(trace_dir).resolve() if trace_dir else None
        self.rpc_factory = rpc_factory
        self.active_rpc = None
        self.on_event = on_event

    def cancel(self):
        rpc = self.active_rpc
        if rpc is not None and rpc.process.poll() is None:
            try: rpc.process.kill()
            except OSError: pass

    def generate(self, prompt, *, model=None, images=()):
        from .ai import AIResponse
        if not isinstance(prompt, str) or not prompt.strip():
            raise ValueError("prompt가 비어 있습니다.")
        images = tuple(Path(path).resolve() for path in images)
        for path in images:
            if not path.is_file():
                raise FileNotFoundError(str(path))
        command = (self.executable, "app-server", "--stdio")
        rpc = self.rpc_factory(command, cwd=self.cwd, timeout=self.timeout)
        self.active_rpc = rpc
        rpc.event_callback = self.on_event
        result = None
        try:
            rpc.request("initialize", {"clientInfo": {"name": "pptx_layout_pipeline",
                "title": "PPTX layout reviewer", "version": "1.0.0"}})
            rpc.send({"method": "initialized", "params": {}})
            chosen_model = model or self.model
            if not chosen_model:
                available = rpc.request("model/list", {})["data"]
                eligible = [m for m in available if not images or "image" in m.get("inputModalities", [])]
                selected = next((m for m in eligible if m.get("isDefault")), eligible[0] if eligible else None)
                if selected is None:
                    raise CodexRPCError("이미지 검토에 사용할 수 있는 Codex 모델이 없습니다.")
                chosen_model = selected["model"]
            options = {"cwd": str(self.cwd), "approvalPolicy": "never", "sandbox": "read-only",
                "ephemeral": True,
                "developerInstructions": "Review the attached slide images and return JSON only. "
                    "Do not use tools, modify files, browse, or launch other agents.",
                "config": {"model_reasoning_effort": self.effort} if self.effort else {}}
            if chosen_model:
                options["model"] = chosen_model
            thread = rpc.request("thread/start", options)
            self.model = thread.get("model", chosen_model)
            thread_id = thread["thread"]["id"]
            inputs = [{"type": "text", "text": prompt}]
            inputs.extend({"type": "localImage", "path": str(path)} for path in images)
            turn_options={"threadId": thread_id, "input": inputs}
            if self.effort: turn_options["effort"]=self.effort
            turn = rpc.request("turn/start", turn_options)["turn"]
            turn_id = turn["id"]
            messages = {}
            cursor = 0
            while True:
                if cursor == len(rpc.events):
                    rpc.receive()
                event = rpc.events[cursor]
                cursor += 1
                params = event.get("params", {})
                if params.get("threadId") != thread_id:
                    continue
                if event.get("method") == "item/completed" and params.get("turnId") == turn_id:
                    item = params.get("item", {})
                    if item.get("type") == "agentMessage":
                        messages[item["id"]] = item
                if event.get("method") == "turn/completed" and params.get("turn", {}).get("id") == turn_id:
                    final = params["turn"]
                    if final.get("status") != "completed":
                        raise CodexRPCError("Codex 검토 실패: " + json.dumps(final.get("error") or final, ensure_ascii=False))
                    finals = [m["text"] for m in messages.values() if m.get("phase") == "final_answer"]
                    # Older protocol versions omit phase; use only the final message.
                    if not finals:
                        finals = [m["text"] for m in messages.values() if m.get("phase") is None][-1:]
                    text = "\n".join(finals)
                    if not text.strip():
                        raise EmptyModelResponseError("Codex가 빈 응답을 반환했습니다.")
                    # Native attachments are consumed as input, no file-view tool evidence is needed.
                    result = AIResponse(text, images)
                    return result
        except (CodexRPCError, EmptyModelResponseError) as exc:
            exc.result = CLIResult(command, -1, "\n".join(json.dumps(e, ensure_ascii=False) for e in rpc.events),
                                   "".join(rpc.stderr), tuple(rpc.events))
            raise
        finally:
            rpc.close()
            self.active_rpc = None
            if self.trace_dir:
                self.trace_dir.mkdir(parents=True, exist_ok=True)
                trace = {"model": self.model, "effort": self.effort, "images": list(map(str, images)),
                         "response": result.text if result else None, "events": rpc.events,
                         "stderr": "".join(rpc.stderr)}
                (self.trace_dir / f"{uuid4().hex}.json").write_text(
                    json.dumps(trace, ensure_ascii=False, indent=2), encoding="utf-8")
