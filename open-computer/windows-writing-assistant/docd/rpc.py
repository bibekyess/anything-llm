"""Line-delimited JSON-RPC over stdio (design doc §1.2).

Threading model (§1.3): the stdin reader thread only parses JSON and enqueues
jobs; ALL driver work (COM traffic) happens on the main thread, which calls
pythoncom.CoInitialize() when a COM backend is active. Responses are written
from the worker with a stdout lock.

    request:  {"id": "42", "method": "doc_read", "params": {...}}
    response: {"id": "42", "result": {...}}
              {"id": "42", "error": {"code": "STALE_RANGE", "message": "...", "data": {...}}}
"""

import json
import queue
import sys
import threading
import traceback

from .errors import DocdError, BAD_PARAMS, COM_ERROR

_SHUTDOWN = object()


class RpcServer:
    def __init__(self, dispatcher, use_com=False):
        self.dispatcher = dispatcher
        self.use_com = use_com
        self.jobs = queue.Queue()
        self.out_lock = threading.Lock()

    def send(self, obj):
        line = json.dumps(obj, ensure_ascii=False)
        with self.out_lock:
            sys.stdout.write(line + "\n")
            sys.stdout.flush()

    def _reader(self):
        for raw in sys.stdin:
            raw = raw.strip()
            if not raw:
                continue
            try:
                msg = json.loads(raw)
            except json.JSONDecodeError as e:
                self.send({"id": None, "error": {
                    "code": BAD_PARAMS, "message": f"Invalid JSON: {e}", "data": {},
                }})
                continue
            self.jobs.put(msg)
        self.jobs.put(_SHUTDOWN)  # EOF: parent went away

    def serve(self):
        """Run until stdin closes. Call from the main thread."""
        if self.use_com:
            import pythoncom
            pythoncom.CoInitialize()  # STA; all COM stays on this thread
        threading.Thread(target=self._reader, daemon=True, name="stdin-reader").start()
        try:
            while True:
                msg = self.jobs.get()
                if msg is _SHUTDOWN:
                    break
                self._handle(msg)
        finally:
            self.dispatcher.shutdown()
            if self.use_com:
                import gc
                import pythoncom
                gc.collect()  # release COM proxies before CoUninitialize (§1.3)
                pythoncom.CoUninitialize()

    def _handle(self, msg):
        msg_id = msg.get("id")
        method = msg.get("method")
        params = msg.get("params") or {}
        try:
            result = self.dispatcher.handle(method, params)
            self.send({"id": msg_id, "result": result})
        except DocdError as e:
            self.send({"id": msg_id, "error": e.to_json()})
        except TypeError as e:
            # Wrong/missing kwargs from the client -> BAD_PARAMS, not a crash.
            self.send({"id": msg_id, "error": {
                "code": BAD_PARAMS, "message": str(e), "data": {},
            }})
        except Exception as e:  # never let one job kill the sidecar
            self.send({"id": msg_id, "error": {
                "code": COM_ERROR,
                "message": f"{type(e).__name__}: {e}",
                "data": {"traceback": traceback.format_exc(limit=5)},
            }})
