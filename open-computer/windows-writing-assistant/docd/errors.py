"""Closed error-code enum shared with the TS extension (design doc §1.2)."""

NO_SUCH_DOC = "NO_SUCH_DOC"
STALE_RANGE = "STALE_RANGE"
READ_ONLY = "READ_ONLY"
APP_BUSY_MODAL = "APP_BUSY_MODAL"
APP_NOT_RUNNING = "APP_NOT_RUNNING"
UNSUPPORTED_ON_BACKEND = "UNSUPPORTED_ON_BACKEND"
SAVE_FORMAT_UNSUPPORTED = "SAVE_FORMAT_UNSUPPORTED"
COM_ERROR = "COM_ERROR"
TIMEOUT = "TIMEOUT"
BAD_PARAMS = "BAD_PARAMS"
UNKNOWN_METHOD = "UNKNOWN_METHOD"


class DocdError(Exception):
    """Error carried back to the TS layer as {"error": {code, message, data}}."""

    def __init__(self, code, message, data=None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.data = data or {}

    def to_json(self):
        return {"code": self.code, "message": self.message, "data": self.data}
