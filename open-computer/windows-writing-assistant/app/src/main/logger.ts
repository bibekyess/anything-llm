/**
 * Central logger (electron-log). Files land in:
 *   Windows: %APPDATA%\writing-assistant-app\logs\main.log
 * Console output mirrors the file in dev.
 */
import log from "electron-log/main";

log.initialize();
log.transports.file.level = "debug";
log.transports.console.level = "debug";
log.transports.file.maxSize = 5 * 1024 * 1024;
log.transports.file.format = "[{y}-{m}-{d} {h}:{i}:{s}.{ms}] [{level}] {text}";

export default log;
