"""Entry point: python -m docd [--backend word|fake]

Default backend is Word on Windows, fake elsewhere. The fake backend is also
always registered in non-Word builds so protocol tests run cross-platform.
"""

import argparse
import sys


def force_utf8_stdio():
    """The RPC protocol is UTF-8, always. On Windows, Python defaults piped
    stdio to the legacy ANSI code page (cp1252/cp949), which shreds Korean
    and any other non-ASCII text into surrogate escapes."""
    for stream, errors in ((sys.stdin, "replace"), (sys.stdout, "replace"),
                           (sys.stderr, "backslashreplace")):
        try:
            stream.reconfigure(encoding="utf-8", errors=errors)
        except (AttributeError, ValueError, OSError):
            pass


def main():
    force_utf8_stdio()
    parser = argparse.ArgumentParser(prog="docd")
    parser.add_argument(
        "--backend",
        choices=["word", "fake"],
        default="word" if sys.platform == "win32" else "fake",
    )
    args = parser.parse_args()

    from .registry import Dispatcher
    from .rpc import RpcServer

    drivers = {}
    use_com = False
    if args.backend == "word":
        from .drivers.hwp import HwpDriver
        from .drivers.ppt import PptDriver
        from .drivers.word import WordDriver
        drivers["word"] = WordDriver()
        drivers["powerpoint"] = PptDriver()
        drivers["hwp"] = HwpDriver()
        use_com = True
    from .drivers.fake import FakeDriver
    from .drivers.fake_pres import FakePresDriver
    drivers["fake"] = FakeDriver()
    drivers["fakepres"] = FakePresDriver()

    RpcServer(Dispatcher(drivers), use_com=use_com).serve()


if __name__ == "__main__":
    main()
