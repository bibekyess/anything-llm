"""Entry point: python -m docd [--backend word|fake]

Default backend is Word on Windows, fake elsewhere. The fake backend is also
always registered in non-Word builds so protocol tests run cross-platform.
"""

import argparse
import sys


def main():
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
        from .drivers.word import WordDriver
        drivers["word"] = WordDriver()
        use_com = True
    from .drivers.fake import FakeDriver
    drivers["fake"] = FakeDriver()

    RpcServer(Dispatcher(drivers), use_com=use_com).serve()


if __name__ == "__main__":
    main()
