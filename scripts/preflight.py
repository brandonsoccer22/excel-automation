"""Check Windows Python access, xlwings initialization, and isolated Excel COM."""

import importlib
import os
import sys


def check_directory(path):
    original_directory = os.getcwd()
    try:
        os.chdir(path)
    finally:
        if os.getcwd() != original_directory:
            os.chdir(original_directory)


def main():
    print(f"Python executable: {sys.executable}", flush=True)
    print(f"Python exec_prefix: {sys.exec_prefix}", flush=True)
    print(f"Working directory: {os.getcwd()}", flush=True)
    print(f"Preflight script: {os.path.abspath(__file__)}", flush=True)
    if sys.platform != "win32":
        print("Windows Python and desktop Excel are required for COM.", file=sys.stderr)
        return 2

    stage = "access to Python exec_prefix"
    try:
        check_directory(sys.exec_prefix)
        print("PASS: can change to Python exec_prefix and return", flush=True)

        stage = "import win32com.client"
        client = importlib.import_module("win32com.client")
        print(f"PASS: {stage} ({client.__file__})", flush=True)

        stage = "import and initialize xlwings Windows backend"
        xw = importlib.import_module("xlwings")
        print(f"PASS: {stage} ({xw.__file__})", flush=True)

        stage = "start isolated native Excel COM instance"
        excel = client.DispatchEx("Excel.Application")
        try:
            stage = "read native Excel COM version"
            print(f"PASS: Excel COM version {excel.Version}", flush=True)
        finally:
            if sys.exc_info()[0] is None:
                stage = "quit isolated native Excel COM instance"
            excel.Quit()

        stage = "start isolated xlwings Excel instance"
        app = xw.App(visible=False, add_book=False)
        try:
            stage = "read Excel version through xlwings"
            print(f"PASS: xlwings Excel version {app.api.Version}", flush=True)
        finally:
            if sys.exc_info()[0] is None:
                stage = "quit isolated xlwings Excel instance"
            app.quit()
    except Exception as exc:
        print(f"FAIL: {stage}: {type(exc).__name__}: {exc}", file=sys.stderr)
        if hasattr(exc, "hresult"):
            print(f"HRESULT: 0x{exc.hresult & 0xFFFFFFFF:08X}", file=sys.stderr)
        raise

    return 0


if __name__ == "__main__":
    sys.exit(main())
