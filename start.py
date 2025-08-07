import atexit
import multiprocessing as mp

import webview

from app import app


def runApp():
    app.run(debug=True, port=5000)


if __name__ == "_main_":
    mp.freeze_support()
    procs = []
    proc = mp.Process(target=runApp)
    procs.append(proc)
    proc.start()

    def close():
        proc.terminate()

    atexit.register(close)
    webview.settings = {
        "ALLOW_DOWNLOADS": True,
        "ALLOW_FILE_URLS": True,
        "OPEN_EXTERNAL_LINKS_IN_BROWSER": True,
        "OPEN_DEVTOOLS_IN_DEBUG": True,
        "REMOTE_DEBUGGING_PORT": None,
    }

    window = webview.create_window(
        "PyTransform", "http:// 127.0.0.1:5000", height=720, width=1600, resizable=False
    )
    webview.start(ssl=True)
