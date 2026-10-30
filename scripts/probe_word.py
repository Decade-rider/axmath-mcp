"""Read installed Word/AxMath integration in an isolated Word instance."""

import json

import pythoncom
import win32com.client

pythoncom.CoInitialize()
word = None
try:
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    print(
        json.dumps(
            {
                "version": word.Version,
                "path": word.Path,
                "addins": [
                    {"name": a.Name, "path": a.Path, "installed": a.Installed} for a in word.AddIns
                ],
            },
            ensure_ascii=False,
        ),
        flush=True,
    )
finally:
    if word is not None:
        word.Quit(0)
    pythoncom.CoUninitialize()
