"""Refresh Word fields in a generated DOCX without an extra Python dependency."""
from dataclasses import asdict, dataclass
from pathlib import Path
import os
import shutil
import socket
import subprocess


@dataclass(frozen=True)
class FieldUpdateReport:
    status: str
    engine: str | None
    messages: tuple[str, ...]

    def to_dict(self):
        data = asdict(self)
        data["messages"] = list(self.messages)
        return data


def update_docx_fields(docx_path, powershell_path=None, work_dir=None,
                       soffice_path=None, libreoffice_python_path=None):
    """Update TOC/PAGE fields, preferring bundled LibreOffice UNO on Windows."""
    if os.name != "nt":
        return FieldUpdateReport(
            "not_run", None, ("Microsoft Word field update is available only on Windows",),
        )
    if powershell_path is None:
        tools = _libreoffice_tools(soffice_path, libreoffice_python_path)
        if tools:
            root = Path(work_dir or Path(docx_path).resolve().parent / ".ni-memo-field-update")
            return _update_with_libreoffice(Path(docx_path).resolve(), root, *tools)
    powershell = powershell_path or shutil.which("powershell.exe") or shutil.which("powershell")
    if not powershell:
        return FieldUpdateReport("not_run", None, ("Windows PowerShell was not found",))

    source = str(Path(docx_path).resolve())
    script = r"""
& {
  param([string]$docxPath)
  $word = $null
  $document = $null
  try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($docxPath, $false, $false)
    for ($index = 1; $index -le $document.TablesOfContents.Count; $index++) {
      $document.TablesOfContents.Item($index).Update()
    }
    $null = $document.Fields.Update()
    $document.Save()
    Write-Output 'WORD_FIELDS_UPDATED'
  } finally {
    if ($null -ne $document) { $document.Close($false) }
    if ($null -ne $word) { $word.Quit() }
  }
}
""".strip()
    timeout_seconds = 120
    try:
        completed = subprocess.run(
            [str(powershell), "-NoProfile", "-NonInteractive", "-Command", script, source],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=timeout_seconds, check=False,
        )
    except subprocess.TimeoutExpired:
        return FieldUpdateReport(
            "fail", "Microsoft Word COM",
            (f"Microsoft Word field update timed out after {timeout_seconds} seconds",),
        )
    messages = tuple(
        line for line in ((completed.stdout or "") + "\n" + (completed.stderr or "")).splitlines()
        if line.strip()
    )
    status = "pass" if completed.returncode == 0 and "WORD_FIELDS_UPDATED" in messages else "fail"
    return FieldUpdateReport(status, "Microsoft Word COM", messages)


def _libreoffice_tools(soffice_path=None, python_path=None):
    candidates = (
        Path(soffice_path) if soffice_path else None,
        Path(shutil.which("soffice.com")) if shutil.which("soffice.com") else None,
        Path(r"C:\Program Files\LibreOffice\program\soffice.com"),
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
    )
    soffice = next((path for path in candidates if path and path.is_file()), None)
    if soffice is None:
        return None
    python = Path(python_path) if python_path else soffice.parent / "python.exe"
    return (soffice, python) if python.is_file() else None


def _update_with_libreoffice(source, work_dir, soffice, python):
    work_dir.mkdir(parents=True, exist_ok=True)
    profile = work_dir / "profile"
    profile.mkdir(exist_ok=True)
    with socket.socket() as probe:
        probe.bind(("127.0.0.1", 0))
        port = probe.getsockname()[1]
    accept = f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    process = subprocess.Popen(
        [
            str(soffice), "--headless", "--norestore", "--nodefault",
            f"-env:UserInstallation={profile.as_uri()}", accept,
        ],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        creationflags=creationflags,
    )
    try:
        completed = subprocess.run(
            [str(python), "-c", _UNO_UPDATE_SCRIPT, str(port), str(source)],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=120, check=False, creationflags=creationflags,
        )
    except subprocess.TimeoutExpired:
        return FieldUpdateReport(
            "fail", "LibreOffice UNO",
            ("LibreOffice field update timed out after 120 seconds",),
        )
    finally:
        if process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=10)
    messages = tuple(
        line for line in ((completed.stdout or "") + "\n" + (completed.stderr or "")).splitlines()
        if line.strip()
    )
    status = (
        "pass"
        if completed.returncode == 0 and any("LIBREOFFICE_FIELDS_UPDATED" in line for line in messages)
        else "fail"
    )
    return FieldUpdateReport(status, "LibreOffice UNO", messages)


_UNO_UPDATE_SCRIPT = r'''
import sys
import time
import uno
from com.sun.star.beans import PropertyValue

def prop(name, value):
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item

port = int(sys.argv[1])
path = sys.argv[2]
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", local
)
context = None
for _ in range(80):
    try:
        context = resolver.resolve(
            f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
        )
        break
    except Exception:
        time.sleep(0.25)
if context is None:
    raise RuntimeError("LibreOffice UNO listener did not start")
desktop = context.ServiceManager.createInstanceWithContext(
    "com.sun.star.frame.Desktop", context
)
document = desktop.loadComponentFromURL(
    uno.systemPathToFileUrl(path), "_blank", 0,
    (prop("Hidden", True), prop("ReadOnly", False), prop("UpdateDocMode", 3)),
)
if document is None:
    raise RuntimeError("LibreOffice could not open DOCX")
try:
    indexes = document.getDocumentIndexes()
    for index in range(indexes.getCount()):
        indexes.getByIndex(index).update()
    document.getTextFields().refresh()
    document.store()
    print(f"LIBREOFFICE_FIELDS_UPDATED indexes={indexes.getCount()}")
finally:
    document.close(True)
    desktop.terminate()
'''.strip()
