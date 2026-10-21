"""PowerPoint-only native backend for verify_office_macos.py.

Uses indexed document access (collection object iteration can hang), ordinary
Launch Services open and HFS export paths. Never saves the source or interacts
with recovery prompts. Receipts and app-specific temporary files stay outside
this reusable package.
"""
import hashlib
from pathlib import Path
import shutil
import subprocess
import tempfile
import time

LOOKUP = '''on normalizePath(fileName)
 if fileName starts with "/" then return fileName
 return POSIX path of (fileName as alias)
end normalizePath
on run argv
 set targetPath to item 1 of argv
 tell application "Microsoft PowerPoint"
  repeat with i from 1 to count of presentations
   if my normalizePath(full name of presentation i as text) is targetPath then
    return ((count of slides of presentation i) as text) & "|" & version
   end if
  end repeat
 end tell
 return "NONE"
end run'''
EXPORT = '''on normalizePath(fileName)
 if fileName starts with "/" then return fileName
 return POSIX path of (fileName as alias)
end normalizePath
on run argv
 set targetPath to item 1 of argv
 set outHFS to (POSIX file (item 2 of argv)) as text
 tell application "Microsoft PowerPoint"
  repeat with i from 1 to count of presentations
   if my normalizePath(full name of presentation i as text) is targetPath then
    set taskDeck to presentation i
    if saved of taskDeck is false then error "Target has unsaved changes; preserve it"
    save taskDeck in outHFS as save as PDF
    return "PPT_EXPORT_OK"
   end if
  end repeat
 end tell
 error "Target presentation missing"
end run'''
CLOSE = '''on normalizePath(fileName)
 if fileName starts with "/" then return fileName
 return POSIX path of (fileName as alias)
end normalizePath
on run argv
 set targetPath to item 1 of argv
 tell application "Microsoft PowerPoint"
  repeat with i from (count of presentations) to 1 by -1
   if my normalizePath(full name of presentation i as text) is targetPath then
    if saved of presentation i is false then error "Target has unsaved changes; left open"
    close presentation i saving no
   end if
  end repeat
 end tell
end run'''
# Optional handler for two known localisations. No generic dialog acceptance.
# The only allowed target is the private export folder just made by this call.
GRANT = '''on run argv
 set expectedFolder to item 1 of argv
 set expectedPath to item 2 of argv
 tell application "System Events"
  tell process "Microsoft PowerPoint"
   set targetWindow to missing value
   repeat with dialogName in {"Grant File Access", "授予文件访问权限"}
    if exists window (dialogName as text) then set targetWindow to window (dialogName as text)
   end repeat
   if targetWindow is missing value then return "NO_MATCHING_PROMPT"
   set w to targetWindow
   if not (exists sheet 1 of w) then
    set exactTarget to false
    set promptItems to entire contents of w
    repeat with e in promptItems
     try
      if (value of e as text) is expectedPath then set exactTarget to true
     end try
    end repeat
    if not exactTarget then return "UNEXPECTED_ACCESS_TARGET"
    repeat with labelText in {"Select...", "Select…", "选择...", "选择…"}
     if exists button (labelText as text) of w then
      click button (labelText as text) of w
      return "SELECTED_TASK_EXPORT_FOLDER"
     end if
    end repeat
   else
    set g to splitter group 1 of sheet 1 of w
    set matchesFolder to false
    repeat with t in static texts of g
     if (value of t as text) contains expectedFolder then set matchesFolder to true
    end repeat
    if not matchesFolder then return "UNEXPECTED_ACCESS_PICKER"
    repeat with labelText in {"Grant Access", "授予访问权限"}
     if exists button (labelText as text) of g then
      click button (labelText as text) of g
      return "GRANTED_TASK_EXPORT_FOLDER"
     end if
    end repeat
   end if
   return "UNRECOGNISED_ACCESS_CONTROLS"
  end tell
 end tell
end run'''


def apple(script, *args, timeout=8):
    run = subprocess.run(['osascript', '-', *map(str, args)], input=script,
                         text=True, capture_output=True, timeout=timeout)
    if run.returncode:
        raise RuntimeError(run.stderr.strip() or 'AppleScript failed')
    return run.stdout.strip()


def complete_pdf(path, expected_pages):
    """Require an exported PDF trailer and a parser-confirmed page count."""
    if not path.is_file():
        return False
    data = path.read_bytes()
    if not data.startswith(b'%PDF-') or not data.rstrip().endswith(b'%%EOF'):
        return False
    run = subprocess.run(['pdfinfo', str(path)], capture_output=True, text=True, timeout=8)
    if run.returncode:
        return False
    counts = [line.split()[-1] for line in run.stdout.splitlines() if line.startswith('Pages:')]
    return counts == [str(expected_pages)]


def start_export(source, native_pdf, folder):
    script = folder / 'export.applescript'
    script.write_text(EXPORT, encoding='utf-8')
    return subprocess.Popen(['osascript', str(script), str(source), str(native_pdf)],
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)


def finish_client(proc):
    try:
        out, err = proc.communicate(timeout=3)
        return {'returncode': proc.returncode, 'stdout': out.strip(), 'stderr': err.strip()}
    except subprocess.TimeoutExpired:
        proc.terminate()  # Only this helper's client; never terminate PowerPoint.
        out, err = proc.communicate(timeout=3)
        return {'returncode': proc.returncode, 'stdout': out.strip(), 'stderr': err.strip(),
                'reply': 'client stopped after bounded wait; assess actual PDF separately'}


def verify_ppt(source, pdf, result, grant_export_dir=False, open_seconds=30, export_seconds=60):
    """Update the caller's receipt; paths/output refusal are checked upstream."""
    if not shutil.which('pdfinfo'):
        raise RuntimeError('PowerPoint export verification requires pdfinfo (Poppler)')
    result.update(native_open='not_performed', native_export='not_performed',
                  opened_by_helper=False, owned_target_left_open=False,
                  events=[], count_unit='slides')
    if apple(LOOKUP, source) != 'NONE':
        raise RuntimeError('Exact target is already open; preserve its editing state. Use a separate candidate or inspect this target manually.')
    subprocess.run(['open', '-a', 'Microsoft PowerPoint', str(source)], check=True, timeout=8)
    result.update(opened_by_helper=True, owned_target_left_open=True)
    deadline = time.monotonic() + open_seconds
    opened = 'NONE'
    while time.monotonic() < deadline:
        try:
            opened = apple(LOOKUP, source)
        except subprocess.TimeoutExpired:
            result['events'].append('Open lookup timed out while application was loading')
            continue
        if opened != 'NONE':
            break
        time.sleep(.5)
    if opened == 'NONE':
        raise RuntimeError('Ordinary open did not yield the exact target before the deadline')
    count, version = opened.split('|', 1)
    count = int(count)
    if count <= 0:
        raise RuntimeError('Opened presentation has no slides')
    result.update(native_open='passed', object_count=count, app_version=version,
                  open_method='Launch Services ordinary open; indexed full-path lookup; no repair/recovery')
    folder = Path(tempfile.mkdtemp(prefix='office-ppt-')).resolve()
    native_pdf = folder / 'export.pdf'
    result['temporary_export'] = str(native_pdf)
    result['native_export'] = 'failed'
    proc = start_export(source, native_pdf, folder)
    complete = False
    previous_size, stable = -1, 0
    try:
        deadline = time.monotonic() + export_seconds
        while time.monotonic() < deadline:
            if native_pdf.is_file():
                size = native_pdf.stat().st_size
                stable = stable + 1 if size == previous_size else 0
                previous_size = size
                if stable >= 2 and complete_pdf(native_pdf, count):
                    complete = True
                    break
            if grant_export_dir:
                try:
                    event = apple(GRANT, folder.name, folder, timeout=5)
                    if event != 'NO_MATCHING_PROMPT':
                        result['events'].append(event)
                    if event.startswith('UNEXPECTED') or event.startswith('UNRECOGNISED'):
                        raise RuntimeError(event + '; no further UI interaction performed')
                except subprocess.TimeoutExpired:
                    result['events'].append('Access prompt read timed out; no success assumed')
            time.sleep(.5)
        if not complete:
            raise RuntimeError('No stable, complete native PDF with the expected page count; inspect the target application/access prompt before any retry')
        if hashlib.sha256(source.read_bytes()).hexdigest() != result['source_sha256']:
            raise RuntimeError('Source bytes changed during native verification')
        with pdf.open('xb') as stream:
            stream.write(native_pdf.read_bytes())
        result.update(native_export='passed', pdf=str(pdf), pdf_sha256=hashlib.sha256(pdf.read_bytes()).hexdigest(),
                      export_method='Single HFS-path native export; stable file, PDF trailer and pdfinfo page count checked; copied unchanged')
    finally:
        result['export_client'] = finish_client(proc)
        if complete and source.is_file() and hashlib.sha256(source.read_bytes()).hexdigest() == result['source_sha256']:
            # PDF completion is established before close, not inferred from the reply.
            apple(CLOSE, source)
            result['owned_target_left_open'] = False
            shutil.rmtree(folder)
        # An incomplete export leaves this helper's target open and temp evidence
        # intact; never repeat save blindly or dismiss an unknown application UI.
