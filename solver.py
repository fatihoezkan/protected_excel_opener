"""Remove worksheet, workbook, and VBA protection from copies of Excel files."""
from pathlib import Path
from datetime import datetime
import hashlib
import json
import os
import re
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

from protection import records, unlock_vba

ROOT = Path(__file__).resolve().parent
# Resolve paths relative to this script, so launching from another folder works.
SUPPORTED = {'.xlsb', '.xlsm', '.xlsx'}

def digest(path):
    # Compare the input hash before and after processing to detect changes.
    # file_digest needs Python 3.11+; START.bat currently selects Python 3.14.
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()

def xml_unprotect(data, tag='sheetProtection'):
    # Rewriting the entire XML tree could change namespaces or formatting.
    # Remove only the protection element and preserve the surrounding bytes.
    # Reject structures where the byte-level match disagrees with the XML parser.
    root = ET.fromstring(data)
    expected = sum(e.tag.rsplit('}', 1)[-1] == tag for e in root)
    tag_bytes = tag.encode('ascii')
    pattern = rb'<(?:[\w.-]+:)?' + tag_bytes + rb'\b[^>]*(?:/>|>\s*</(?:[\w.-]+:)?' + tag_bytes + rb'\s*>)'
    result, count = re.subn(pattern, b'', data)
    if count != expected:
        raise ValueError('Unsupported protection XML structure.')
    ET.fromstring(result)
    return result, count

def process(source, output):
    # Read the source without modifying it; publish the copy only after validation.
    if source.suffix.lower() not in SUPPORTED:
        raise ValueError('Unsupported format. Use .xlsb, .xlsm, or .xlsx files.')
    if not zipfile.is_zipfile(source):
        raise ValueError('Cannot read this file as an Excel package. It may be encrypted or damaged.')
    original_hash = digest(source)
    fd, temporary = tempfile.mkstemp(prefix='.processing-', suffix='.tmp', dir=output)
    os.close(fd)
    temporary = Path(temporary)
    changed, sheet_count, vba = [], 0, False
    book_count = 0
    try:
        with zipfile.ZipFile(source) as zin:
            if len(zin.namelist()) != len(set(zin.namelist())):
                raise ValueError('The package contains duplicate entries.')
            if not any(n in zin.namelist() for n in ('xl/workbook.bin', 'xl/workbook.xml')):
                raise ValueError('No valid Excel workbook was found in the package.')
            if any('signature' in n.lower() for n in zin.namelist()):
                # Changing signed content could invalidate its digital signature.
                raise ValueError('Signed files are not supported because editing them may invalidate the signature.')
            with zipfile.ZipFile(temporary, 'w') as zout:
                zout.comment = zin.comment
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    new = data
                    if item.filename.startswith('xl/worksheets/'):
                        if item.filename.endswith('.bin'):
                            # MS-XLSB: 535 = BrtSheetProtection;
                            # 678 = BrtSheetProtectionIso. Preserve every other record as written.
                            entries = list(records(data))
                            sheet_count += sum(t in (535, 678) for t, _ in entries)
                            new = b''.join(raw for t, raw in entries if t not in (535, 678))
                        elif item.filename.endswith('.xml'):
                            new, count = xml_unprotect(data)
                            sheet_count += count
                    elif item.filename == 'xl/workbook.bin':
                        # Workbook structure protection is separate from worksheet protection:
                        # 534 = BrtBookProtection; 677 = BrtBookProtectionIso.
                        entries = list(records(data))
                        book_count += sum(t in (534, 677) for t, _ in entries)
                        new = b''.join(raw for t, raw in entries if t not in (534, 677))
                    elif item.filename == 'xl/workbook.xml':
                        new, book_count = xml_unprotect(data, 'workbookProtection')
                    elif item.filename == 'xl/vbaProject.bin':
                        # Edit the VBA container directly, without opening Excel or running macros.
                        new = unlock_vba(data)
                        vba = True
                    if new != data:
                        changed.append(item.filename)
                    zout.writestr(item, new)
        # Check ZIP integrity and compare content that should remain unchanged.
        # unlock_vba checks the individual streams inside the VBA container.
        with zipfile.ZipFile(source) as zin, zipfile.ZipFile(temporary) as zout:
            if zout.testzip() is not None or zin.namelist() != zout.namelist():
                raise ValueError('Output package validation failed.')
            for name in zin.namelist():
                old, new = zin.read(name), zout.read(name)
                if name not in changed:
                    assert old == new, 'Unexpected content change'
                elif name.startswith('xl/worksheets/') and name.endswith('.bin'):
                    assert list(records(new)) == [(t, raw) for t, raw in records(old) if t not in (535, 678)]
                elif name.startswith('xl/worksheets/') and name.endswith('.xml'):
                    assert new == xml_unprotect(old)[0]
                elif name == 'xl/workbook.bin':
                    assert list(records(new)) == [(t, raw) for t, raw in records(old) if t not in (534, 677)]
                elif name == 'xl/workbook.xml':
                    assert new == xml_unprotect(old, 'workbookProtection')[0]
        if digest(source) != original_hash:
            raise ValueError('The input changed during processing. No output was published.')
        target = output / source.name
        number = 2
        while True:
            try:
                # On Windows, rename refuses to overwrite an existing destination, even
                # during concurrent runs. POSIX rename does not provide this guarantee.
                temporary.rename(target)
                break
            except FileExistsError:
                target = output / f'{source.stem} ({number}){source.suffix}'
                number += 1
        return {'input': source.name, 'status': 'success', 'output': target.name,
                'sheet_protection_records': sheet_count, 'vba_processed': vba,
                'workbook_protection_records': book_count,
                'modified_parts': changed, 'source_sha256': original_hash}
    finally:
        # Remove unfinished output if any step fails.
        temporary.unlink(missing_ok=True)

def main():
    # Scan the files directly inside input once per run.
    # Ignore the temporary lock files Excel creates with a ~$ prefix.
    incoming, outgoing = ROOT / 'input', ROOT / 'output'
    incoming.mkdir(exist_ok=True)
    outgoing.mkdir(exist_ok=True)
    files = sorted(p for p in incoming.iterdir() if p.is_file() and not p.name.startswith('~$'))
    if not files:
        print('The input folder is empty. Add your Excel files and run the program again.')
        return 0
    report = []
    for source in files:
        print(f'Processing: {source.name}', flush=True)
        try:
            result = process(source, outgoing)
            print(f"  SAVED -> output/{result['output']}", flush=True)
        except Exception as error:
            # A damaged or unsupported file should not stop the rest of the batch.
            message = str(error) or type(error).__name__
            result = {'input': source.name, 'status': 'error', 'error': message}
            print(f'  FAILED: {message}', flush=True)
        report.append(result)
    name = 'run_report_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f') + '.json'
    (outgoing / name).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    success = sum(r['status'] == 'success' for r in report)
    print(f'\nDone. Successful: {success} / {len(report)}. Report: output/{name}')
    return int(success != len(report))

if __name__ == '__main__':
    # Validation uses assertions, so do not allow Python to disable them.
    if sys.flags.optimize:
        raise SystemExit('Do not use Python -O: it disables required validation checks.')
    try:
        raise SystemExit(main())
    except (OSError, ImportError) as error:
        print(f'The program could not complete: {error}')
        raise SystemExit(1)
