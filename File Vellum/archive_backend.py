"""Secure ZIP/TAR archive operations for FileVellum.

All extracted content is staged in private temporary directories. Archive members
are validated before writing, and output files use unpredictable names.
"""
import io
import json
import os
import shutil
import stat
import tarfile
import tempfile
import zipfile
from pathlib import PurePosixPath
from uuid import uuid4

MAX_UPLOAD_BYTES = 512 * 1024 * 1024
MAX_ENTRIES = 5000
MAX_EXPANDED_BYTES = 2 * 1024 * 1024 * 1024
MAX_RATIO = 200
CHUNK = 1024 * 1024


def _clean_name(name):
    name = (name or '').replace('\\', '/')
    p = PurePosixPath(name)
    if not name or name.startswith('/') or any(part in ('..', '') for part in p.parts):
        raise ValueError('Archive contains an unsafe path.')
    if p.parts and ':' in p.parts[0]:
        raise ValueError('Archive contains an unsafe path.')
    return str(p)


def _save_uploads(files, folder):
    good = [f for f in files if f and f.filename]
    if not good:
        raise ValueError('Please select at least one file.')
    total = 0
    saved = []
    for f in good:
        name = os.path.basename(f.filename.replace('\\', '/'))
        if not name or name in ('.', '..'):
            raise ValueError('Invalid upload filename.')
        path = os.path.join(folder, f'{uuid4().hex}_{name}')
        size = 0
        with open(path, 'wb') as out:
            while True:
                chunk = f.stream.read(CHUNK)
                if not chunk:
                    break
                size += len(chunk); total += len(chunk)
                if size > MAX_UPLOAD_BYTES or total > MAX_UPLOAD_BYTES:
                    raise ValueError('Uploads are limited to 512 MB total.')
                out.write(chunk)
        saved.append((path, name))
    return saved


def _output_path(upload_folder, suffix):
    return os.path.join(upload_folder, f'filevellum_{uuid4().hex}{suffix}')


def _write_zip_from_files(items, output, password=None):
    if password:
        try:
            import pyzipper
        except ImportError as exc:
            raise RuntimeError('AES ZIP support requires pyzipper. Install requirements.txt.') from exc
        with pyzipper.AESZipFile(output, 'w', compression=pyzipper.ZIP_DEFLATED,
                                 encryption=pyzipper.WZ_AES) as zf:
            zf.setpassword(password.encode('utf-8'))
            zf.setencryption(pyzipper.WZ_AES, nbits=256)
            for source, arcname in items:
                zf.write(source, arcname)
    else:
        with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
            for source, arcname in items:
                zf.write(source, arcname)


def _extract_zip(path, target, password=None):
    try:
        # pyzipper reads AES and standard ZIP encryption. Fallback handles normal ZIPs.
        try:
            import pyzipper
            zf = pyzipper.AESZipFile(path, 'r')
        except ImportError:
            zf = zipfile.ZipFile(path, 'r')
        with zf:
            infos = zf.infolist()
            if len(infos) > MAX_ENTRIES:
                raise ValueError('Archive contains too many entries.')
            expanded = 0
            for info in infos:
                name = _clean_name(info.filename)
                mode = (info.external_attr >> 16) & 0xFFFF
                if stat.S_ISLNK(mode):
                    raise ValueError('ZIP symbolic links are not allowed.')
                if info.flag_bits & 1 and not password:
                    raise ValueError('This ZIP is password-protected. Enter its password.')
                expanded += info.file_size
                if expanded > MAX_EXPANDED_BYTES:
                    raise ValueError('Archive expands beyond the 2 GB safety limit.')
                if info.compress_size and info.file_size / info.compress_size > MAX_RATIO:
                    raise ValueError('Archive compression ratio exceeds the safety limit.')
                dest = os.path.abspath(os.path.join(target, *PurePosixPath(name).parts))
                if os.path.commonpath([os.path.abspath(target), dest]) != os.path.abspath(target):
                    raise ValueError('Archive contains an unsafe path.')
                if info.is_dir():
                    os.makedirs(dest, exist_ok=True); continue
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                try:
                    with zf.open(info, 'r', pwd=password.encode('utf-8') if password else None) as src, open(dest, 'wb') as out:
                        copied = 0
                        while True:
                            b = src.read(CHUNK)
                            if not b: break
                            copied += len(b)
                            if copied > info.file_size or copied > MAX_EXPANDED_BYTES:
                                raise ValueError('Invalid or oversized ZIP member.')
                            out.write(b)
                except RuntimeError as exc:
                    raise ValueError('Incorrect password or invalid encrypted ZIP.') from exc
                except zipfile.BadZipFile as exc:
                    raise ValueError('ZIP archive is corrupted.') from exc
    except (zipfile.BadZipFile, tarfile.TarError, OSError) as exc:
        raise ValueError('Invalid or corrupted ZIP archive.') from exc


def _extract_tar(path, target, mode):
    try:
        with tarfile.open(path, mode) as tf:
            members = tf.getmembers()
            if len(members) > MAX_ENTRIES:
                raise ValueError('Archive contains too many entries.')
            expanded = 0
            for member in members:
                name = _clean_name(member.name)
                if not (member.isfile() or member.isdir()):
                    raise ValueError('TAR links, devices, FIFOs and special entries are not allowed.')
                expanded += member.size
                if expanded > MAX_EXPANDED_BYTES:
                    raise ValueError('Archive expands beyond the 2 GB safety limit.')
                dest = os.path.abspath(os.path.join(target, *PurePosixPath(name).parts))
                if os.path.commonpath([os.path.abspath(target), dest]) != os.path.abspath(target):
                    raise ValueError('Archive contains an unsafe path.')
                if member.isdir():
                    os.makedirs(dest, exist_ok=True); continue
                src = tf.extractfile(member)
                if src is None: raise ValueError('Invalid TAR member.')
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                with src, open(dest, 'wb') as out:
                    shutil.copyfileobj(src, out, CHUNK)
    except (tarfile.TarError, OSError) as exc:
        raise ValueError('Invalid or corrupted TAR archive.') from exc


def _unique_upload_items(uploads):
    used = set()
    items = []
    for path, original in uploads:
        base, ext = os.path.splitext(original)
        candidate = original
        index = 2
        while candidate.casefold() in used:
            candidate = f'{base}_{index}{ext}'
            index += 1
        used.add(candidate.casefold())
        items.append((path, candidate))
    return items


def _folder_files(root):
    items = []
    for base, dirs, files in os.walk(root):
        dirs.sort(); files.sort()
        for filename in files:
            full = os.path.join(base, filename)
            rel = os.path.relpath(full, root).replace(os.sep, '/')
            items.append((full, rel))
    return items


def process_archive(tool, files, form, upload_folder):
    out = None
    try:
        with tempfile.TemporaryDirectory(prefix='fv_archive_') as work:
            uploads = _save_uploads(files, work)
            password = (form.get('password') or '').strip()
            if tool in ('password_protect_zip', 'extract_password_protected_zip') and not password:
                raise ValueError('Please enter the ZIP password.')
            out = None

            if tool in ('create_zip_archive', 'compress_files_to_zip', 'compressor'):
                out = _output_path(upload_folder, '.zip')
                _write_zip_from_files(_unique_upload_items(uploads), out)
            elif tool == 'merge_zip_archives':
                if len(uploads) < 2: raise ValueError('Select at least two ZIP archives to merge.')
                merged = os.path.join(work, 'merged')
                os.makedirs(merged)
                for index, (path, _) in enumerate(uploads, 1):
                    if not zipfile.is_zipfile(path): raise ValueError('Every uploaded file must be a valid ZIP archive.')
                    stage = os.path.join(merged, f'archive_{index}')
                    os.makedirs(stage)
                    _extract_zip(path, stage)
                out = _output_path(upload_folder, '.zip')
                _write_zip_from_files(_folder_files(merged), out)
            elif tool in ('extract_zip_archive', 'extract_multiple_files', 'extract_zip', 'extract_password_protected_zip'):
                if len(uploads) != 1: raise ValueError('Select exactly one ZIP archive.')
                target = os.path.join(work, 'extracted'); os.makedirs(target)
                _extract_zip(uploads[0][0], target, password or None)
                out = _output_path(upload_folder, '_extracted.zip')
                _write_zip_from_files(_folder_files(target), out)
            elif tool == 'password_protect_zip':
                out = _output_path(upload_folder, '_encrypted.zip')
                _write_zip_from_files(_unique_upload_items(uploads), out, password)
            elif tool == 'split_large_files_into_parts':
                if len(uploads) != 1: raise ValueError('Select exactly one file to split.')
                try: part_mb = int(form.get('part_size_mb', '10'))
                except ValueError: raise ValueError('Part size must be a number.')
                if not 1 <= part_mb <= 100: raise ValueError('Part size must be between 1 and 100 MB.')
                size = part_mb * 1024 * 1024
                src, original = uploads[0]
                parts = []
                with open(src, 'rb') as f:
                    idx = 1
                    while True:
                        data = f.read(size)
                        if not data: break
                        p = os.path.join(work, f'{original}.part{idx:04d}')
                        with open(p, 'wb') as w: w.write(data)
                        parts.append((p, os.path.basename(p))); idx += 1
                out = _output_path(upload_folder, '_parts.zip')
                _write_zip_from_files(parts, out)
            elif tool in ('create_tar_archive', 'create_tar_gz_archive'):
                suffix = '.tar' if tool == 'create_tar_archive' else '.tar.gz'
                out = _output_path(upload_folder, suffix)
                mode = 'w' if suffix == '.tar' else 'w:gz'
                with tarfile.open(out, mode) as tf:
                    for p, n in _unique_upload_items(uploads): tf.add(p, arcname=n, recursive=False)
            elif tool in ('extract_tar_archive', 'extract_tar_gz_archive'):
                if len(uploads) != 1: raise ValueError('Select exactly one TAR archive.')
                target = os.path.join(work, 'extracted'); os.makedirs(target)
                mode = 'r:gz' if tool == 'extract_tar_gz_archive' else 'r:*'
                _extract_tar(uploads[0][0], target, mode)
                out = _output_path(upload_folder, '_extracted.zip')
                _write_zip_from_files(_folder_files(target), out)
            elif tool == 'zip_file_information':
                if len(uploads) != 1: raise ValueError('Select exactly one ZIP archive.')
                report_path = os.path.join(work, 'zip_info.json')
                with zipfile.ZipFile(uploads[0][0], 'r') as zf:
                    infos = zf.infolist()
                    if len(infos) > MAX_ENTRIES: raise ValueError('Archive contains too many entries.')
                    if sum(i.file_size for i in infos) > MAX_EXPANDED_BYTES:
                        raise ValueError('Archive expands beyond the 2 GB safety limit.')
                    if any(i.compress_size and i.file_size / i.compress_size > MAX_RATIO for i in infos):
                        raise ValueError('Archive compression ratio exceeds the safety limit.')
                    if any(i.flag_bits & 1 for i in infos):
                        bad = 'encrypted members require a password to verify'
                    else:
                        bad = zf.testzip()
                        if bad: raise ValueError(f'ZIP archive is corrupted (member: {bad}).')
                    report = {'file': uploads[0][1], 'entries': len(infos), 'corrupt_member': bad,
                              'total_uncompressed_bytes': sum(i.file_size for i in infos),
                              'files': [{'name': i.filename, 'size_bytes': i.file_size,
                                         'compressed_bytes': i.compress_size, 'is_directory': i.is_dir(),
                                         'encrypted': bool(i.flag_bits & 1)} for i in infos]}
                with open(report_path, 'w', encoding='utf-8') as f: json.dump(report, f, indent=2)
                out = _output_path(upload_folder, '_information.json'); shutil.copyfile(report_path, out)
            elif tool in ('zip_to_tar', 'tar_to_zip'):
                if len(uploads) != 1: raise ValueError('Select exactly one archive to convert.')
                target = os.path.join(work, 'converted'); os.makedirs(target)
                if tool == 'zip_to_tar':
                    _extract_zip(uploads[0][0], target)
                    out = _output_path(upload_folder, '.tar')
                    with tarfile.open(out, 'w') as tf:
                        for p, n in _folder_files(target): tf.add(p, arcname=n, recursive=False)
                else:
                    _extract_tar(uploads[0][0], target, 'r:*')
                    out = _output_path(upload_folder, '.zip')
                    _write_zip_from_files(_folder_files(target), out)
            else:
                raise ValueError('Unknown ZIP & Archives operation.')
            return {'success': True, 'filename': os.path.basename(out)}
    except ValueError as exc:
        if out and os.path.isfile(out):
            try: os.remove(out)
            except OSError: pass
        return {'success': False, 'error': str(exc)}, 400
    except RuntimeError as exc:
        if out and os.path.isfile(out):
            try: os.remove(out)
            except OSError: pass
        return {'success': False, 'error': str(exc)}, 500
    except Exception:
        if out and os.path.isfile(out):
            try: os.remove(out)
            except OSError: pass
        return {'success': False, 'error': 'Archive processing failed. Verify the files and try again.'}, 500
