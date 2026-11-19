"""Publish a DOCX and its reports as one recoverable batch."""
import hashlib
import json
import os
from pathlib import Path
import tempfile
from contextlib import contextmanager


def _result_names(document):
    return [document.name, document.stem + "_conflicts.txt",
            document.stem + "_merged_changes_log.txt", document.stem + "_text_correction.txt"]


@contextmanager
def _publication_lock(key):
    # OS locks disappear when the process exits. Keeping the inode avoids an
    # unlink/reopen race between publishers using the same output filename.
    lock = Path(tempfile.gettempdir()) / ("mddm-output-" + key + ".lock")
    with lock.open("a+b") as handle:
        handle.seek(0)
        try:
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as error:
            raise OSError("另一项合并正在写入相同的输出文件，请稍后重试。") from error
        yield


def _finish_previous(journal, destination, names):
    """Restore an incomplete publication or discard backups of a committed one."""
    active = journal / "active.json"
    committed = journal / "committed.json"
    state = committed if committed.exists() else active
    if state.exists():
        existed = json.loads(state.read_text(encoding="utf-8"))
        if len(existed) not in (3, 4) or not all(type(flag) is bool for flag in existed):
            raise ValueError(f"无效的输出恢复记录: {state}")
        for index, had_previous in enumerate(existed):
            backup = journal / f"{index}.bak"
            target = destination / names[index]
            if state == committed:
                if os.path.lexists(backup):
                    backup.unlink()
            elif os.path.lexists(backup):
                os.replace(backup, target)
            elif not had_previous and target.exists():
                target.unlink()
        state.unlink()
    draft = journal / "manifest.tmp"
    if draft.exists():
        draft.unlink()
    journal.rmdir()


def publish_outputs(staged_document, destination_document):
    staged = Path(staged_document)
    destination = Path(destination_document)
    names = _result_names(destination)
    entry = destination.parent.resolve() / destination.name
    key = hashlib.sha256(os.fsencode(os.path.normcase(entry))).hexdigest()
    journal = destination.parent / (".mddm-" + key[:20] + ".pending")
    with _publication_lock(key):
        if journal.is_symlink():
            raise OSError(f"输出恢复目录为符号链接: {journal}")
        if journal.exists():
            _finish_previous(journal, destination.parent, names)
        files = names if (staged.parent / names[3]).exists() else names[:3]
        for name in files:
            if not (staged.parent / name).is_file():
                raise FileNotFoundError(f"本次合并未生成所需文件: {name}")
            target = destination.parent / name
            if target.exists() and not target.is_file():
                raise OSError(f"输出位置不是文件: {target}")
        existed = [os.path.lexists(destination.parent / name) for name in files]
        journal.mkdir()
        try:
            draft = journal / "manifest.tmp"
            draft.write_text(json.dumps(existed), encoding="utf-8")
            os.replace(draft, journal / "active.json")
            # Reserve all previous paths first. Windows refuses a rename when a
            # reader (e.g. Word) has an old file open without delete sharing.
            for index, (name, had_previous) in enumerate(zip(files, existed)):
                if had_previous:
                    os.replace(destination.parent / name, journal / f"{index}.bak")
            for name in files:
                os.replace(staged.parent / name, destination.parent / name)
            os.replace(journal / "active.json", journal / "committed.json")
        except BaseException:
            _finish_previous(journal, destination.parent, names)
            raise
        _finish_previous(journal, destination.parent, names)
