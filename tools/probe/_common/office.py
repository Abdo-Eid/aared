"""Shared helpers for ground-truth probes (tools/probe/<topic>/).

The COM application guard, the fixture temp-copy, and the unzip snapshot
live here once instead of copied into every probe script. Windows only:
every helper needs pywin32 and an installed Office application.
"""
import contextlib
import os
import shutil
import tempfile
import zipfile

import win32com.client


@contextlib.contextmanager
def word_app():
    """Hidden Word Application that never quits an instance someone else had open."""
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    had_open = app.Documents.Count > 0
    try:
        yield app
    finally:
        for open_doc in list(app.Documents):
            if not had_open:
                open_doc.Close(0)
        if not had_open:
            app.Quit(0)


@contextlib.contextmanager
def powerpoint_app():
    """Hidden PowerPoint Application with the same never-quit-another's rule."""
    app = win32com.client.DispatchEx("PowerPoint.Application")
    had_open = app.Presentations.Count > 0
    try:
        yield app
    finally:
        if not had_open and app.Presentations.Count == 0:
            app.Quit()


@contextlib.contextmanager
def excel_app():
    """Hidden Excel Application: no alerts, no events, no link updates."""
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.EnableEvents = False
    app.AskToUpdateLinks = False
    had_open = app.Workbooks.Count > 0
    try:
        yield app
    finally:
        for book in list(app.Workbooks):
            if not had_open:
                book.Close(False)
        if not had_open:
            app.Quit()


@contextlib.contextmanager
def fixture_copy(fixture_path):
    """Work on a temp copy of a fixture; copy back over the fixture on clean exit."""
    work = os.path.join(tempfile.mkdtemp(), os.path.basename(fixture_path))
    shutil.copy(fixture_path, work)
    yield work
    shutil.copy(work, fixture_path)


def snapshot(package_path, out_dir):
    """Unzip every part of an OOXML package into out_dir: the probe snapshot."""
    with zipfile.ZipFile(package_path) as src:
        src.extractall(out_dir)
