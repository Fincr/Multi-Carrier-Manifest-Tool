"""
Finding the carrier sheets in a folder the operator points at.

Batch processing and batch printing both start by asking "which files in this
folder are workbooks?", and they used to answer it with their own inline scan.
Keeping the answer here means the two cannot drift over details like Excel's
lock files, the way the printed-locally carrier list drifted before
`carrier_workflow` collected it.

What the two paths do next differs: processing goes on to read cell B3 and
reject anything that is not a known carrier, while printing takes every
workbook as-is, because the operator asked for the folder to be printed.
"""

import os


_WORKBOOK_EXTENSIONS = ('.xlsx', '.xls')

# Excel writes a "~$name.xlsx" owner file beside any workbook that is open.
# It holds a username, not paperwork, and Excel refuses to open it.
_LOCK_FILE_PREFIX = '~$'


def find_printable_sheets(folder) -> list[str]:
    """
    Full paths of the workbooks directly inside `folder`, alphabetically.

    Subfolders are ignored: the operator chose one folder, and a nested
    archive of last month's sheets must not reach the printer.

    A missing or unreadable folder yields an empty list rather than raising.
    The caller reports "no sheets found" either way, so it has no use for the
    distinction, and this keeps a mistyped path from surfacing as a traceback.
    """
    try:
        names = sorted(os.listdir(folder))
    except OSError:
        return []

    sheets = []
    for name in names:
        if name.startswith(_LOCK_FILE_PREFIX):
            continue
        if not name.lower().endswith(_WORKBOOK_EXTENSIONS):
            continue
        path = os.path.join(folder, name)
        # A folder called "sheets.xlsx" is not a workbook.
        if not os.path.isfile(path):
            continue
        sheets.append(path)

    return sheets
