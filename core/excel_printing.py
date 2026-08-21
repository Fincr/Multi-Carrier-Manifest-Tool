"""
What to do with an open Excel workbook on its way to the printer.

Opening and closing workbooks over COM stays in the GUI layer; the decisions —
which sheets to lay out, which object to print, which printer to name — live
here so they can be tested against a stand-in workbook instead of a real Excel
instance and a real tray of paper.
"""


# The largest value FitToPagesTall accepts. Excel has no "unlimited" here, so
# this is the idiomatic way to say: constrain the width, leave the height be.
_UNCONSTRAINED_PAGES_TALL = 32767


def fit_columns_to_one_page(sheet) -> None:
    """
    Lay a sheet out so no column runs off the side of the page.

    Rows are deliberately left free to spill onto further pages. Forcing a long
    sheet onto one page instead shrinks it until nobody can read it.

    Chart sheets have no usable PageSetup and raise; a sheet that cannot be
    laid out is skipped rather than being allowed to abort the print.
    """
    try:
        # Must come first: while Zoom holds a number, Excel ignores
        # FitToPages* entirely. Setting it False is what enables the fit.
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = _UNCONSTRAINED_PAGES_TALL
    except Exception:
        pass


def print_workbook(wb, printer_name: str = None, active_sheet_only: bool = False) -> None:
    """
    Send an open workbook to the printer, laying out its columns first.

    By default the whole workbook prints, every tab: that is what a populated
    manifest template needs. With `active_sheet_only`, just the sheet Excel
    opened on prints, which is what a carrier sheet needs — its Notes and Lists
    tabs are reference material, not paperwork.

    `printer_name` of None means the Windows default printer. It is omitted
    from the call rather than passed as None, which Excel would reject.
    """
    target = wb.ActiveSheet if active_sheet_only else wb
    sheets = [wb.ActiveSheet] if active_sheet_only else list(wb.Sheets)

    for sheet in sheets:
        fit_columns_to_one_page(sheet)

    if printer_name:
        target.PrintOut(ActivePrinter=printer_name)
    else:
        target.PrintOut()


def should_abandon_batch_print(failure_count: int, attempted: int) -> bool:
    """
    Whether a failed sheet means giving up on the rest of the batch.

    Asked only when a sheet has just failed and sheets remain.

    The dividing line is whether anything has printed yet. If nothing has, the
    printer is the likely culprit — jammed, offline, or a bad queue name — and
    the remaining sheets would each pay a slow Excel round-trip only to fail
    the same way. Once one sheet has come out the queue is demonstrably alive,
    so a later failure belongs to that file, and abandoning the sheets behind
    it would force the operator to re-run the folder and reprint the pages that
    already succeeded.

    Args:
        failure_count: Sheets that have failed so far, including this one.
        attempted: Sheets tried so far, including this one.

    Returns:
        True to stop the batch, False to keep printing.
    """
    return failure_count == attempted
