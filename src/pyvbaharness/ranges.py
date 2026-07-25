"""Pure range helpers.

Kept out of the worker so the write-chunking rule, which exists to dodge a
measured Excel wedge, is unit tested without Excel.
"""
from __future__ import annotations

from typing import Iterator, NamedTuple


class WriteChunk(NamedTuple):
    """Half-open block indices into the caller's 2-D data."""

    row_start: int
    row_end: int
    column_start: int
    column_end: int

    @property
    def cells(self) -> int:
        return ((self.row_end - self.row_start)
                * (self.column_end - self.column_start))


def plan_write_chunks(height: int, width: int,
                      max_cells: int) -> Iterator[WriteChunk]:
    """Split a height x width block into pieces of at most ``max_cells``.

    Rows are split first because a wide single row still has to fit, so the
    column step is capped as well; a block wider than ``max_cells`` is split
    on both axes.
    """
    if height <= 0 or width <= 0:
        return
    if max_cells <= 0:
        raise ValueError("max_cells must be positive")
    column_step = min(width, max_cells)
    row_step = max(1, max_cells // column_step)
    for row_start in range(0, height, row_step):
        row_end = min(row_start + row_step, height)
        for column_start in range(0, width, column_step):
            column_end = min(column_start + column_step, width)
            yield WriteChunk(row_start, row_end, column_start, column_end)


def validate_block(data: list[list[object]]) -> int:
    """Return the block width, or raise for a ragged or empty block."""
    if not data or not all(isinstance(row, list) and row for row in data):
        raise ValueError("Expected a non-empty 2-D list of rows.")
    width = len(data[0])
    if any(len(row) != width for row in data):
        raise ValueError("All rows must have the same width.")
    return width
