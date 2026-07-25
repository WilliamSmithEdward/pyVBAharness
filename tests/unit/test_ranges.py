import pytest

from pyvbaharness.ranges import plan_write_chunks, validate_block


def chunks(height, width, cap):
    return list(plan_write_chunks(height, width, cap))


class TestPlanning:
    def test_small_block_is_one_chunk(self):
        assert chunks(2, 2, 2000) == [(0, 2, 0, 2)]

    def test_every_chunk_is_within_the_cap(self):
        for height, width in ((100, 100), (1000, 3), (7, 5000), (999, 17)):
            planned = chunks(height, width, 2000)
            assert planned
            assert all(c.cells <= 2000 for c in planned)

    def test_chunks_tile_the_block_exactly(self):
        height, width = 37, 41
        covered = set()
        for chunk in chunks(height, width, 200):
            for row in range(chunk.row_start, chunk.row_end):
                for column in range(chunk.column_start, chunk.column_end):
                    key = (row, column)
                    assert key not in covered  # no overlap
                    covered.add(key)
        assert len(covered) == height * width  # no gaps

    def test_wide_block_splits_columns_too(self):
        planned = chunks(3, 5000, 2000)
        assert max(c.column_end - c.column_start for c in planned) == 2000
        assert all(c.cells <= 2000 for c in planned)

    def test_single_row_wider_than_cap(self):
        planned = chunks(1, 4500, 2000)
        assert [(c.column_start, c.column_end) for c in planned] == [
            (0, 2000), (2000, 4000), (4000, 4500)]

    def test_empty_block_yields_nothing(self):
        assert chunks(0, 5, 100) == []
        assert chunks(5, 0, 100) == []

    def test_invalid_cap_rejected(self):
        with pytest.raises(ValueError):
            chunks(2, 2, 0)

    def test_10k_block_matches_measured_shape(self):
        # 100x100 at the shipped cap: 20-row chunks, five of them.
        planned = chunks(100, 100, 2000)
        assert len(planned) == 5
        assert all(c.row_end - c.row_start == 20 for c in planned)


class TestValidation:
    def test_returns_width(self):
        assert validate_block([[1, 2, 3], [4, 5, 6]]) == 3

    def test_rejects_empty_and_ragged(self):
        for bad in ([], [[]], [[1, 2], [3]], [1, 2], [None]):
            with pytest.raises(ValueError):
                validate_block(bad)
