###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import BytesIO

from xlsxwriter.workbook import Workbook


class TestWriteRejectsBytes(unittest.TestCase):
    """bytes/bytearray must raise at write time, not on workbook.close()."""

    def test_write_bytes_raises(self):
        workbook = Workbook(BytesIO(), {"in_memory": True})
        worksheet = workbook.add_worksheet()
        with self.assertRaises(TypeError) as ctx:
            worksheet.write(0, 0, b"hello")
        self.assertIn("bytes", str(ctx.exception))
        workbook.close()

    def test_write_bytearray_raises(self):
        workbook = Workbook(BytesIO(), {"in_memory": True})
        worksheet = workbook.add_worksheet()
        with self.assertRaises(TypeError) as ctx:
            worksheet.write(0, 0, bytearray(b"hello"))
        self.assertIn("bytearray", str(ctx.exception))
        workbook.close()

    def test_write_numeric_bytes_raises(self):
        """float(b'1.5') would coerce; bytes must still be rejected."""
        workbook = Workbook(BytesIO(), {"in_memory": True})
        worksheet = workbook.add_worksheet()
        with self.assertRaises(TypeError):
            worksheet.write(0, 0, b"1.5")
        workbook.close()

    def test_write_string_bytes_raises(self):
        workbook = Workbook(BytesIO(), {"in_memory": True})
        worksheet = workbook.add_worksheet()
        with self.assertRaises(TypeError) as ctx:
            worksheet.write_string(0, 0, b"hello")
        self.assertIn("bytes", str(ctx.exception))
        workbook.close()

    def test_write_str_still_ok(self):
        workbook = Workbook(BytesIO(), {"in_memory": True})
        worksheet = workbook.add_worksheet()
        self.assertEqual(worksheet.write(0, 0, "hello"), 0)
        workbook.close()
