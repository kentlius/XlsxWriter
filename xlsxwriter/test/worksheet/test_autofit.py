###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest

from xlsxwriter.format import Format
from xlsxwriter.sharedstrings import SharedStringTable
from xlsxwriter.worksheet import Worksheet


class TestAutofit(unittest.TestCase):
    """
    Test the Worksheet autofit() method.

    """

    def setUp(self):
        self.worksheet = Worksheet()
        self.worksheet.str_table = SharedStringTable()

    def test_autofit_no_rotation(self):
        """Test autofit of unrotated text."""

        self.worksheet.write_string(0, 0, "Hello")
        self.worksheet.autofit()

        self.assertEqual(40, self.worksheet.col_info[0].width)

    def test_autofit_rotation_90(self):
        """Test autofit of text rotated 90 degrees."""

        cell_format = Format()
        cell_format.set_rotation(90)

        self.worksheet.write_string(0, 0, "Very long text", cell_format)
        self.worksheet.autofit()

        # The column width is the font height plus cell padding.
        self.assertEqual(22, self.worksheet.col_info[0].width)

    def test_autofit_rotation_minus_90(self):
        """Test autofit of text rotated -90 degrees."""

        cell_format = Format()
        cell_format.set_rotation(-90)

        self.worksheet.write_string(0, 0, "Very long text", cell_format)
        self.worksheet.autofit()

        self.assertEqual(22, self.worksheet.col_info[0].width)

    def test_autofit_rotation_270(self):
        """Test autofit of vertically stacked text."""

        cell_format = Format()
        cell_format.set_rotation(270)

        self.worksheet.write_string(0, 0, "Very long text", cell_format)
        self.worksheet.autofit()

        self.assertEqual(22, self.worksheet.col_info[0].width)

    def test_autofit_rotation_45(self):
        """Test autofit of text rotated 45 degrees."""

        cell_format = Format()
        cell_format.set_rotation(45)

        self.worksheet.write_string(0, 0, "Very long text", cell_format)
        self.worksheet.autofit()

        self.assertEqual(79, self.worksheet.col_info[0].width)

    def test_autofit_rotation_45_font_size(self):
        """Test autofit of rotated text with a non-default font size."""

        cell_format = Format()
        cell_format.set_rotation(45)
        cell_format.set_font_size(22)

        self.worksheet.write_string(0, 0, "Very long text", cell_format)
        self.worksheet.autofit()

        self.assertEqual(89, self.worksheet.col_info[0].width)

    def test_autofit_rotation_multiline(self):
        """Test autofit of multi-line rotated text."""

        cell_format = Format()
        cell_format.set_rotation(90)

        self.worksheet.write_string(0, 0, "abc\ndef", cell_format)
        self.worksheet.autofit()

        # The column width is the height of two lines plus cell padding.
        self.assertEqual(37, self.worksheet.col_info[0].width)

    def test_autofit_rotation_column_format(self):
        """Test autofit of rotated text using a column format."""

        cell_format = Format()
        cell_format.set_rotation(90)

        self.worksheet.set_column(0, 0, None, cell_format)
        self.worksheet.write_string(0, 0, "Very long text")
        self.worksheet.autofit()

        self.assertEqual(22, self.worksheet.col_info[0].width)

    def test_autofit_rotation_row_format(self):
        """Test autofit of rotated text using a row format."""

        cell_format = Format()
        cell_format.set_rotation(90)

        self.worksheet.set_row(0, None, cell_format)
        self.worksheet.write_string(0, 0, "Very long text")
        self.worksheet.autofit()

        self.assertEqual(22, self.worksheet.col_info[0].width)

    def test_autofit_rotation_cell_format_precedence(self):
        """Test that a cell format overrides a rotated column format."""

        column_format = Format()
        column_format.set_rotation(90)

        self.worksheet.set_column(0, 0, None, column_format)
        self.worksheet.write_string(0, 0, "Very long text", Format())
        self.worksheet.autofit()

        self.assertEqual(94, self.worksheet.col_info[0].width)
