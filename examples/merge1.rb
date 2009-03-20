#!/usr/bin/ruby -w

###############################################################################
#
# Simple example of merging cells using the Spreadsheet::WriteExcel module.
#
# This merges three cells using the "Centre Across Selection" alignment.
# This was the Excel 5 method of achieving a merge. Use the merge_range()
# worksheet method instead. See merge3.pl - merge6.pl.
#
# reverse('©'), August 2002, John McNamara, jmcnamara@cpan.org
#

require 'writeexcel'
include Spreadsheet

# Create a new workbook and add a worksheet
workbook  = WriteExcel.new('merge1.xls')
worksheet = workbook.add_worksheet


# Increase the cell size of the merged cells to highlight the formatting.
worksheet.set_column('B:D', 20)
worksheet.set_row(2, 30)


# Create a merge format
format = workbook.add_format(:center_across => 1)


# Only one cell should contain text, the others should be blank.
worksheet.write(2, 1, "Center across selection", format)
worksheet.write_blank(2, 2,                 format)
worksheet.write_blank(2, 3,                 format)

workbook.close
