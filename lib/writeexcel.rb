###############################################################################
#
# WriteExcel.
#
# WriteExcel - Write to a cross-platform Excel binary file.
#
# Copyright 2000-2008, John McNamara, jmcnamara@cpan.org
#
# original written in Perl by John McNamara
# converted to Ruby by Hideo Nakamura, cxn03651@msj.biglobe.ne.jp
#
require "writeexcel/workbook"
#
# = WriteExcel - Write to a cross-platform Excel binary file.
#
# == Synopsis
#
# To write a string, a formatted string, a number and a formula to the first
# worksheet in an Excel workbook called ruby.xls:
#
#     require 'WriteExcel'
#
#     # Create a new Excel workbook
#     workbook = WriteExcel->new('ruby.xls')
#
#     # Add a worksheet
#     worksheet = workbook.add_worksheet
#
#     #  Add and define a format
#     format = workbook.add_format # Add a format
#     format.set_bold()
#     format.set_color('red')
#     format.set_align('center')
#
#     # Write a formatted and unformatted string, row and column notation.
#     col = row = 0
#     worksheet.write(row, col, 'Hi Excel!', format)
#     worksheet.write(1,   col, 'Hi Excel!')
#
#     # Write a number and a formula using A1 notation
#     worksheet.write('A3', 1.2345)
#     worksheet.write('A4', '=SIN(PI()/4)')
#
# == Description
#
# WriteExcel can be used to create a cross-platform Excel binary file.
# Multiple worksheets can be added to a workbook and formatting can be applied
# to cells. Text, numbers, formulas, hyperlinks and images can be written to
# the cells.
#
# The Excel file produced by this gem is compatible with 97, 2000, 2002 and 2003.
#
# WriteExcel will work on the majority of Windows, UNIX and Macintosh platforms.
# Generated files are also compatible with the Linux/UNIX spreadsheet
# applications Gnumeric and OpenOffice.org.
#
# This module cannot be used to write to an existing Excel file
#
# == Quick Start
#
# WriteExcel tries to provide an interface to as many of Excel's features as
# possible. As a result there is a lot of documentation to accompany the
# interface and it can be difficult at first glance to see what it important
# and what is not. So for those of you who prefer to assemble Ikea furniture
# first and then read the instructions, here are three easy steps:
#
# 1. Create a new Excel workbook (i.e. file) using new().
#
# 2. Add a worksheet to the new workbook using add_worksheet().
#
# 3. Write to the worksheet using write().
#
# Like this:
#
#     require 'WriteExcel'                     # Step 0
#
#     workbook  = WriteExcel.new('ruby.xls')   # Step 1
#     worksheet = workbook.add_worksheet       # Step 2
#     worksheet.write('A1', 'Hi Excel!')       # Step 3
#
# This will create an Excel file called ruby.xls with a single worksheet and the
# text 'Hi Excel!' in the relevant cell. And that's it. Okay, so there is
# actually a zeroth step as well, but use WriteExcel goes without saying. There
# are also more than 80 examples that come with the distribution and which you can
# use to get you started. See EXAMPLES.
#

class WriteExcel < Workbook
  VERSION = "0.2.2"
end
