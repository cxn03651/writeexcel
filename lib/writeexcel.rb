###############################################################################
#
# WriteExcel.
#
# Spreadsheet::WriteExcel - Write to a cross-platform Excel binary file.
#
# Copyright 2000-2008, John McNamara, jmcnamara@cpan.org
#
# original written in Perl by John McNamara
# converted to Ruby by Hideo Nakamura, cxn03651@msj.biglobe.ne.jp
#
require "writeexcel/workbook"

class Spreadsheet
  class WriteExcel < Workbook
    VERSION = "0.2.1"
  end
end
