###############################################################################
#
# Area - A writer class for Excel Area charts.
#
# Used in conjunction with Spreadsheet::WriteExcel::Chart.
#
# See formatting note in Spreadsheet::WriteExcel::Chart.
#
# Copyright 2000-2009, John McNamara, jmcnamara@cpan.org
#
# original written in Perl by John McNamara
# converted to Ruby by Hideo Nakamura, cxn03651@msj.biglobe.ne.jp
#

require 'writeexcel'

class Chart
  class Area < Chart
    ###############################################################################
    #
    # new()
    #
    #
    def initialize(*args)
      super
    end

    ###############################################################################
    #
    # _store_chart_type()
    #
    # Implementation of the abstract method from the specific chart class.
    #
    # Write the AREA chart BIFF record. Defines a area chart type.
    #
    def store_chart_type
      record = 0x101A     # Record identifier.
      length = 0x0002     # Number of bytes to follow.
      grbit  = 0x0001     # Option flags.

      header = [record, length].pack('vv')
      data = [grbit].pack('v')

      append(header, data)
    end
  end
end
