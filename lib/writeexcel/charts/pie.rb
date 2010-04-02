###############################################################################
#
# Pie - A writer class for Excel Pie charts.
#
# Used in conjunction with Spreadsheet::WriteExcel::Chart.
#
# See formatting note in Spreadsheet::WriteExcel::Chart.
#
# Copyright 2000-2010, John McNamara, jmcnamara@cpan.org
#
# original written in Perl by John McNamara
# converted to Ruby by Hideo Nakamura, cxn03651@msj.biglobe.ne.jp
#

require 'writeexcel'

class Chart
  class Pie < Chart  # :nodoc:
    ###############################################################################
    #
    # new()
    #
    #
    def initialize(*args)
      super
      @vary_data_color = 1
    end

    ###############################################################################
    #
    # _store_chart_type()
    #
    # Implementation of the abstract method from the specific chart class.
    #
    # Write the Pie chart BIFF record.
    #
    def store_chart_type
      record = 0x1019     # Record identifier.
      length = 0x0006     # Number of bytes to follow.
      angle  = 0x0000     # Angle.
      donut  = 0x0000     # Donut hole size.
      grbit  = 0x0002     # Option flags.

      header = [record, length].pack('vv')
      data  = [angle].pack('v')
      data += [donut].pack('v')
      data += [grbit].pack('v')

      append(header, data)
    end

    ###############################################################################
    #
    # _store_axisparent_stream(). Overridden.
    #
    # Write the AXISPARENT chart substream.
    #
    # A Pie chart has no X or Y axis so we override this method to remove them.
    #
    def store_axisparent_stream
      store_axisparent(*@config[:axisparent])

      store_begin
      store_pos(*@config[:axisparent_pos])

      store_chartformat_stream
      store_end
    end
  end
end
