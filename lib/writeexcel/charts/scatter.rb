###############################################################################
#
# Scatter - A writer class for Excel Scatter charts.
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
  class Scatter < Chart  # :nodoc:
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
      record       = 0x101B     # Record identifier.
      length       = 0x0006     # Number of bytes to follow.
      bubble_ratio = 0x0064     # Bubble ratio.
      bubble_type  = 0x0001     # Bubble type.
      grbit        = 0x0000     # Option flags.

      header = [record, length].pack('vv')
      data   = [bubble_ratio].pack('v')
      data  += [bubble_type].pack('v')
      data  += [grbit].pack('v')

      append(header, data)
    end

    ###############################################################################
    #
    # _store_axis_category_stream(). Overridden.
    #
    # Write the AXIS chart substream for the chart category.
    #
    # For a Scatter chart the category stream is replace with a values stream. We
    # override this method and turn it into a values stream.
    #
    def store_axis_category_stream
      store_axis(0)

      store_begin
      store_valuerange
      store_tick
      store_end
    end

    ###############################################################################
    #
    # _store_marker_dataformat_stream(). Overridden.
    #
    # This is an implementation of the parent abstract method  to define
    # properties of markers, linetypes, pie formats and other.
    #
    def store_marker_dataformat_stream
      store_dataformat(0x0000, 0xFFFD, 0x0000)

      store_begin
      store_3dbarshape
      store_lineformat(0x00000000, 0x0005, 0xFFFF, 0x0008, 0x004D)
      store_areaformat(0x00FFFFFF, 0x0000, 0x01, 0x01, 0x4E, 0x4D)
      store_pieformat
      store_markerformat(0x00, 0x00, 0x02, 0x01, 0x4D, 0x4D, 0x3C)
      store_end
    end
  end
end
