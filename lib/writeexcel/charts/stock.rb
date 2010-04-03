###############################################################################
#
# Stock - A writer class for Excel Stock charts.
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
  class Stock < Chart  # :nodoc:
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
    # Write the LINE chart BIFF record. A stock chart uses the same LINE record
    # as a line chart but with additional DROPBAR and CHARTLINE records to define
    # the stock style.
    #
    def store_chart_type
      record = 0x1018     # Record identifier.
      length = 0x0002     # Number of bytes to follow.
      grbit  = 0x0000     # Option flags.

      header = [record, length].pack('vv')
      data   = [grbit].pack('v')

      append(header, data)
    end

    ###############################################################################
    #
    # _store_marker_dataformat_stream(). Overridden.
    #
    # This is an implementation of the parent abstract method to define
    # properties of markers, linetypes, pie formats and other.
    #
    def store_marker_dataformat_stream
      store_dropbar
      store_begin
      store_lineformat(0x00000000, 0x0000, 0xFFFF, 0x0001, 0x004F)
      store_areaformat(0x00FFFFFF, 0x0000, 0x01, 0x01, 0x09, 0x08)
      store_end

      store_dropbar
      store_begin
      store_lineformat(0x00000000, 0x0000, 0xFFFF, 0x0001, 0x004F)
      store_areaformat(0x0000, 0x00FFFFFF, 0x01, 0x01, 0x08, 0x09)
      store_end

      store_chartline
      store_lineformat(0x00000000, 0x0000, 0xFFFF, 0x0000, 0x004F)


      store_dataformat(0x0000, 0xFFFD, 0x0000)
      store_begin
      store_3dbarshape
      store_lineformat(0x00000000, 0x0005, 0xFFFF, 0x0000, 0x004F)
      store_areaformat(0x00000000, 0x0000, 0x00, 0x01, 0x4D, 0x4D)
      store_pieformat
      store_markerformat(0x00, 0x00, 0x00, 0x00, 0x4D, 0x4D, 0x3C)
      store_end
    end
  end
end
