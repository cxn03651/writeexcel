###############################################################################
#
# Bar - A writer class for Excel Bar charts.
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

require 'writeexcel/chart'

class Chart
  class Bar < Chart
    ###############################################################################
    #
    # new()
    #
    #
    def initialize(*args)
      super
      @config[:x_axis_text]     = [ 0x2D,   0x6D9,  0x5F,   0x1CC, 0x281,  0x0, 90 ]
      @config[:x_axis_text_pos] = [ 2,      2,      0,      0,     0x17,   0x2A ]
      @config[:y_axis_text]     = [ 0x078A, 0x0DFC, 0x011D, 0x9C,  0x0081, 0x0000 ]
      @config[:y_axis_text_pos] = [ 2,      2,      0,      0,     0x45,   0x17 ]
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
      record    = 0x1017     # Record identifier.
      length    = 0x0006     # Number of bytes to follow.
      pcOverlap = 0x0000     # Space between bars.
      pcGap     = 0x0096     # Space between cats.
      grbit     = 0x0001     # Option flags.

      header = [record, length].pack('vv')
      data  = [pcOverlap].pack('v')
      data += [pcGap].pack('v')
      data += [grbit].pack('v')

      append(header, data)
    end

    ###############################################################################
    #
    # _set_embedded_config_data()
    #
    # Override some of the default configuration data for an embedded chart.
    #
    def set_embedded_config_data
      # Set the parent configuration first.
      super

      # The axis positions are reversed for a bar chart so we change the config.
      @config[:x_axis_text]     = [ 0x57,   0x5BC,  0xB5,   0x214, 0x281, 0x0, 90 ]
      @config[:x_axis_text_pos] = [ 2,      2,      0,      0,     0x17,  0x2A ]
      @config[:y_axis_text]     = [ 0x074A, 0x0C8F, 0x021F, 0x123, 0x81,  0x0000 ]
      @config[:y_axis_text_pos] = [ 2,      2,      0,      0,     0x45,  0x17 ]
    end
  end
end
