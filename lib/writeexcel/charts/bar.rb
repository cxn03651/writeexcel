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
      super(*args)
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
    # _store_x_axis_text_stream()
    #
    # Write the X-axis TEXT substream. Override the parent class because the axes
    # are reversed.
    #
    def store_x_axis_text_stream
      formula = @x_axis_formula.nil? ? '' : @x_axis_formula
      ai_type = _formula_type(2, 1, formula)

      store_text(0x002D, 0x06D9, 0x5F, 0x1CC, 0x0281, 0x00, 90)

      store_begin
      store_pos(2, 2, 0, 0, 0x17, 0x2A)
      store_fontx(8)
      store_ai(0, ai_type, formula)

      unless @x_axis_formula.nil?
        store_seriestext(@x_axis_name, @x_axis_encoding)
      end

      store_objectlink(3)
      store_end
    end

    ###############################################################################
    #
    # _store_y_axis_text_stream()
    #
    # Write the X-axis TEXT substream. Override the parent class because the axes
    # are reversed.
    #
    def store_y_axis_text_stream
      formula = @y_axis_formula.nil? ? '' : @y_axis_formula
      ai_type = _formula_type(2, 1, formula)

      store_text(0x078A, 0x0DFC, 0x011D, 0x9C, 0x0081, 0x0000)

      store_begin
      store_pos(2, 2, 0, 0, 0x45, 0x17)
      store_fontx(8)
      store_ai(0, ai_type, formula)

      unless @y_axis_formula.nil?
        store_seriestext(@y_axis_name, @y_axis_encoding)
      end

      store_objectlink(2)
      store_end
    end
  end
end
