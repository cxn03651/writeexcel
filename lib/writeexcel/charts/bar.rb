###############################################################################
#
# Bar - A writer class for Excel Bar charts.
#
# Used in conjunction with Chart.
#
# See formatting note in Chart.
#
# Copyright 2000-2009, John McNamara, jmcnamara@cpan.org
#
# original written in Perl by John McNamara
# converted to Ruby by Hideo Nakamura, cxn03651@msj.biglobe.ne.jp
#

require 'writeexcel/chart'

#
# == SYNOPSIS ^
#
# To create a simple Excel file with a Bar chart using WriteExcel:
#
#     #!/usr/bin/ruby -w
#
#     require 'writeexcel'
#
#     workbook  = WriteExcel.new('chart.xls')
#     worksheet = workbook.add_worksheet
#
#     chart     = workbook.add_chart(:type => Chart::Bar)
#
#     # Configure the chart.
#     chart.add_series(
#       :categories => '=Sheet1!$A$2:$A$7',
#       :values     => '=Sheet1!$B$2:$B$7'
#     )
#
#     # Add the data to the worksheet the chart refers to.
#     data = [
#         [ 'Category', 2, 3, 4, 5, 6, 7 ],
#         [ 'Value',    1, 4, 5, 2, 1, 5 ]
#     ]
#
#     worksheet.write('A1', data)
#
#     workbook.close
#
# == DESCRIPTION ^
#
# This module implements Bar charts for WriteExcel. The chart object is
# created via the Workbook add_chart method:
#
#     chart = workbook.add_chart(:type => Chart::Bar)
#
# Once the object is created it can be configured via the following methods
# that are common to all chart classes:
#
#     chart.add_series
#     chart.set_x_axis
#     chart.set_y_axis
#     chart.set_title
#
# These methods are explained in detail in Chart. Class specific methods or
# settings, if any, are explained below.
#
# == Bar Chart Methods ^
#
# There aren't currently any bar chart specific methods.
# See the TODO section of Spreadsheet::WriteExcel::Chart.
#
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
    def store_chart_type  # :nodoc:
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
    def set_embedded_config_data  # :nodoc:
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
