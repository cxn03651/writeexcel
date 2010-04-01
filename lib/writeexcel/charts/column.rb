###############################################################################
#
# Column - A writer class for Excel Column charts.
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
# To create a simple Excel file with a Column chart using WriteExcel:
#
#     #!/usr/bin/ruby -w
#
#     require 'writeexcel'
#
#     workbook  = WriteExcel.new('chart.xls')
#     worksheet = workbook.add_worksheet
#
#     chart     = workbook.add_chart(:type => Chart::Column)
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
#     ];
#
#     worksheet.write('A1', data)
#
#     workbook.close
#
# == DESCRIPTION ^
#
# This module implements Column charts for WriteExcel. The chart object is
# created via the Workbook add_chart method:
#
#     chart = workbook.add_chart(:type => Chart::Column)
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
# == Column Chart Methods ^
#
# There aren't currently any column chart specific methods.
# See the TODO section of Chart.
#
class Chart
  class Column < Chart
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
    # Write the BAR chart BIFF record. Defines a bar or column chart type.
    #
    def store_chart_type   # :nodoc:
      record    = 0x1017     # Record identifier.
      length    = 0x0006     # Number of bytes to follow.
      pcOverlap = 0x0000     # Space between bars.
      pcGap     = 0x0096     # Space between cats.
      grbit     = 0x0000     # Option flags.

      header = [record, length].pack('vv')
      data  = [pcOverlap].pack('v')
      data += [pcGap].pack('v')
      data += [grbit].pack('v')

      append(header, data)
    end
  end
end
