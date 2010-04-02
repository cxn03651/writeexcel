###############################################################################
#
# Line - A writer class for Excel Line charts.
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

class Chart
  class Line < Chart  # :nodoc:
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
    def store_chart_type
      record = 0x1018     # Record identifier.
      length = 0x0002     # Number of bytes to follow.
      grbit  = 0x0000     # Option flags.

      header = [record, length].pack('vv')
      data   = [grbit].pack('v')

      append(header, data)
    end
  end
end
