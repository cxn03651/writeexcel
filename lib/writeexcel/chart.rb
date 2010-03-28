###############################################################################
#
# Chart - A writer class for Excel Charts.
#
#
# Used in conjunction with WriteExcel
#
# Copyright 2000-2008, John McNamara, jmcnamara@cpan.org
#
# original written in Perl by John McNamara
# converted to Ruby by Hideo Nakamura, cxn03651@msj.biglobe.ne.jp
#

require 'writeexcel/worksheet'

# = Chart
# Charts and WriteExcel - A short introduction on how include externally
# generated charts into a WriteExcel file.
#
# = DESCRIPTION
# This document explains how to import Excel charts into a WriteExcel file.
#
# Please note that this feature is experimental. It may not work
# in all cases and it is best to start with a simple Excel file
# and gradually add complexity.
#
# = METHODOLOGY
# The general methodology is to create a chart in Excel, extract
# the chart from the binary file, import it into
# 'WriteExcel' and add new data to the series that
# the chart uses.
#
# The steps involved are as follows:
#
# 1. Create a new workbook in Excel or with Spreadsheet::WriteExcel. The file
# should be in Excel 97 or later format.
# 2. Add one or more worksheets with sample data of the type and format that
# you would like to have in the final version.
# 3. Create a chart on a new chart sheet that refers to a data range in one
# of the worksheets. Charts embedded in worksheets are also supported.
# See the 'add_chart_ext()' section of the main WriteExcel documentation.
# 4. Extend the chart data series to cover a sufficient range for any
# additional data that might be added. For example, if you initially have
# only 10 data points but you think that you may add up to 2000 at a later
# stage then increase the chart data series to 2000 points. In this case you
# should probably also leave the axes on automatic scaling.
# 5. Format the chart as you would like it to appear in the final version.
# 6. Save the workbook.
# 7. Using the 'chartex'* or 'chartex.pl' utility extract the chart(s) from
# the Excel file:
#             chartex file.xls
#
#             or
#
#             perl chartex.pl file.xls
#
#         * If you performed a normal installation then the 'chartex'
#         utility should be installed to your 'perl/bin' directory and
#         should be available from the command line.
#
# 8. Create a new 'Spreadsheet::WriteExcel' file with the same worksheets as
# the original file.
# 9. Add the external chart data to the 'Spreadsheet::WriteExcel' file:
#             my $chart = $workbook->add_chart_ext('chart01.bin', 'Chart1');
#
# In this case the 'chart01.bin' file is the chart data that
# was extracted in the Step 7.
#
# 10. Create a link between the chart and the target worksheet using a dummy
# formula, this is discussed in more detail below:
#             $worksheet->store_formula('=Sheet1!A1');
#
# 11. Add 3 or more additional formats to match any formats used in the chart,
# for example in the axis labels or the title. You may also have to adjust
# some of the font properties such as 'bold' or 'italic' to obtain the required
# font formats in the final chart.
#             $workbook->add_format(color => $_, bold => 1) for 1 ..5;
#
# If you do not supply enough additional formats then you may
# see the following error when you open the file in Excel:
# File error: data may have been lost. The file will still
# load but some formatting will have been lost.
#
# 12. Add new data to the data ranges defined in the chart using the standard
# WriteExcel interface.
#
# EXAMPLE
#     This following is a short example which uses line chart to
#     display some X-Y data:
#
#         #!/usr/bin/perl -w
#
#         use strict;
#         use Spreadsheet::WriteExcel;
#
#         my $workbook  = Spreadsheet::WriteExcel->new("demo01.xls");
#         my $worksheet = $workbook->add_worksheet();
#
#         my $chart     = $workbook->add_chart_ext('chart01.bin', 'Chart1');
#
#         $worksheet->store_formula('=Sheet1!A1');
#
#         $workbook->add_format(color => 1);
#         $workbook->add_format(color => 2, bold => 1);
#         $workbook->add_format(color => 3);
#
#         my @nums    = (0, 1, 2, 3, 4,  5,  6,  7,  8,  9,  10 );
#         my @squares = (0, 1, 4, 9, 16, 25, 36, 49, 64, 81, 100);
#
#         $worksheet->write_col('A1', \@nums   );
#         $worksheet->write_col('B1', \@squares);
#
#     This can be viewed in terms of the steps outlined above:
#
#     Steps 1-6. Create a workbook with a chart based on data in the
#     first worksheet. Otherwise use the 'Chart1.xls' file in the
#     'charts' directory of the distro as a template.
#
#     Step 7. Extract the chart data:
#
#         perl chartex.pl file.xls
#
#         Extracting "Chart1" to chart01.bin
#
#         ============================================================
#         Add the following near the start of your program.
#         Change variable name $worksheet if required.
#
#             $worksheet->store_formula("=Sheet1!A1");
#
#     Step 8. Create the new 'Spreadsheet::WriteExcel' file with the
#     same worksheets as the original file.
#
#         #!/usr/bin/perl -w
#
#         use strict;
#         use Spreadsheet::WriteExcel;
#
#         my $workbook  = Spreadsheet::WriteExcel->new("demo01.xls");
#         my $worksheet = $workbook->add_worksheet();
#
#     Step 9. Add the external chart data to the
#     'Spreadsheet::WriteExcel' file:
#
#         my $chart     = $workbook->add_chart_ext('chart01.bin', 'Chart1');
#
#     Step 10. Create a link between the chart and the worksheet using
#     a dummy formula:
#
#         $worksheet->store_formula('=Sheet1!A1');
#
#     Step 11. Add 3 or more additional formats to match any formats
#     used in the chart.
#
#         $workbook->add_format(color => 1);
#         $workbook->add_format(color => 2, bold => 1);
#         $workbook->add_format(color => 3);
#
#     Step 12. Add new data to the data ranges defined in the chart.
#
#         my @nums    = (0, 1, 2, 3, 4,  5,  6,  7,  8,  9,  10 );
#         my @squares = (0, 1, 4, 9, 16, 25, 36, 49, 64, 81, 100);
#
#         $worksheet->write_col('A1', \@nums   );
#         $worksheet->write_col('B1', \@squares);
#
#     See also the 'demo1.pl', 'demo2.pl' and 'demo3.pl' example
#     programs in the 'charts' directory of the distro.
#
# LINKING CHARTS AND DATA
#     Excel maintains links between charts and their data using
#     references. For example 'Sheet3' in the following chart series
#     would be stored internally in a reference table using a zero-
#     based integer.
#
#         =SERIES(,Sheet3!$A$2:$A$100,Sheet3!$B$2:$B$100,1)
#
#     These references are also shared with formulas that refer to
#     sheetnames. For example the following would share the same
#     reference as the previous chart series:
#
#         =Sheet3!A1
#
#     Therefore, we can simulate the link between the chart and the
#     worksheet data using a dummy formula:
#
#         $worksheet->store_formula('=Sheet3!A1');
#
#     When you run the 'chartex' program it will suggest the required
#     links:
#
#         ============================================================
#         Add the following near the start of your program.
#         Change variable name $worksheet if required.
#
#             $worksheet->store_formula("=Sheet3!A1");
#
#     This method is a workaround and will hopefully be made more
#     transparent in a future release.
#
# SEE ALSO
#     The Spreadsheet::WriteExcel documentation.
#
#     The 'demo1.pl', 'demo2.pl' and 'demo3.pl' example programs in
#     the 'charts' directory of the distro.
#
# BUGS
#     If you wish to submit a bug report run the 'bug_report.pl'
#     program in the 'examples' directory of the distro.
#
# AUTHOR
#     John McNamara jmcnamara@cpan.org
#
# COPYRIGHT
#     MMV, John McNamara.
#
#     All Rights Reserved. This module is free software. It may be
#     used, redistributed and/or modified under the same terms as Perl
#     itself.
#
#
class Chart < Worksheet

  ###############################################################################
  #
  # new()
  #
  # Constructor. Creates a new Chart object from a BIFFwriter object
  #
  def initialize(workbook, filename, name, index, encoding, activesheet, firstsheet, external_bin = nil)
    super(workbook, name, index, encoding)

    @filename          = filename
    @name              = name
    @index             = index
    @encoding          = encoding
    @activesheet       = activesheet
    @firstsheet        = firstsheet
    @external_bin      = external_bin

    @type              = 0x0200
    @using_tmpfile     = false
    @filehandle        = nil
    @xls_rowmax        = 0
    @xls_colmax        = 0
    @xls_strmax        = 0
    @dim_rowmin        = 0
    @dim_rowmax        = 0
    @dim_colmin        = 0
    @dim_colmax        = 0
    @dim_changed       = 0

    _initialize
  end

  ###############################################################################
  #
  # _close()
  #
  # Add data to the beginning of the workbook (note the reverse order)
  # and to the end of the workbook.
  #
  def close(*args)
  end

  ###############################################################################
  #
  # _initialize()
  #
  def _initialize
    if @external_bin
      filehandle = open(@filename, "rb") or
        die "Couldn't open #{@filename} in add_chart_ext(): $!.\n"
      @filehandle = filehandle
      @datasize   = FileTest.size(@filename)
      @using_tmpfile = false

      # Read the entire external chart binary into the the data buffer.
      # This will be retrieved by _get_data() when the chart is closed().
      @data = @filehandle.read(@datasize)
    end
  end
  private :_initialize

  ###############################################################################
  #
  # _store_fbi()
  #
  # Write the FBI chart BIFF record. Specifies the font information at the time
  # it was applied to the chart.
  #
  def store_fbi(index)
    record       = 0x1060    # Record identifier.
    length       = 0x000A    # Number of bytes to follow.
    # index                  # Font index.
    height       = 0x00C8    # Default font height in twips.
    width_basis  = 0x38B8    # Width basis, in twips.
    height_basis = 0x22A1    # Height basis, in twips.
    scale_basis  = 0x0000    # Scale by chart area or plot area.

    header = [record, length].pack('vv')
    data   = [width_basis].pack('v')
    data  += [height_basis].pack('v')
    data  += [height].pack('v')
    data  += [scale_basis].pack('v')
    data  += [index].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_chart()
  #
  # Write the CHART BIFF record. This indicates the start of the chart sub-stream
  # and contains dimensions of the chart on the display. Units are in 1/72 inch
  # and are 2 byte integer with 2 byte fraction.
  #
  def store_chart
    record = 0x1002        # Record identifier.
    length = 0x0010        # Number of bytes to follow.
    x_pos  = 0x00000000    # X pos of top left corner.
    y_pos  = 0x00000000    # Y pos of top left corner.
    dx     = 0x02DD51E0    # X size.
    dy     = 0x01C2B838    # Y size.

    header = [record, length].pack('vv')
    data   = [x_pos].pack('V')
    data  += [y_pos].pack('V')
    data  += [dx].pack('V')
    data  += [dy].pack('V')

    append(header, data)
  end
end
