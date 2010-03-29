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
    @embedded          = 0
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
  # BIFF Records.
  #
  ###############################################################################

  ###############################################################################
  #
  # _store_3dbarshape()
  #
  # Write the 3DBARSHAPE chart BIFF record.
  #
  def store_3dbarshape
    record = 0x105F    # Record identifier.
    length = 0x0002    # Number of bytes to follow.
    riser  = 0x00      # Shape of base.
    taper  = 0x00      # Column taper type.

    header = [record, length].pack('vv')
    data   = [riser].pack('C')
    data  += [taper].pack('C')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_ai()
  #
  # Write the AI chart BIFF record.
  #
  def store_ai(id, type, format_index, formula)
    record       = 0x1051     # Record identifier.
    length       = 0x0008     # Number of bytes to follow.
    # id                      # Link index.
    # type                    # Reference type.
    # format_index            # Num format index.
    # formula                 # Pre-parsed formula.
    grbit        = 0x0000     # Option flags.

    formula_length  = formula.length
    length += formula_length

    header = [record, length].pack('vv')
    data   = [id].pack('C')
    data  += [type].pack('C')
    data  += [grbit].pack('v')
    data  += [format_index].pack('v')
    data  += [formula_length].pack('v')
    data  += formula

    append(header, data)
  end

  ###############################################################################
  #
  # _store_areaformat()
  #
  # Write the AREAFORMAT chart BIFF record. Contains the patterns and colours
  # of a chart area.
  #
  def store_areaformat(rgbFore, rgbBack, pattern, grbit, indexFore, indexBack)
    record    = 0x100A     # Record identifier.
    length    = 0x0010     # Number of bytes to follow.
    # rgbFore              # Foreground RGB colour.
    # rgbBack              # Background RGB colour.
    # pattern              # Pattern.
    # grbit                # Option flags.
    # indexFore            # Index to Foreground colour.
    # indexBack            # Index to Background colour.

    header = [record, length].pack('vv')
    data  = [rgbFore].pack('V')
    data += [rgbBack].pack('V')
    data += [pattern].pack('v')
    data += [grbit].pack('v')
    data += [indexFore].pack('v')
    data += [indexBack].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_axcext()
  #
  # Write the AXCEXT chart BIFF record.
  #
  def store_axcext
    record       = 0x1062     # Record identifier.
    length       = 0x0012     # Number of bytes to follow.
    catMin       = 0x0000     # Minimum category on axis.
    catMax       = 0x0000     # Maximum category on axis.
    catMajor     = 0x0001     # Value of major unit.
    unitMajor    = 0x0000     # Units of major unit.
    catMinor     = 0x0001     # Value of minor unit.
    unitMinor    = 0x0000     # Units of minor unit.
    unitBase     = 0x0000     # Base unit of axis.
    catCrossDate = 0x0000     # Crossing point.
    grbit        = 0x00EF     # Option flags.

    header = [record, length].pack('vv')
    data  = [catMin].pack('v')
    data += [catMax].pack('v')
    data += [catMajor].pack('v')
    data += [unitMajor].pack('v')
    data += [catMinor].pack('v')
    data += [unitMinor].pack('v')
    data += [unitBase].pack('v')
    data += [catCrossDate].pack('v')
    data += [grbit].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_axesused()
  #
  # Write the AXESUSED chart BIFF record.
  #
  def store_axesused(num_axes)
    record   = 0x1046     # Record identifier.
    length   = 0x0002     # Number of bytes to follow.
    # num_axes            # Number of axes used.

    header = [record, length].pack('vv')
    data = [num_axes].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_axis()
  #
  # Write the AXIS chart BIFF record to define the axis type.
  #
  def store_axis(type)
    record    = 0x101D;        # Record identifier.
    length    = 0x0012;        # Number of bytes to follow.
    # type                     # Axis type.
    reserved1 = 0x00000000;    # Reserved.
    reserved2 = 0x00000000;    # Reserved.
    reserved3 = 0x00000000;    # Reserved.
    reserved4 = 0x00000000;    # Reserved.

    header = [record, length].pack('vv')
    data  = [type].pack('v')
    data += [reserved1].pack('V')
    data += [reserved2].pack('V')
    data += [reserved3].pack('V')
    data += [reserved4].pack('V')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_axislineformat()
  #
  # Write the AXISLINEFORMAT chart BIFF record.
  #
  def store_axislineformat
    record      = 0x1021     # Record identifier.
    length      = 0x0002     # Number of bytes to follow.
    line_format = 0x0001     # Axis line format.

    header = [record, length].pack('vv')
    data = [line_format].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_axisparent()
  #
  # Write the AXISPARENT chart BIFF record.
  #
  def store_axisparent(*args)
    record = 0x1041         # Record identifier.
    length = 0x0012         # Number of bytes to follow.
    iax    = args[0]        # Axis index.
    x      = args[1]        # X-coord.
    y      = args[2]        # Y-coord.
    dx     = args[3]        # Length of x axis.
    dy     = args[4]        # Length of y axis.

    header = [record, length].pack('vv')
    data   = [iax].pack('v')
    data  += [x].pack('V')
    data  += [y].pack('V')
    data  += [dx].pack('V')
    data  += [dy].pack('V')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_begin()
  #
  # Write the BEGIN chart BIFF record to indicate the start of a sub stream.
  #
  def store_begin
    record = 0x1033     # Record identifier.
    length = 0x0000     # Number of bytes to follow.

    header = [record, length].pack('vv')

    append(header)
  end

  ###############################################################################
  #
  # _store_catserrange()
  #
  # Write the CATSERRANGE chart BIFF record.
  #
  def store_catserrange
    record   = 0x1020     # Record identifier.
    length   = 0x0008     # Number of bytes to follow.
    catCross = 0x0001     # Value/category crossing.
    catLabel = 0x0001     # Frequency of labels.
    catMark  = 0x0001     # Frequency of ticks.
    grbit    = 0x0001     # Option flags.

    header = [record, length].pack('vv')
    data   = [catCross].pack('v')
    data  += [catLabel].pack('v')
    data  += [catMark].pack('v')
    data  += [grbit].pack('v')

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
  def store_chart(x_pos, y_pos, dx, dy)
    record   = 0x1002     # Record identifier.
    length   = 0x0010     # Number of bytes to follow.
    # x_pos               # X pos of top left corner.
    # y_pos               # Y pos of top left corner.
    # dx                  # X size.
    # dy                  # Y size.

    header = [record, length].pack('vv')
    data   = [x_pos].pack('V')
    data  += [y_pos].pack('V')
    data  += [dx].pack('V')
    data  += [dy].pack('V')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_chartformat()
  #
  # Write the CHARTFORMAT chart BIFF record. The parent record for formatting
  # of a chart group.
  #
  def store_chartformat(grbit = 0)
    record    = 0x1014         # Record identifier.
    length    = 0x0014         # Number of bytes to follow.
    reserved1 = 0x00000000     # Reserved.
    reserved2 = 0x00000000     # Reserved.
    reserved3 = 0x00000000     # Reserved.
    reserved4 = 0x00000000     # Reserved.
    # grbit                    # Option flags.
    icrt      = 0x0000         # Drawing order.

    header = [record, length].pack('vv')
    data   = [reserved1].pack('V')
    data  += [reserved2].pack('V')
    data  += [reserved3].pack('V')
    data  += [reserved4].pack('V')
    data  += [grbit].pack('v')
    data  += [icrt].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_chartline()
  #
  # Write the CHARTLINE chart BIFF record.
  #
  def store_chartline
    record = 0x101C     # Record identifier.
    length = 0x0002     # Number of bytes to follow.
    type   = 0x0001     # Drop/hi-lo line type.

    header = [record, length].pack('vv')
    data   = [type].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_charttext()
  #
  # Write the TEXT chart BIFF record.
  #
  def store_charttext
    record           = 0x1025         # Record identifier.
    length           = 0x0020         # Number of bytes to follow.
    horz_align       = 0x02           # Horizontal alignment.
    vert_align       = 0x02           # Vertical alignment.
    bg_mode          = 0x0001         # Background display.
    text_color_rgb   = 0x00000000     # Text RGB colour.
    text_x           = 0xFFFFFF46     # Text x-pos.
    text_y           = 0xFFFFFF06     # Text y-pos.
    text_dx          = 0x00000000     # Width.
    text_dy          = 0x00000000     # Height.
    grbit1           = 0x00B1         # Options
    text_color_index = 0x004D         # Auto Colour.
    grbit2           = 0x0000         # Data label placement.
    rotation         = 0x0000         # Text rotation.

    header = [record, length].pack('vv')
    data  = [horz_align].pack('C')
    data += [vert_align].pack('C')
    data += [bg_mode].pack('v')
    data += [text_color_rgb].pack('V')
    data += [text_x].pack('V')
    data += [text_y].pack('V')
    data += [text_dx].pack('V')
    data += [text_dy].pack('V')
    data += [grbit1].pack('v')
    data += [text_color_index].pack('v')
    data += [grbit2].pack('v')
    data += [rotation].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_dataformat()
  #
  # Write the DATAFORMAT chart BIFF record. This record specifies the series
  # that the subsequent sub stream refers to.
  #
  def store_dataformat(series_index, series_number, point_number)
    record        = 0x1006     # Record identifier.
    length        = 0x0008     # Number of bytes to follow.
    # series_index             # Series index.
    # series_number            # Series number. (Same as index).
    # point_number             # Point number.
    grbit         = 0x0000     # Format flags.

    header = [record, length].pack('vv')
    data  = [point_number].pack('v')
    data += [series_index].pack('v')
    data += [series_number].pack('v')
    data += [grbit].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_defaulttext()
  #
  # Write the DEFAULTTEXT chart BIFF record. Identifier for subsequent TEXT
  # record.
  #
  def store_defaulttext
    record = 0x1024;    # Record identifier.
    length = 0x0002;    # Number of bytes to follow.
    type   = 0x0002;    # Type.

    header = [record, length].pack('vv')
    data  = [type].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_dropbar()
  #
  # Write the DROPBAR chart BIFF record.
  #
  def store_dropbar
    record      = 0x103D     # Record identifier.
    length      = 0x0002     # Number of bytes to follow.
    percent_gap = 0x0096     # Drop bar width gap (%).

    header = [record, length].pack('vv')
    data  = [percent_gap].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_end()
  #
  # Write the END chart BIFF record to indicate the end of a sub stream.
  #
  def store_end
    record = 0x1034     # Record identifier.
    length = 0x0000     # Number of bytes to follow.

    header = [record, length].pack('vv')

    append(header, data)
  end

  ###############################################################################
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
  # _store_fontx()
  #
  # Write the FONTX chart BIFF record which contains the index of the FONT
  # record in the Workbook.
  #
  def store_fontx(index)
    record = 0x1026     # Record identifier.
    length = 0x0002     # Number of bytes to follow.
    # index             # Font index.

    header = [record, length].pack('vv')
    data   = [index].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_frame()
  #
  # Write the FRAME chart BIFF record.
  #
  def store_frame(frame_type, grbit)
    record     = 0x1032     # Record identifier.
    length     = 0x0004     # Number of bytes to follow.
    # frame_type            # Frame type.
    # grbit                 # Option flags.

    header = [record, length].pack('vv')
    data  = [frame_type].pack('v')
    data += [grbit].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_legend()
  #
  # Write the LEGEND chart BIFF record. The Marcus Horan method.
  #
  def store_legend(x, y, width, height, wType, wSpacing, grbit)
    record   = 0x1015     # Record identifier.
    length   = 0x0014     # Number of bytes to follow.
    # x                   # X-position.
    # y                   # Y-position.
    # width               # Width.
    # height              # Height.
    # wType               # Type.
    # wSpacing            # Spacing.
    # grbit               # Option flags.

    header = [record, length].pack('vv')
    data  = [x].pack('V')
    data += [y].pack('V')
    data += [width].pack('V')
    data += [height].pack('V')
    data += [wType].pack('C')
    data += [wSpacing].pack('C')
    data += [grbit].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_lineformat()
  #
  # Write the LINEFORMAT chart BIFF record.
  #
  def store_lineformat(rgb, lns, we, grbit, index)
    record = 0x1007     # Record identifier.
    length = 0x000C     # Number of bytes to follow.
    # rgb               # Line RGB colour.
    # lns               # Line pattern.
    # we                # Line weight.
    # grbit             # Option flags.
    # index             # Index to colour of line.

    header = [record, length].pack('vv')
    data  = [rgb].pack('V')
    data += [lns].pack('v')
    data += [we].pack('v')
    data += [grbit].pack('v')
    data += [index].pack('v')

    append(header, data)
 end

  ###############################################################################
  #
  # _store_markerformat()
  #
  # Write the MARKERFORMAT chart BIFF record.
  #
  def store_markerformat(rgbFore, rgbBack, marker, grbit, icvFore, icvBack, miSize)
    record  = 0x1009     # Record identifier.
    length  = 0x0014     # Number of bytes to follow.
    # rgbFore            # Foreground RGB color.
    # rgbBack            # Background RGB color.
    # marker             # Type of marker.
    # grbit              # Format flags.
    # icvFore            # Color index marker border.
    # icvBack            # Color index marker fill.
    # miSize             # Size of line markers.

    header = [record, length].pack('vv')
    data  = [rgbFore].pack('V')
    data += [rgbBack].pack('V')
    data += [marker].pack('v')
    data += [grbit].pack('v')
    data += [icvFore].pack('v')
    data += [icvBack].pack('v')
    data += [miSize].pack('V')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_objectlink()
  #
  # Write the OBJECTLINK chart BIFF record.
  #
  def store_objectlink(link_type)
    record      = 0x1027     # Record identifier.
    length      = 0x0006     # Number of bytes to follow.
    # link_type              # Object text link type.
    link_index1 = 0x0000     # Link index 1.
    link_index2 = 0x0000     # Link index 2.

    header = [record, length].pack('vv')
    data  = [link_type].pack('v')
    data += [link_index1].pack('v')
    data += [link_index2].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_pieformat()
  #
  # Write the PIEFORMAT chart BIFF record.
  #
  def store_pieformat
    record  = 0x100B     # Record identifier.
    length  = 0x0002     # Number of bytes to follow.
    percent = 0x0000     # Distance % from center.

    header = [record, length].pack('vv')
    data   = [percent].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_plotarea()
  #
  # Write the PLOTAREA chart BIFF record. This indicates that the subsequent
  # FRAME record belongs to a plot area.
  #
  def store_plotarea
    record = 0x1035     # Record identifier.
    length = 0x0000     # Number of bytes to follow.

    header = [record, length].pack('vv')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_plotgrowth()
  #
  # Write the PLOTGROWTH chart BIFF record.
  #
  def store_plotgrowth
    record  = 0x1064         # Record identifier.
    length  = 0x0008         # Number of bytes to follow.
    dx_plot = 0x00010000     # Horz growth for font scale.
    dy_plot = 0x00010000     # Vert growth for font scale.

    header = [record, length].pack('vv')
    data  = [dx_plot].pack('V')
    data += [dy_plot].pack('V')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_pos()
  #
  # Write the POS chart BIFF record. Generally not required when using
  # automatic positioning.
  #
  def store_pos(mdTopLt, mdBotRt, x1, y1, x2, y2)
    record  = 0x104F     # Record identifier.
    length  = 0x0014     # Number of bytes to follow.
    # mdTopLt            # Top left.
    # mdBotRt            # Bottom right.
    # x1                 # X coordinate.
    # y1                 # Y coordinate.
    # x2                 # Width.
    # y2                 # Height.

    header = [record, length].pack('vv')
    data  = [mdTopLt].pack('v')
    data += [mdBotRt].pack('v')
    data += [x1].pack('V')
    data += [y1].pack('V')
    data += [x2].pack('V')
    data += [y2].pack('V')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_serauxtrend()
  #
  # Write the SERAUXTREND chart BIFF record.
  #
  def store_serauxtrend(reg_type, poly_order, equation, r_squared)
    record     = 0x104B     # Record identifier.
    length     = 0x001C     # Number of bytes to follow.
    # reg_type              # Regression type.
    # poly_order            # Polynomial order.
    # equation              # Display equation.
    # r_squared             # Display R-squared.
    # intercept             # Forced intercept.
    # forecast              # Forecast forward.
    # backcast              # Forecast backward.

    # TODO. When supported, intercept needs to be NAN if not used.
    # Also need to reverse doubles.
    intercept = ['FFFFFFFF0001FFFF'].pack('H*')
    forecast  = ['0000000000000000'].pack('H*')
    backcast  = ['0000000000000000'].pack('H*')

    header = [record, length].pack('vv')
    data  = [reg_type].pack('C')
    data += [poly_order].pack('C')
    data += intercept
    data += [equation].pack('C')
    data += [r_squared].pack('C')
    data += forecast
    data += backcast

    append(header, data)
  end

  ###############################################################################
  #
  # _store_series()
  #
  # Write the SERIES chart BIFF record.
  #
  def store_series(category_count, value_count)
    record         = 0x1003     # Record identifier.
    length         = 0x000C     # Number of bytes to follow.
    category_type  = 0x0001     # Type: category.
    value_type     = 0x0001     # Type: value.
    # category_count            # Num of categories.
    # value_count               # Num of values.
    bubble_type    = 0x0001     # Type: bubble.
    bubble_count   = 0x0000     # Num of bubble values.

    header = [record, length].pack('vv')
    data  = [category_type].pack('v')
    data += [value_type].pack('v')
    data += [category_count].pack('v')
    data += [value_count].pack('v')
    data += [bubble_type].pack('v')
    data += [bubble_count].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_seriestext()
  #
  # Write the SERIESTEXT chart BIFF record.
  #
  def store_seriestext(str, encoding)
    record   = 0x100D          # Record identifier.
    length   = 0x0000          # Number of bytes to follow.
    id       = 0x0000          # Text id.
    # str                      # Text.
    # encoding                 # String encoding.
    cch      = str.length      # String length.

    # Character length is num of chars not num of bytes
    cch /= 2 if encoding != 0

    # Change the UTF-16 name from BE to LE
    str = str.unpack('v*').pack('n*') if encoding != 0

    length = 4 + str.length

    header = [record, length].pack('vv')
    data  = [id].pack('v')
    data += [cch].pack('C')
    data += [encoding].pack('C')

    append(header, data, str)
  end

  ###############################################################################
  #
  # _store_serparent()
  #
  # Write the SERPARENT chart BIFF record.
  #
  def store_serparent(series)
    record = 0x104A     # Record identifier.
    length = 0x0002     # Number of bytes to follow.
    # series            # Series parent.

    header = [record, length].pack('vv')
    data = [series].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_sertocrt()
  #
  # Write the SERTOCRT chart BIFF record to indicate the chart group index.
  #
  def store_sertocrt
    record     = 0x1045     # Record identifier.
    length     = 0x0002     # Number of bytes to follow.
    chartgroup = 0x0000     # Chart group index.

    header = [record, length].pack('vv')
    data = [chartgroup].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_shtprops()
  #
  # Write the SHTPROPS chart BIFF record.
  #
  def store_shtprops
    record      = 0x1044     # Record identifier.
    length      = 0x0004     # Number of bytes to follow.
    grbit       = 0x000E     # Option flags.
    empty_cells = 0x0000     # Empty cell handling.

    grbit = 0x000A if @embedded != 0

    header = [record, length].pack('vv')
    data  = [grbit].pack('v')
    data += [empty_cells].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_text()
  #
  # Write the TEXT chart BIFF record.
  #
  def store_text(x, y, dx, dy, grbit1, grbit2, rotation = 0x00)
    record   = 0x1025;           # Record identifier.
    length   = 0x0020;           # Number of bytes to follow.
    at       = 0x02;             # Horizontal alignment.
    vat      = 0x02;             # Vertical alignment.
    wBkgMode = 0x0001;           # Background display.
    rgbText  = 0x0000;           # Text RGB colour.
    # x                          # Text x-pos.
    # y                          # Text y-pos.
    # dx                         # Width.
    # dy                         # Height.
    # grbit1                     # Option flags.
    icvText  = 0x004D;           # Auto Colour.
    # grbit2                     # Show legend.
    # rotation                   # Show value.

    header = [record, length].pack('vv')
    data  = [at].pack('C')
    data += [vat].pack('C')
    data += [wBkgMode].pack('v')
    data += [rgbText].pack('V')
    data += [x].pack('V')
    data += [y].pack('V')
    data += [dx].pack('V')
    data += [dy].pack('V')
    data += [grbit1].pack('v')
    data += [icvText].pack('v')
    data += [grbit2].pack('v')
    data += [rotation].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_tick()
  #
  # Write the TICK chart BIFF record.
  #
  def store_tick
    record    = 0x101E         # Record identifier.
    length    = 0x001E         # Number of bytes to follow.
    tktMajor  = 0x02           # Type of major tick mark.
    tktMinor  = 0x00           # Type of minor tick mark.
    tlt       = 0x03           # Tick label position.
    wBkgMode  = 0x01           # Background mode.
    rgb       = 0x00000000     # Tick-label RGB colour.
    reserved1 = 0x00000000     # Reserved.
    reserved2 = 0x00000000     # Reserved.
    reserved3 = 0x00000000     # Reserved.
    reserved4 = 0x00000000     # Reserved.
    grbit     = 0x0023         # Option flags.
    index     = 0x004D         # Colour index.
    reserved5 = 0x0000         # Reserved.

    header = [record, length].pack('vv')
    data  = [tktMajor].pack('C')
    data += [tktMinor].pack('C')
    data += [tlt].pack('C')
    data += [wBkgMode].pack('C')
    data += [rgb].pack('V')
    data += [reserved1].pack('V')
    data += [reserved2].pack('V')
    data += [reserved3].pack('V')
    data += [reserved4].pack('V')
    data += [grbit].pack('v')
    data += [index].pack('v')
    data += [reserved5].pack('v')

    append(header, data)
  end

  ###############################################################################
  #
  # _store_valuerange()
  #
  # Write the VALUERANGE chart BIFF record.
  #
  def store_valuerange
    record   = 0x101F         # Record identifier.
    length   = 0x002A         # Number of bytes to follow.
    numMin   = 0x00000000     # Minimum value on axis.
    numMax   = 0x00000000     # Maximum value on axis.
    numMajor = 0x00000000     # Value of major increment.
    numMinor = 0x00000000     # Value of minor increment.
    numCross = 0x00000000     # Value where category axis crosses.
    grbit    = 0x011F         # Format flags.

    # TODO. Reverse doubles when they are handled.

    header = [record, length].pack('vv')
    data  = [numMin].pack('d')
    data += [numMax].pack('d')
    data += [numMajor].pack('d')
    data += [numMinor].pack('d')
    data += [numCross].pack('d')
    data += [grbit].pack('v')

    append(header, data)
  end
end
