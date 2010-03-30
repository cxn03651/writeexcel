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

###############################################################################
#
# Formatting information.
#
# perltidy with options: -mbl=2 -pt=0 -nola
#
# Any camel case Hungarian notation style variable names in the BIFF record
# writing sub-routines below are for similarity with names used in the Excel
# documentation. Otherwise lowercase underscore style names are used.
#


###############################################################################
#
# The chart class hierarchy is as follows. Chart.pm acts as a factory for the
# sub-classes.
#
#
#     Spreadsheet::WriteExcel::BIFFwriter
#                     ^
#                     |
#     Spreadsheet::WriteExcel::Worksheet
#                     ^
#                     |
#     Spreadsheet::WriteExcel::Chart
#                     ^
#                     |
#     Spreadsheet::WriteExcel::Chart::* (sub-types)
#


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
  NonAscii = /[^!"#\$%&'\(\)\*\+,\-\.\/\:\;<=>\?@0-9A-Za-z_\[\\\]^` ~\0\n]/

  ###############################################################################
  #
  # factory()
  #
  # Factory method for returning chart objects based on their class type.
  #
  def self.factory(klass, *args)
    klass.new(*args)
  end

  ###############################################################################
  #
  # :call-seq:
  #   new(filename, name, index, encoding, activesheet, firstsheet, external_bin = nil)
  #
  # Default constructor for sub-classes.
  #
  def initialize(*args)
    super

    @sheet_type  = 0x0200
    @orientation = 0x0
    @series      = []
    @external_bin = false
    @x_axis_formula = nil
    @x_axis_name = nil
    @y_axis_formula = nil
    @y_axis_name = nil
    @title_name = nil
    @title_formula = nil
  end

  ###############################################################################
  #
  # add_series()
  #
  # Add a series and it's properties to a chart.
  #
  def add_series(params)
    raise "Must specify 'values' in add_series()" if params[:values].nil?

    # Parse the ranges to validate them and extract salient information.
    value_data    = parse_series_formula(params[:values])
    category_data = parse_series_formula(params[:categories])
    name_formula  = parse_series_formula(params[:name_formula])

    # Default category count to the same as the value count if not defined.
    category_data[1] = value_data[1] if category_data.size < 2

    # Add the parsed data to the user supplied data.
    params[:values]       = value_data
    params[:categories]   = category_data
    params[:name_formula] = name_formula

    # Encode the Series name.
    name, encoding = encode_utf16(params[:name], params[:name_encoding])

    params[:name]          = name
    params[:name_encoding] = encoding

    @series << params
  end

  ###############################################################################
  #
  # set_x_axis()
  #
  # Set the properties of the X-axis.
  #
  def set_x_axis(params)
    name, encoding = encode_utf16(params[:name], params[:name_encoding])
    formula = parse_series_formula(params[:name_formula])

    @x_axis_name     = name
    @x_axis_encoding = encoding
    @x_axis_formula  = formula
  end

  ###############################################################################
  #
  # set_y_axis()
  #
  # Set the properties of the Y-axis.
  #
  def set_y_axis(params)
    name, encoding = encode_utf16(params[:name], params[:name_encoding])
    formula = parse_series_formula(params[:name_formula])

    @y_axis_name     = name
    @y_axis_encoding = encoding
    @y_axis_formula  = formula
  end

  ###############################################################################
  #
  # set_title()
  #
  # TODO
  #
  def set_title(params)
    name, encoding = encode_utf16( params[:name], params[:name_encoding])

    formula = parse_series_formula(params[:name_formula])

    @title_name     = name
    @title_encoding = encoding
    @title_formula  = formula
  end

  ###############################################################################
  #
  # _prepend(), overridden.
  #
  # The parent Worksheet class needs to store some data in memory and some in
  # temporary files for efficiency. The Chart* classes don't need to do this
  # since they are dealing with smaller amounts of data so we override
  # _prepend() to turn it into an _append() method. This allows for a more
  # natural method calling order.
  #
  def prepend(*args)
    @using_tmpfile = false
    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
    append(*args)
  end

  ###############################################################################
  #
  # _close(), overridden.
  #
  # Create and store the Chart data structures.
  #
  def close
    # TODO
    return nil if @external_bin

    # Ignore any data that has been written so far since it is probably
    # from unwanted Worksheet method calls.
    @data = ''

    # TODO. Check for charts without a series?

    # Store the chart BOF.
    store_bof(0x0020)

    # Store the page header
    store_header

    # Store the page footer
    store_footer

    # Store the page horizontal centering
    store_hcenter

    # Store the page vertical centering
    store_vcenter

    # Store the left margin
    store_margin_left

    # Store the right margin
    store_margin_right

    # Store the top margin
    store_margin_top

    # Store the bottom margin
    store_margin_bottom

    # Store the page setup
    store_setup

    # Store the sheet password
    store_password

    # Start of Chart specific records.

    # Store the FBI font records.
    store_fbi(5, 10)
    store_fbi(6, 10)
    store_fbi(7, 12)
    store_fbi(8, 10)

    # Ignore UNITS record.

    # Store the Chart sub-stream.
    store_chart_stream

    # Append the sheet dimensions
    store_dimensions

    store_eof
  end

  #
  # TODO temp debug code
  #
  def store_tmp_records
    data = [

           ].pack('C*')

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
    append(data)
  end

  ###############################################################################
  #
  # _parse_series_formula()
  #
  # Parse the formula used to define a series. We also extract some range
  # information required for _store_series() and the SERIES record.
  #
  def parse_series_formula(formula)
    encoding = 0
    length   = 0
    count    = 0
    tokens = []

    return [''] if formula.nil?

    # Strip the = sign at the beginning of the formula string
    formula = formula.sub(/^=/, '')

    # Parse the formula using the parser in Formula.pm
    parser = @parser

    # In order to raise formula errors from the point of view of the calling
    # program we use an eval block and re-raise the error from here.
    #
    tokens = parser.parse_formula(formula)

    # Force ranges to be a reference class.
    tokens.collect! { |t| t.gsub(/_ref3d/, '_ref3dR') }
    tokens.collect! { |t| t.gsub(/_range3d/, '_range3dR') }
    tokens.collect! { |t| t.gsub(/_name/, '_nameR') }

    # Parse the tokens into a formula string.
    formula = parser.parse_tokens(tokens)

    # Return formula for a single cell as used by title and series name.
    return formula if formula[0] == 0x3A

    # Extract the range from the parse formula.
    if formula[0] == 0x3B
        ptg, ext_ref, row_1, row_2, col_1, col_2 = formula.unpack('Cv5')

        # TODO. Remove high bit on relative references.
        count = row_2 - row_1 + 1
    end

    [formula, count]
  end

  ###############################################################################
  #
  # _encode_utf16()
  #
  # Convert UTF8 strings used in the chart to UTF16.
  #
  def encode_utf16(str, encoding = 0)
    # Exit if the $string isn't defined, i.e., hasn't been set by user.
    return [nil, nil] if str.nil?

    string = str.dup
    # Return if encoding is set, i.e., string has been manually encoded.
    #return ( undef, undef ) if $string == 1;

    # Handle utf8 strings in perl 5.8.
    if string =~ NonAscii
      string = NKF.nkf('-w16B0 -m0 -W', string)
      encoding = 1
    end

    # Chart strings are limited to 255 characters.
    limit = encoding != 0 ? 255 * 2 : 255

    if string.length >= limit
      # truncate the string and raise a warning.
      string = string[0, limit]
    end

    [string, encoding]
  end

  ###############################################################################
  #
  # _store_chart_stream()
  #
  # Store the CHART record and it's substreams.
  #
  def store_chart_stream # :nodoc:
    store_chart
    store_begin

    # Store the chart SCL record.
    store_plotgrowth

    # Store SERIES stream for each series.
    index = 0
    @series.each do |series|
      store_series_stream(
          :index            => index,
          :value_formula    => series[:values][0],
          :value_count      => series[:values][1],
          :category_count   => series[:categories][1],
          :category_formula => series[:categories][0],
          :name             => series[:name],
          :name_encoding    => series[:name_encoding],
          :name_formula     => series[:name_formula]
        )
        index += 1
    end

    store_shtprops

    # Write the TEXT stream for each series.
    font_index = 5
    (0...@series.size).each do |i|
      store_defaulttext
      store_series_text_stream(font_index)
      font_index += 1
    end

    store_axesused(1)
    store_axisparent_stream

    if !@title_name.nil? || !@title_formula.nil?
      store_title_text_stream
    end

    store_end
  end

  def _formula_type_from_param(t, f, params, key)
    if params.has_key?(key)
      v = params[key]
      (v.nil? || v == [""] || v == '' || v == 0) ? f : t
    end
  end
  private :_formula_type_from_param

  ###############################################################################
  #
  # _store_series_stream()
  #
  # Write the SERIES chart substream.
  #
  def store_series_stream(params)
    name_type     = _formula_type_from_param(2, 1, params, :name_formula)
    value_type    = _formula_type_from_param(2, 0, params, :value_formula)
    category_type = _formula_type_from_param(2, 0, params, :category_formula)

    store_series(params[:value_count], params[:category_count])

    store_begin

    # Store the Series name AI record.
    store_ai(0, name_type, params[:name_formula])
    unless params[:name].nil?
      store_seriestext(params[:name], params[:name_encoding])
    end

    store_ai(1, value_type,    params[:value_formula])
    store_ai(2, category_type, params[:category_formula])
    store_ai(3, 1,             '' )

    store_dataformat_stream(params[:index])
    store_sertocrt
    store_end
  end

  ###############################################################################
  #
  # _store_dataformat_stream()
  #
  # Write the DATAFORMAT chart substream.
  #
  def store_dataformat_stream(series_index)
    store_dataformat(series_index, series_index, 0xFFFF)

    store_begin
    store_3dbarshape
    store_end
  end

  ###############################################################################
  #
  # _store_series_text_stream()
  #
  # Write the series TEXT substream.
  #
  def store_series_text_stream(font_index)
    store_text( 0xFFFFFF46, 0xFFFFFF06, 0, 0, 0x00B1, 0x1020 )

    store_begin
    store_pos( 2, 2, 0, 0, 0, 0 )
    store_fontx( font_index )
    store_ai( 0, 1, '' )
    store_end
  end

  def _formula_type(t, f, formula)
    (formula.nil? || formula == [""] || formula == '' || formula == 0) ? f : t
  end
  private :_formula_type

  ###############################################################################
  #
  # _store_x_axis_text_stream()
  #
  # Write the X-axis TEXT substream.
  #
  def store_x_axis_text_stream
    formula = @x_axis_formula.nil? ? '' : @x_axis_formula
    ai_type = _formula_type(2, 1, formula)

    store_text(0x07E1, 0x0DFC, 0xB2, 0x9C, 0x0081, 0x0000)

    store_begin
    store_pos(2, 2, 0, 0, 0x2B, 0x17)
    store_fontx(8)
    store_ai(0, ai_type, formula)

    unless @x_axis_name.nil?
      store_seriestext(@x_axis_name, @x_axis_encoding)
    end

    store_objectlink(3)
    store_end
  end

  ###############################################################################
  #
  # _store_y_axis_text_stream()
  #
  # Write the Y-axis TEXT substream.
  #
  def store_y_axis_text_stream
    formula = @y_axis_formula
    ai_type = _formula_type(2, 1, formula)

    store_text(0x002D, 0x06AA, 0x5F, 0x1CC, 0x0281, 0x00, 90)

    store_begin
    store_pos(2, 2, 0, 0, 0x17, 0x44)
    store_fontx(8)
    store_ai(0, ai_type, formula)

    unless @y_axis_name.nil?
      store_seriestext(@y_axis_name, @y_axis_encoding)
    end

    store_objectlink(2)
    store_end
  end

  ###############################################################################
  #
  # _store_legend_text_stream()
  #
  # Write the legend TEXT substream.
  #
  def store_legend_text_stream
    store_text(0xFFFFFF46, 0xFFFFFF06, 0, 0, 0x00B1, 0x0000)

    store_begin
    store_pos(2, 2, 0, 0, 0x00, 0x00)
    store_ai(0, 1, '')

    store_end
  end

  ###############################################################################
  #
  # _store_title_text_stream()
  #
  # Write the title TEXT substream.
  #
  def store_title_text_stream
    formula = @title_formula
    ai_type = _formula_type(2, 1, formula)

    store_text(0x06E4, 0x0051, 0x01DB, 0x00C4, 0x0081, 0x1030)

    store_begin
    store_pos(2, 2, 0, 0, 0x73, 0x1D)
    store_fontx(7)
    store_ai(0, ai_type, formula)

    unless @title_name.nil?
      store_seriestext(@title_name, @title_encoding)
    end

    store_objectlink(1)
    store_end
  end

  ###############################################################################
  #
  # _store_axisparent_stream()
  #
  # Write the AXISPARENT chart substream.
  #
  def store_axisparent_stream
    store_axisparent(0)

    store_begin
    store_pos(2, 2, 0x008C, 0x01AA, 0x0EEA, 0x0C52)
    store_axis_category_stream
    store_axis_values_stream

    if !@x_axis_name.nil? || !@x_axis_formula.nil?
      store_x_axis_text_stream
    end

    if !@y_axis_name.nil? || !@y_axis_formula.nil?
      store_y_axis_text_stream();
    end

    store_plotarea
    store_frame_stream
    store_chartformat_stream
    store_end
  end

  ###############################################################################
  #
  # _store_axis_category_stream()
  #
  # Write the AXIS chart substream for the chart category.
  #
  def store_axis_category_stream
    store_axis(0)

    store_begin
    store_catserrange
    store_axcext
    store_tick
    store_end
  end

  ###############################################################################
  #
  # _store_axis_values_stream()
  #
  # Write the AXIS chart substream for the chart values.
  #
  def store_axis_values_stream
    store_axis(1)

    store_begin
    store_valuerange
    store_tick
    store_axislineformat
    store_lineformat
    store_end
  end

  ###############################################################################
  #
  # _store_frame_stream()
  #
  # Write the FRAME chart substream.
  #
  def store_frame_stream
    store_frame

    store_begin
    store_lineformat
    store_areaformat
    store_end
  end

  ###############################################################################
  #
  # _store_chartformat_stream()
  #
  # Write the CHARTFORMAT chart substream.
  #
  def store_chartformat_stream
    store_chartformat

    store_begin

    # Store the BIFF record that will define the chart type.
    store_chart_type

    # CHARTFORMATLINK is not used.
    store_legend_stream
    store_end
  end

  ###############################################################################
  #
  # _store_chart_type()
  #
  # This is an abstract method that is overridden by the sub-classes to define
  # the chart types such as Column, Line, Pie, etc.
  #
  def store_chart_type

  end

  ###############################################################################
  #
  # _store_marker_dataformat_stream()
  #
  # This is an abstract method that is overridden by the sub-classes to define
  # properties of markers, linetypes, pie formats and other.
  #
  def store_marker_dataformat_stream

  end

  ###############################################################################
  #
  # _store_legend_stream()
  #
  # Write the LEGEND chart substream.
  #
  def store_legend_stream
    store_legend

    store_begin
    store_pos(5, 2, 0x05F9, 0x0EE9, 0, 0)
    store_legend_text_stream
    store_end
  end

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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
    append(header, data)
  end

  ###############################################################################
  #
  # _store_ai()
  #
  # Write the AI chart BIFF record.
  #
  def store_ai(id, type, formula, format_index = 0)
    formula = '' if formula == [""]

    record       = 0x1051     # Record identifier.
    length       = 0x0008     # Number of bytes to follow.
    # id                      # Link index.
    # type                    # Reference type.
    # formula                 # Pre-parsed formula.
    # format_index            # Num format index.
    grbit        = 0x0000     # Option flags.

    formula_length  = formula.length
    length += formula_length

    header = [record, length].pack('vv')
    data   = [id].pack('C')
    data  += [type].pack('C')
    data  += [grbit].pack('v')
    data  += [format_index].pack('v')
    data  += [formula_length].pack('v')
    data  += formula[0].kind_of?(String) ? formula[0] : formula

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
    append(header, data)
  end

  ###############################################################################
  #
  # _store_areaformat()
  #
  # Write the AREAFORMAT chart BIFF record. Contains the patterns and colours
  # of a chart area.
  #
  def store_areaformat
    record    = 0x100A     # Record identifier.
    length    = 0x0010     # Number of bytes to follow.
    rgbFore   = 0x00C0C0C0     # Foreground RGB colour.
    rgbBack   = 0x00000000     # Background RGB colour.
    pattern   = 0x0001         # Pattern.
    grbit     = 0x0000         # Option flags.
    indexFore = 0x0016         # Index to Foreground colour.
    indexBack = 0x004F         # Index to Background colour.

    header = [record, length].pack('vv')
    data  = [rgbFore].pack('V')
    data += [rgbBack].pack('V')
    data += [pattern].pack('v')
    data += [grbit].pack('v')
    data += [indexFore].pack('v')
    data += [indexBack].pack('v')

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
    append(header, data)
  end

  ###############################################################################
  #
  # _store_axis()
  #
  # Write the AXIS chart BIFF record to define the axis type.
  #
  def store_axis(type)
    record    = 0x101D         # Record identifier.
    length    = 0x0012         # Number of bytes to follow.
    # type                     # Axis type.
    reserved1 = 0x00000000     # Reserved.
    reserved2 = 0x00000000     # Reserved.
    reserved3 = 0x00000000     # Reserved.
    reserved4 = 0x00000000     # Reserved.

    header = [record, length].pack('vv')
    data  = [type].pack('v')
    data += [reserved1].pack('V')
    data += [reserved2].pack('V')
    data += [reserved3].pack('V')
    data += [reserved4].pack('V')

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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
    x      = 0x000000F8     # X-coord.
    y      = 0x000001F5     # Y-coord.
    dx     = 0x00000E7F     # Length of x axis.
    dy     = 0x00000B36     # Length of y axis.

    header = [record, length].pack('vv')
    data   = [iax].pack('v')
    data  += [x].pack('V')
    data  += [y].pack('V')
    data  += [dx].pack('V')
    data  += [dy].pack('V')

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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
    record   = 0x1002     # Record identifier.
    length   = 0x0010     # Number of bytes to follow.
    x_pos  = 0x00000000     # X pos of top left corner.
    y_pos  = 0x00000000     # Y pos of top left corner.
    dx     = 0x02DD51E0     # X size.
    dy     = 0x01C2B838     # Y size.

    header = [record, length].pack('vv')
    data   = [x_pos].pack('V')
    data  += [y_pos].pack('V')
    data  += [dx].pack('V')
    data  += [dy].pack('V')

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
    append(header, data)
  end

  ###############################################################################
  #
  # _store_chartformat()
  #
  # Write the CHARTFORMAT chart BIFF record. The parent record for formatting
  # of a chart group.
  #
  def store_chartformat
    record    = 0x1014         # Record identifier.
    length    = 0x0014         # Number of bytes to follow.
    reserved1 = 0x00000000     # Reserved.
    reserved2 = 0x00000000     # Reserved.
    reserved3 = 0x00000000     # Reserved.
    reserved4 = 0x00000000     # Reserved.
    grbit     = 0x0000         # Option flags.
    icrt      = 0x0000         # Drawing order.

    header = [record, length].pack('vv')
    data   = [reserved1].pack('V')
    data  += [reserved2].pack('V')
    data  += [reserved3].pack('V')
    data  += [reserved4].pack('V')
    data  += [grbit].pack('v')
    data  += [icrt].pack('v')

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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
    record = 0x1024     # Record identifier.
    length = 0x0002     # Number of bytes to follow.
    type   = 0x0002     # Type.

    header = [record, length].pack('vv')
    data  = [type].pack('v')

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
    append(header)
  end

  ###############################################################################
  # _store_fbi()
  #
  # Write the FBI chart BIFF record. Specifies the font information at the time
  # it was applied to the chart.
  #
  def store_fbi(index, height)
    record       = 0x1060    # Record identifier.
    length       = 0x000A    # Number of bytes to follow.
    # index                  # Font index.
    height       = height * 20    # Default font height in twips.
    width_basis  = 0x38B8    # Width basis, in twips.
    height_basis = 0x22A1    # Height basis, in twips.
    scale_basis  = 0x0000    # Scale by chart area or plot area.

    header = [record, length].pack('vv')
    data   = [width_basis].pack('v')
    data  += [height_basis].pack('v')
    data  += [height].pack('v')
    data  += [scale_basis].pack('v')
    data  += [index].pack('v')

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
    append(header, data)
  end

  ###############################################################################
  #
  # _store_frame()
  #
  # Write the FRAME chart BIFF record.
  #
  def store_frame
    record     = 0x1032     # Record identifier.
    length     = 0x0004     # Number of bytes to follow.
    frame_type = 0x0000     # Frame type.
    grbit      = 0x0003     # Option flags.

    header = [record, length].pack('vv')
    data  = [frame_type].pack('v')
    data += [grbit].pack('v')

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
    append(header, data)
  end

  ###############################################################################
  #
  # _store_legend()
  #
  # Write the LEGEND chart BIFF record. The Marcus Horan method.
  #
  def store_legend
    record   = 0x1015     # Record identifier.
    length   = 0x0014     # Number of bytes to follow.
    x        = 0x000005F9     # X-position.
    y        = 0x00000EE9     # Y-position.
    width    = 0x0000047D     # Width.
    height   = 0x0000009C     # Height.
    wType    = 0x00           # Type.
    wSpacing = 0x01           # Spacing.
    grbit    = 0x000F         # Option flags.

    header = [record, length].pack('vv')
    data  = [x].pack('V')
    data += [y].pack('V')
    data += [width].pack('V')
    data += [height].pack('V')
    data += [wType].pack('C')
    data += [wSpacing].pack('C')
    data += [grbit].pack('v')

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
    append(header, data)
  end

  ###############################################################################
  #
  # _store_lineformat()
  #
  # Write the LINEFORMAT chart BIFF record.
  #
  def store_lineformat
    record = 0x1007     # Record identifier.
    length = 0x000C     # Number of bytes to follow.
    rgb    = 0x00000000     # Line RGB colour.
    lns    = 0x0000         # Line pattern.
    we     = 0xFFFF         # Line weight.
    grbit  = 0x0009         # Option flags.
    index  = 0x004D         # Index to colour of line.

    header = [record, length].pack('vv')
    data  = [rgb].pack('V')
    data += [lns].pack('v')
    data += [we].pack('v')
    data += [grbit].pack('v')
    data += [index].pack('v')

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
    append(header)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    encoding ||= 0

    # Character length is num of chars not num of bytes
    cch /= 2 if encoding != 0

    # Change the UTF-16 name from BE to LE
    str = str.unpack('v*').pack('n*') if encoding != 0

    length = 4 + str.length

    header = [record, length].pack('vv')
    data  = [id].pack('v')
    data += [cch].pack('C')
    data += [encoding].pack('C')

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    header = [record, length].pack('vv')
    data  = [grbit].pack('v')
    data += [empty_cells].pack('v')

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
    append(header, data)
  end

  ###############################################################################
  #
  # _store_text()
  #
  # Write the TEXT chart BIFF record.
  #
  def store_text(x, y, dx, dy, grbit1, grbit2, rotation = 0x00)
    record   = 0x1025            # Record identifier.
    length   = 0x0020            # Number of bytes to follow.
    at       = 0x02              # Horizontal alignment.
    vat      = 0x02              # Vertical alignment.
    wBkgMode = 0x0001            # Background display.
    rgbText  = 0x0000            # Text RGB colour.
    # x                          # Text x-pos.
    # y                          # Text y-pos.
    # dx                         # Width.
    # dy                         # Height.
    # grbit1                     # Option flags.
    icvText  = 0x004D            # Auto Colour.
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
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

    print "sheet #{@name} : #{__FILE__}(#{__LINE__}) \n" if defined?($debug)
    append(header, data)
  end
end
