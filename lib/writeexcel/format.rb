##############################################################################
#
# Format - A class for defining Excel formatting.
#
#
# Used in conjunction with WriteExcel
#
# Copyright 2000-2008, John McNamara, jmcnamara@cpan.org
#
# original written in Perl by John McNamara
# converted to Ruby by Hideo Nakamura, cxn03651@msj.biglobe.ne.jp
#
require 'nkf'

#
# =CELL FORMATTING
#
# This section describes the methods and properties that are available for
# formatting cells in Excel. The properties of a cell that can be formatted
# include: fonts, colours, patterns, borders, alignment and number formatting.
#
# ==Creating and using a Format object
#
# Cell formatting is defined through a Format object. Format objects are
# created by calling the workbook add_format() method as follows:
#
#     format1 = workbook.add_format                   # Set properties later
#     format2 = workbook.add_format(property hash..)  # Set at creation
#
# The format object holds all the formatting properties that can be applied
# to a cell, a row or a column. The process of setting these properties is
# discussed in the next section.
#
# Once a Format object has been constructed and it properties have been set
# it can be passed as an argument to the worksheet write methods as follows:
#
#     worksheet.write(0, 0, 'One', format)
#     worksheet.write_string(1, 0, 'Two', format)
#     worksheet.write_number(2, 0, 3, format)
#     worksheet.write_blank(3, 0, format)
#
# Formats can also be passed to the worksheet set_row() and set_column()
# methods to define the default property for a row or column.
#
#     worksheet.set_row(0, 15, format)
#     worksheet.set_column(0, 0, 15, format)
#
# ==Format methods and Format properties
#
# The following table shows the Excel format categories, the formatting
# properties that can be applied and the equivalent object method:
#
#     Category   Description       Property        Method Name
#     --------   -----------       --------        -----------
#     Font       Font type         font            set_font()
#                Font size         size            set_size()
#                Font color        color           set_color()
#                Bold              bold            set_bold()
#                Italic            italic          set_italic()
#                Underline         underline       set_underline()
#                Strikeout         font_strikeout  set_font_strikeout()
#                Super/Subscript   font_script     set_font_script()
#                Outline           font_outline    set_font_outline()
#                Shadow            font_shadow     set_font_shadow()
#
#     Number     Numeric format    num_format      set_num_format()
#
#     Protection Lock cells        locked          set_locked()
#                Hide formulas     hidden          set_hidden()
#
#     Alignment  Horizontal align  align           set_align()
#                Vertical align    valign          set_align()
#                Rotation          rotation        set_rotation()
#                Text wrap         text_wrap       set_text_wrap()
#                Justify last      text_justlast   set_text_justlast()
#                Center across     center_across   set_center_across()
#                Indentation       indent          set_indent()
#                Shrink to fit     shrink          set_shrink()
#
#     Pattern    Cell pattern      pattern         set_pattern()
#                Background color  bg_color        set_bg_color()
#                Foreground color  fg_color        set_fg_color()
#
#     Border     Cell border       border          set_border()
#                Bottom border     bottom          set_bottom()
#                Top border        top             set_top()
#                Left border       left            set_left()
#                Right border      right           set_right()
#                Border color      border_color    set_border_color()
#                Bottom color      bottom_color    set_bottom_color()
#                Top color         top_color       set_top_color()
#                Left color        left_color      set_left_color()
#                Right color       right_color     set_right_color()
#
# There are two ways of setting Format properties: by using the object method
# interface or by setting the property directly. For example, a typical use of
# the method interface would be as follows:
#
#     format = workbook.add_format
#     format.set_bold
#     format.set_color('red')
#
# By comparison the properties can be set directly by passing a hash of
# properties to the Format constructor:
#
#     format = workbook.add_format(:bold => 1, :color => 'red')
#
# or after the Format has been constructed by means of the
# set_format_properties() method as follows:
#
#     format = workbook.add_format
#     format.set_format_properties(:bold => 1, :color => 'red')
#
# You can also store the properties in one or more named hashes and pass them
# to the required method:
#
#     font    = {
#                  :font  => 'Arial',
#                  :size  => 12,
#                  :color => 'blue',
#                  :bold  => 1
#               }
#
#     shading = {
#                  :bg_color => 'green',
#                  :pattern  => 1
#               }
#
#     format1 = workbook.add_format(font)           # Font only
#     format2 = workbook.add_format(font, shading)  # Font and shading
#
# The provision of two ways of setting properties might lead you to wonder
# which is the best way. The method mechanism may be better is you prefer
# setting properties via method calls (which the author did when they were
# code was first written) otherwise passing properties to the constructor has
# proved to be a little more flexible and self documenting in practice. An
# additional advantage of working with property hashes is that it allows you to
# share formatting between workbook objects as shown in the example above.
#
#--
#
# did not converted ????
#
# The Perl/Tk style of adding properties is also supported:
#
#     %font    = (
#                     -font      => 'Arial',
#                     -size      => 12,
#                     -color     => 'blue',
#                     -bold      => 1,
#                   )
#++
#
# ==Working with formats
#
# The default format is Arial 10 with all other properties off.
#
# Each unique format in Spreadsheet::WriteExcel must have a corresponding
# Format object. It isn't possible to use a Format with a write() method and
# then redefine the Format for use at a later stage. This is because a Format
# is applied to a cell not in its current state but in its final state.
# Consider the following example:
#
#     format = workbook.add_format
#     format.set_bold
#     format.set_color('red')
#     worksheet.write('A1', 'Cell A1', format)
#     format.set_color('green')
#     worksheet.write('B1', 'Cell B1', format)
#
# Cell A1 is assigned the Format _format_ which is initially set to the colour
# red. However, the colour is subsequently set to green. When Excel displays
# Cell A1 it will display the final state of the Format which in this case
# will be the colour green.
#
# In general a method call without an argument will turn a property on,
# for example:
#
#     format1 = workbook.add_format
#     format1.set_bold()   # Turns bold on
#     format1.set_bold(1)  # Also turns bold on
#     format1.set_bold(0)  # Turns bold off
#
# =FORMAT METHODS
#
# The Format object methods are described in more detail in the following
# sections. In addition, there is a Perl program called formats.rb in the
# examples directory of the WriteExcel distribution. This program creates an
# Excel workbook called formats.xls which contains examples of almost all
# the format types.
#
# The following Format methods are available:
#
#     set_font()
#     set_size()
#     set_color()
#     set_bold()
#     set_italic()
#     set_underline()
#     set_font_strikeout()
#     set_font_script()
#     set_font_outline()
#     set_font_shadow()
#     set_num_format()
#     set_locked()
#     set_hidden()
#     set_align()
#     set_rotation()
#     set_text_wrap()
#     set_text_justlast()
#     set_center_across()
#     set_indent()
#     set_shrink()
#     set_pattern()
#     set_bg_color()
#     set_fg_color()
#     set_border()
#     set_bottom()
#     set_top()
#     set_left()
#     set_right()
#     set_border_color()
#     set_bottom_color()
#     set_top_color()
#     set_left_color()
#     set_right_color()
#
# The above methods can also be applied directly as properties. For example
# format.set_bold() is equivalent to workbook.add_format(bold => 1).
#
# =COLOURS IN EXCEL
#
# Excel provides a colour palette of 56 colours. In WriteExcel these colours
# are accessed via their palette index in the range 8..63. This index is used
# to set the colour of fonts, cell patterns and cell borders. For example:
#
#     format = workbook.add_format(
#                   :color => 12, # index for blue
#                   :font  => 'Arial',
#                   :size  => 12,
#                   :bold  => 1
#                 )
#
# The most commonly used colours can also be accessed by name. The name acts
# as a simple alias for the colour index:
#
#     :black     =>    8
#     :blue      =>   12
#     :brown     =>   16
#     :cyan      =>   15
#     :gray      =>   23
#     :green     =>   17
#     :lime      =>   11
#     :magenta   =>   14
#     :navy      =>   18
#     :orange    =>   53
#     :pink      =>   33
#     :purple    =>   20
#     :red       =>   10
#     :silver    =>   22
#     :white     =>    9
#     :yellow    =>   13
#
# For example:
#
#     font = workbook.add_format(:color => 'red')
#
# Users of VBA in Excel should note that the equivalent colour indices are in
# the range 1..56 instead of 8..63.
#
# If the default palette does not provide a required colour you can override
# one of the built-in values. This is achieved by using the set_custom_color()
# workbook method to adjust the RGB (red green blue) components of the colour:
#
#     ferrari = workbook.set_custom_color(40, 216, 12, 12)
#
#     format  = workbook.add_format(
#                   :bg_color => ferrari,
#                   :pattern  => 1,
#                   :border   => 1
#                 )
#
#     worksheet.write_blank('A1', format)
#--
# The default Excel 97 colour palette is shown in palette.html in the doc
# directory of the distro. You can generate an Excel version of the palette
# using colors.pl in the examples directory.
#
# A comparison of the colour components in the Excel 5 and Excel 97+ colour
# palettes is shown in rgb5-97.txt in the doc directory.
#++
#
# You may also find the following links helpful:
#
# A detailed look at Excel's colour palette:
# http://www.mvps.org/dmcritchie/excel/colors.htm
#
# A decimal RGB chart: http://www.hypersolutions.org/pages/rgbdec.html
#
# A hex RGB chart: : http://www.hypersolutions.org/pages/rgbhex.html
#
# =DATES AND TIME IN EXCEL
#
# There are two important things to understand about dates and times in Excel:
#
# 1 A date/time in Excel is a real number plus an Excel number format.
#
# 2 WriteExcel doesn't automatically convert date/time strings in write() to
# an Excel date/time.
#
# These two points are explained in more detail below along with some
# suggestions on how to convert times and dates to the required format.
# An Excel date/time is a number plus a format
#
# If you write a date string with write() then all you will get is a string:
#
#     worksheet.write('A1', '02/03/04')  # !! Writes a string not a date. !!
#
# Dates and times in Excel are represented by real numbers, for example
# "Jan 1 2001 12:30 AM" is represented by the number 36892.521.
#
# The integer part of the number stores the number of days since the epoch
# and the fractional part stores the percentage of the day.
#
# A date or time in Excel is just like any other number. To have the number
# display as a date you must apply an Excel number format to it. Here are
# some examples.
#
#     #!/usr/bin/ruby -w
#
#     require 'writeexcel'
#
#     workbook  = WriteExcel.new('date_examples.xls')
#     worksheet = workbook.add_worksheet
#
#     worksheet.set_column('A:A', 30)  # For extra visibility.
#
#     number    = 39506.5
#
#     worksheet.write('A1', number)            #     39506.5
#
#     format2 = workbook.add_format(num_format => 'dd/mm/yy')
#     worksheet.write('A2', number , format2); #     28/02/08
#
#     format3 = workbook.add_format(num_format => 'mm/dd/yy')
#     worksheet.write('A3', number , format3); #     02/28/08
#
#     format4 = workbook.add_format(num_format => 'd-m-yyyy')
#     worksheet.write('A4', .number , format4) #     28-2-2008
#
#     format5 = workbook.add_format(num_format => 'dd/mm/yy hh:mm')
#     worksheet.write('A5', number , format5)  #     28/02/08 12:00
#
#     format6 = workbook.add_format(num_format => 'd mmm yyyy')
#     worksheet.write('A6', number , format6)  #     28 Feb 2008
#
#     format7 = workbook.add_format(num_format => 'mmm d yyyy hh:mm AM/PM')
#     worksheet.write('A7', number , format7)  #     Feb 28 2008 12:00 PM
#
# WriteExcel doesn't automatically convert date/time strings
#
# WriteExcel doesn't automatically convert input date strings into Excel's
# formatted date numbers due to the large number of possible date formats
# and also due to the possibility of misinterpretation.
#
# For example, does 02/03/04 mean March 2 2004, February 3 2004 or even March
# 4 2002.
#
# Therefore, in order to handle dates you will have to convert them to numbers
# and apply an Excel format. Some methods for converting dates are listed in
# the next section.
#
# The most direct way is to convert your dates to the ISO8601
# yyyy-mm-ddThh:mm:ss.sss date format and use the write_date_time() worksheet
# method:
#
#     worksheet.write_date_time('A2', '2001-01-01T12:20', format)
#
# See the write_date_time() section of the documentation for more details.
#
# A general methodology for handling date strings with write_date_time() is:
#
#     1. Identify incoming date/time strings with a regex.
#     2. Extract the component parts of the date/time using the same regex.
#     3. Convert the date/time to the ISO8601 format.
#     4. Write the date/time using write_date_time() and a number format.
#
# Here is an example:
#
#     #!/usr/bin/ruby -w
#
#     require 'writeexcel'
#
#     workbook    = WriteExcel.new('example.xls')
#     worksheet   = workbook.add_worksheet
#
#     # Set the default format for dates.
#     date_format = workbook.add_format(num_format => 'mmm d yyyy')
#
#     # Increase column width to improve visibility of data.
#     worksheet.set_column('A:C', 20)
#
#     # Simulate reading from a data source.
#     row = 0
#
#     while (<DATA>) {
#         chomp;
#
#         my $col  = 0;
#         my @data = split ' ';
#
#         for my $item (@data) {
#
#             # Match dates in the following formats: d/m/yy, d/m/yyyy
#             if ($item =~ qr[^(\d{1,2})/(\d{1,2})/(\d{4})$]) {
#
#                 # Change to the date format required by write_date_time().
#                 my $date = sprintf "%4d-%02d-%02dT", $3, $2, $1;
#
#                 $worksheet->write_date_time($row, $col++, $date, $date_format);
#             }
#             else {
#                 # Just plain data
#                 $worksheet->write($row, $col++, $item);
#             }
#         }
#         $row++;
#     }
#
#     __DATA__
#     Item    Cost    Date
#     Book    10      1/9/2007
#     Beer    4       12/9/2007
#     Bed     500     5/10/2007
#
# For a slightly more advanced solution you can modify the write() method to
# handle date formats of your choice via the add_write_handler() method. See
# the add_write_handler() section of the docs and the write_handler3.rb and
# write_handler4.rb programs in the examples directory of the distro.
#
# =OUTLINES AND GROUPING IN EXCEL
#
# Excel allows you to group rows or columns so that they can be hidden or
# displayed with a single mouse click. This feature is referred to as outlines.
#
# Outlines can reduce complex data down to a few salient sub-totals or
# summaries.
#
# This feature is best viewed in Excel but the following is an ASCII
# representation of what a worksheet with three outlines might look like. Rows
# 3-4 and rows 7-8 are grouped at level 2. Rows 2-9 are grouped at level 1.
# The lines at the left hand side are called outline level bars.
#
#             ------------------------------------------
#      1 2 3 |   |   A   |   B   |   C   |   D   |  ...
#             ------------------------------------------
#       _    | 1 |   A   |       |       |       |  ...
#      |  _  | 2 |   B   |       |       |       |  ...
#      | |   | 3 |  (C)  |       |       |       |  ...
#      | |   | 4 |  (D)  |       |       |       |  ...
#      | -   | 5 |   E   |       |       |       |  ...
#      |  _  | 6 |   F   |       |       |       |  ...
#      | |   | 7 |  (G)  |       |       |       |  ...
#      | |   | 8 |  (H)  |       |       |       |  ...
#      | -   | 9 |   I   |       |       |       |  ...
#      -     | . |  ...  |  ...  |  ...  |  ...  |  ...
#
# Clicking the minus sign on each of the level 2 outlines will collapse and
# hide the data as shown in the next figure. The minus sign changes to a plus
# sign to indicate that the data in the outline is hidden.
#
#             ------------------------------------------
#      1 2 3 |   |   A   |   B   |   C   |   D   |  ...
#             ------------------------------------------
#       _    | 1 |   A   |       |       |       |  ...
#      |     | 2 |   B   |       |       |       |  ...
#      | +   | 5 |   E   |       |       |       |  ...
#      |     | 6 |   F   |       |       |       |  ...
#      | +   | 9 |   I   |       |       |       |  ...
#      -     | . |  ...  |  ...  |  ...  |  ...  |  ...
#
# Clicking on the minus sign on the level 1 outline will collapse the
# remaining rows as follows:
#
#             ------------------------------------------
#      1 2 3 |   |   A   |   B   |   C   |   D   |  ...
#             ------------------------------------------
#            | 1 |   A   |       |       |       |  ...
#      +     | . |  ...  |  ...  |  ...  |  ...  |  ...
#
# Grouping in WriteExcel is achieved by setting the outline level via the
# set_row() and set_column() worksheet methods:
#
#     set_row(row, height, format, hidden, level, collapsed)
#     set_column(first_col, last_col, width, format, hidden, level, collapsed)
#
# The following example sets an outline level of 1 for rows 1 and 2
# (zero-indexed) and columns B to G. The parameters _height_ and _XF_ are
# assigned default values since they are undefined:
#
#     worksheet.set_row(1, nil, nil, 0, 1)
#     worksheet.set_row(2, nil, nil, 0, 1)
#     worksheet.set_column('B:G', nil, nil, 0, 1)
#
# Excel allows up to 7 outline levels. Therefore the _level_ parameter should
# be in the range 0 <= _level_ <= 7.
#
# Rows and columns can be collapsed by setting the _hidden_ flag for the hidden
# rows/columns and setting the _collapsed_ flag for the row/column that has
# the collapsed + symbol:
#
#     worksheet.set_row(1, nil, nil, 1, 1)
#     worksheet.set_row(2, nil, nil, 1, 1)
#     worksheet.set_row(3, nil, nil, 0, 0, 1)         # Collapsed flag.
#
#     worksheet.set_column('B:G', nil, nil, 1, 1)
#     worksheet.set_column('H:H', nil, nil, 0, 0, 1)  # Collapsed flag.
#
# Note: Setting the _collapsed_ flag is particularly important for
# compatibility with OpenOffice.org and Gnumeric.
#
# For a more complete example see the outline.pl and outline_collapsed.rb
# programs in the examples directory of the distro.
#
# Some additional outline properties can be set via the outline_settings()
# worksheet method, see above.
#
class Format

  COLORS = {
    'aqua'    => 0x0F,
    'cyan'    => 0x0F,
    'black'   => 0x08,
    'blue'    => 0x0C,
    'brown'   => 0x10,
    'magenta' => 0x0E,
    'fuchsia' => 0x0E,
    'gray'    => 0x17,
    'grey'    => 0x17,
    'green'   => 0x11,
    'lime'    => 0x0B,
    'navy'    => 0x12,
    'orange'  => 0x35,
    'pink'    => 0x21,
    'purple'  => 0x14,
    'red'     => 0x0A,
    'silver'  => 0x16,
    'white'   => 0x09,
    'yellow'  => 0x0D,
  }
  NonAscii = /[^!"#\$%&'\(\)\*\+,\-\.\/\:\;<=>\?@0-9A-Za-z_\[\\\]^` ~\0\n]/

  attr_accessor :xf_index, :used_merge
  attr_accessor :bold, :text_wrap, :text_justlast
  attr_accessor :fg_color, :bg_color, :color, :font_outline, :font_shadow
  attr_accessor :align, :border
  attr_accessor :font_index
  attr_accessor :num_format
  attr_reader   :type
  attr_reader   :font, :size, :font_family, :font_strikeout, :font_script, :font_charset
  attr_reader   :font_encoding, :merge_range, :reading_order
  attr_reader   :diag_type, :diag_color, :diag_border
  attr_reader   :num_format_enc, :locked, :hidden
  attr_reader   :rotation, :indent, :shrink, :pattern, :bottom, :top, :left, :right
  attr_reader   :bottom_color, :top_color, :left_color, :right_color
  attr_reader   :italic, :underline, :font_strikeout
  attr_reader   :text_h_align, :text_v_align, :font_only

  ###############################################################################
  #
  # initialize(xf_index=0, properties = {})
  #    xf_index   :
  #    properties : Hash of property => value
  #
  # Constructor
  #
  def initialize(xf_index = 0, properties = {})
    @xf_index       = xf_index

    @type           = 0
    @font_index     = 0
    @font           = 'Arial'
    @size           = 10
    @bold           = 0x0190
    @italic         = 0
    @color          = 0x7FFF
    @underline      = 0
    @font_strikeout = 0
    @font_outline   = 0
    @font_shadow    = 0
    @font_script    = 0
    @font_family    = 0
    @font_charset   = 0
    @font_encoding  = 0

    @num_format     = 0
    @num_format_enc = 0

    @hidden         = 0
    @locked         = 1

    @text_h_align   = 0
    @text_wrap      = 0
    @text_v_align   = 2
    @text_justlast  = 0
    @rotation       = 0

    @fg_color       = 0x40
    @bg_color       = 0x41

    @pattern        = 0

    @bottom         = 0
    @top            = 0
    @left           = 0
    @right          = 0

    @bottom_color   = 0x40
    @top_color      = 0x40
    @left_color     = 0x40
    @right_color    = 0x40

    @indent         = 0
    @shrink         = 0
    @merge_range    = 0
    @reading_order  = 0

    @diag_type      = 0
    @diag_color     = 0x40
    @diag_border    = 0

    @font_only      = 0

    # Temp code to prevent merged formats in non-merged cells.
    @used_merge     = 0

    set_format_properties(properties) unless properties.empty?
  end


  #
  # :call-seq:
  #    copy(format)
  #
  # Copy the attributes of another Format object.
  #
  # This method is used to copy all of the properties from one Format object
  # to another:
  #
  #     lorry1 = workbook.add_format
  #     lorry1.set_bold
  #     lorry1.set_italic
  #     lorry1.set_color('red')     # lorry1 is bold, italic and red
  #
  #     lorry2 = workbook.add_format
  #     lorry2.copy(lorry1)
  #     lorry2.set_color('yellow')  # lorry2 is bold, italic and yellow
  #
  # The copy() method is only useful if you are using the method interface
  # to Format properties. It generally isn't required if you are setting
  # Format properties directly using hashes.
  #
  # Note: this is not a copy constructor, both objects must exist prior to
  # copying.
  #
  def copy(other)
    return unless other.kind_of?(Format)

    # copy properties except xf, merge_range, used_merge
    # Copy properties
    @type           = other.type
    @font_index     = other.font_index
    @font           = other.font
    @size           = other.size
    @bold           = other.bold
    @italic         = other.italic
    @color          = other.color
    @underline      = other.underline
    @font_strikeout = other.font_strikeout
    @font_outline   = other.font_outline
    @font_shadow    = other.font_shadow
    @font_script    = other.font_script
    @font_family    = other.font_family
    @font_charset   = other.font_charset
    @font_encoding  = other.font_encoding

    @num_format     = other.num_format
    @num_format_enc = other.num_format_enc

    @hidden         = other.hidden
    @locked         = other.locked

    @text_h_align   = other.text_h_align
    @text_wrap      = other.text_wrap
    @text_v_align   = other.text_v_align
    @text_justlast  = other.text_justlast
    @rotation       = other.rotation

    @fg_color       = other.fg_color
    @bg_color       = other.bg_color

    @pattern        = other.pattern

    @bottom         = other.bottom
    @top            = other.top
    @left           = other.left
    @right          = other.right

    @bottom_color   = other.bottom_color
    @top_color      = other.top_color
    @left_color     = other.left_color
    @right_color    = other.right_color

    @indent         = other.indent
    @shrink         = other.shrink
    @reading_order  = other.reading_order

    @diag_type      = other.diag_type
    @diag_color     = other.diag_color
    @diag_border    = other.diag_border

    @font_only      = other.font_only
end

  ###############################################################################
  #
  # get_xf($style)
  #
  # Generate an Excel BIFF XF record.
  #
  def get_xf

    # Local Variable
    #    record;     # Record identifier
    #    length;     # Number of bytes to follow
    #
    #    ifnt;       # Index to FONT record
    #    ifmt;       # Index to FORMAT record
    #    style;      # Style and other options
    #    align;      # Alignment
    #    indent;     #
    #    icv;        # fg and bg pattern colors
    #    border1;    # Border line options
    #    border2;    # Border line options
    #    border3;    # Border line options

    # Set the type of the XF record and some of the attributes.
    if @type == 0xFFF5 then
      style = 0xFFF5
    else
      style  = @locked
      style |= @hidden << 1
    end

    # Flags to indicate if attributes have been set.
    atr_num  = (@num_format   != 0) ? 1 : 0
    atr_fnt  = (@font_index   != 0) ? 1 : 0
    atr_alc  = (@text_h_align != 0 ||
                @text_v_align != 2 ||
                @shrink       != 0 ||
                @merge_range  != 0 ||
                @text_wrap    != 0 ||
                @indent       != 0) ? 1 : 0
    atr_bdr  = (@bottom       != 0 ||
                @top          != 0 ||
                @left         != 0 ||
                @right        != 0 ||
                @diag_type    != 0) ? 1 : 0
    atr_pat  = (@fg_color     != 0x40 ||
                @bg_color     != 0x41 ||
                @pattern      != 0x00) ? 1 : 0
    atr_prot = (@hidden       != 0 ||
                @locked       != 1) ? 1 : 0

    # Set attribute changed flags for the style formats.
    if @xf_index != 0 and @type == 0xFFF5
      if @xf_index >= 16
        atr_num    = 0
        atr_fnt    = 1
      else
        atr_num    = 1
        atr_fnt    = 0
      end
      atr_alc    = 1
      atr_bdr    = 1
      atr_pat    = 1
      atr_prot   = 1
    end

    # Set a default diagonal border style if none was specified.
    @diag_border = 1 if (@diag_border ==0 and @diag_type != 0)

    # Reset the default colours for the non-font properties
    @fg_color     = 0x40 if @fg_color     == 0x7FFF
    @bg_color     = 0x41 if @bg_color     == 0x7FFF
    @bottom_color = 0x40 if @bottom_color == 0x7FFF
    @top_color    = 0x40 if @top_color    == 0x7FFF
    @left_color   = 0x40 if @left_color   == 0x7FFF
    @right_color  = 0x40 if @right_color  == 0x7FFF
    @diag_color   = 0x40 if @diag_color   == 0x7FFF

    # Zero the default border colour if the border has not been set.
    @bottom_color = 0 if @bottom    == 0
    @top_color    = 0 if @top       == 0
    @right_color  = 0 if @right     == 0
    @left_color   = 0 if @left      == 0
    @diag_color   = 0 if @diag_type == 0

    # The following 2 logical statements take care of special cases in relation
    # to cell colours and patterns:
    # 1. For a solid fill (_pattern == 1) Excel reverses the role of foreground
    #    and background colours.
    # 2. If the user specifies a foreground or background colour without a
    #    pattern they probably wanted a solid fill, so we fill in the defaults.
    #
    if (@pattern  <= 0x01 && @bg_color != 0x41 && @fg_color == 0x40)
      @fg_color = @bg_color
      @bg_color = 0x40
      @pattern  = 1
    end

    if (@pattern <= 0x01 && @bg_color == 0x41 && @fg_color != 0x40)
      @bg_color = 0x40
      @pattern  = 1
    end

    # Set default alignment if indent is set.
    @text_h_align = 1 if @indent != 0 and @text_h_align == 0


    record         = 0x00E0
    length         = 0x0014

    ifnt           = @font_index
    ifmt           = @num_format


    align          = @text_h_align
    align         |= @text_wrap     << 3
    align         |= @text_v_align  << 4
    align         |= @text_justlast << 7
    align         |= @rotation      << 8

    indent         = @indent
    indent        |= @shrink        << 4
    indent        |= @merge_range   << 5
    indent        |= @reading_order << 6
    indent        |= atr_num        << 10
    indent        |= atr_fnt        << 11
    indent        |= atr_alc        << 12
    indent        |= atr_bdr        << 13
    indent        |= atr_pat        << 14
    indent        |= atr_prot       << 15


    border1        = @left
    border1       |= @right         << 4
    border1       |= @top           << 8
    border1       |= @bottom        << 12

    border2        = @left_color
    border2       |= @right_color   << 7
    border2       |= @diag_type     << 14

    border3        = 0
    border3       |= @top_color
    border3       |= @bottom_color  << 7
    border3       |= @diag_color    << 14
    border3       |= @diag_border   << 21
    border3       |= @pattern       << 26

    icv            = @fg_color
    icv           |= @bg_color      << 7

    header = [record, length].pack("vv")
    data   = [ifnt, ifmt, style, align, indent,
              border1, border2, border3, icv].pack("vvvvvvvVv")

    return header + data
  end

  ###############################################################################
  #
  # get_font()
  #
  # Generate an Excel BIFF FONT record.
  #
  def get_font

    #   my $record;     # Record identifier
    #   my $length;     # Record length

    #   my $dyHeight;   # Height of font (1/20 of a point)
    #   my $grbit;      # Font attributes
    #   my $icv;        # Index to color palette
    #   my $bls;        # Bold style
    #   my $sss;        # Superscript/subscript
    #   my $uls;        # Underline
    #   my $bFamily;    # Font family
    #   my $bCharSet;   # Character set
    #   my $reserved;   # Reserved
    #   my $cch;        # Length of font name
    #   my $rgch;       # Font name
    #   my $encoding;   # Font name character encoding


    dyHeight   = @size * 20
    icv        = @color
    bls        = @bold
    sss        = @font_script
    uls        = @underline
    bFamily    = @font_family
    bCharSet   = @font_charset
    rgch       = @font
    encoding   = @font_encoding

    # Handle utf8 strings
    if rgch =~ NonAscii
      rgch = NKF.nkf('-w16B0 -m0 -W', rgch)
      encoding = 1
    end

    cch = rgch.length
    #
    # Handle Unicode font names.
    if (encoding == 1)
      raise "Uneven number of bytes in Unicode font name" if cch % 2 != 0
      cch  /= 2 if encoding !=0
      rgch  = rgch.unpack('n*').pack('v*')
    end

    record     = 0x31
    length     = 0x10 + rgch.length
    reserved   = 0x00

    grbit      = 0x00
    grbit     |= 0x02 if @italic != 0
    grbit     |= 0x08 if @font_strikeout != 0
    grbit     |= 0x10 if @font_outline != 0
    grbit     |= 0x20 if @font_shadow != 0


    header = [record, length].pack("vv")
    data   = [dyHeight, grbit, icv, bls,
              sss, uls, bFamily,
              bCharSet, reserved, cch, encoding].pack('vvvvvCCCCCC')

    return header + data + rgch
  end

  ###############################################################################
  #
  # get_font_key()
  #
  # Returns a unique hash key for a font. Used by Workbook->_store_all_fonts()
  #
  def get_font_key
    # The following elements are arranged to increase the probability of
    # generating a unique key. Elements that hold a large range of numbers
    # e.g. _color are placed between two binary elements such as _italic

    key  = "#{@font}#{@size}#{@font_script}#{@underline}#{@font_strikeout}#{@bold}#{@font_outline}"
    key += "#{@font_family}#{@font_charset}#{@font_shadow}#{@color}#{@italic}#{@font_encoding}"
    result =  key.gsub(' ', '_') # Convert the key to a single word

    return result
  end

  ###############################################################################
  #
  # get_xf_index()
  #
  # Returns the used by Worksheet->_XF()
  #
  def get_xf_index
    return @xf_index
  end


  ###############################################################################
  #
  # get_color(colour)
  #
  # Used in conjunction with the set_xxx_color methods to convert a color
  # string into a number. Color range is 0..63 but we will restrict it
  # to 8..63 to comply with Gnumeric. Colors 0..7 are repeated in 8..15.
  #
  def get_color(colour = nil)
    # Return the default color, 0x7FFF, if undef,
    return 0x7FFF if colour.nil?

    if colour.kind_of?(Numeric)
      if colour < 0
        return 0x7FFF

      # or an index < 8 mapped into the correct range,
      elsif colour < 8
        return (colour + 8).to_i

      # or the default color if arg is outside range,
      elsif colour > 63
        return 0x7FFF

      # or an integer in the valid range
      else
        return colour.to_i
      end
    elsif colour.kind_of?(String)
      # or the color string converted to an integer,
      if COLORS.has_key?(colour)
        return COLORS[colour]

      # or the default color if string is unrecognised,
      else
        return 0x7FFF
      end
    else
      return 0x7FFF
    end
  end

  ###############################################################################
  #
  # class method    Format._get_color(colour)
  #
  #  used from Worksheet.rb
  #
  #  this is cut & copy of get_color().
  #
  def self._get_color(colour)
    # Return the default color, 0x7FFF, if undef,
    return 0x7FFF if colour.nil?

    if colour.kind_of?(Numeric)
      if colour < 0
        return 0x7FFF

        # or an index < 8 mapped into the correct range,
      elsif colour < 8
        return (colour + 8).to_i

        # or the default color if arg is outside range,
      elsif 63 < colour
        return 0x7FFF

        # or an integer in the valid range
      else
        return colour.to_i
      end
    elsif colour.kind_of?(String)
      # or the color string converted to an integer,
      if COLORS.has_key?(colour)
        return COLORS[colour]

        # or the default color if string is unrecognised,
      else
        return 0x7FFF
      end
    else
      return 0x7FFF
    end
  end

  ###############################################################################
  #
  # set_type()
  #
  # Set the XF object type as 0 = cell XF or 0xFFF5 = style XF.
  #
  def set_type(type = nil)

    if !type.nil? and type == 0
      @type = 0x0000
    else
      @type = 0xFFF5
    end
  end

  #
  #     Default state:      Font size is 10
  #     Default action:     Set font size to 1
  #     Valid args:         Integer values from 1 to as big as your screen.
  #
  # Set the font size. Excel adjusts the height of a row to accommodate the
  # largest font size in the row. You can also explicitly specify the height
  # of a row using the set_row() worksheet method.
  #
  #     format = workbook.add_format
  #     format.set_size(30)
  #
  def set_size(size = 1)
    if size.kind_of?(Numeric) && size >= 1
      @size = size.to_i
    end
  end

  #
  # Set the font colour.
  #
  #    Default state:      Excels default color, usually black
  #    Default action:     Set the default color
  #    Valid args:         Integers from 8..63 or the following strings:
  #                        'black', 'blue', 'brown', 'cyan', 'gray'
  #                        'green', 'lime', 'magenta', 'navy', 'orange'
  #                        'pink', 'purple', 'red', 'silver', 'white', 'yellow'
  #
  # The set_color() method is used as follows:
  #
  #    format = workbook.add_format()
  #    format.set_color('red')
  #    worksheet.write(0, 0, 'wheelbarrow', format)
  #
  # Note: The set_color() method is used to set the colour of the font in a cell.
  #       To set the colour of a cell use the set_bg_color()
  #       and set_pattern() methods.
  #
  def set_color(color = 0x7FFF)
    @color = get_color(color)
  end

  #
  # Set the italic property of the font:
  #
  #     Default state:      Italic is off
  #     Default action:     Turn italic on
  #     Valid args:         0, 1
  #
  #     format.set_italic    # Turn italic on
  #
  def set_italic(arg = 1)
    begin
      if    arg == 1  then @italic = 1   # italic on
      elsif arg == 0  then @italic = 0   # italic off
      else
        raise ArgumentError,
        "\n\n  set_italic(#{arg.inspect})\n    arg must be 0, 1, or none. ( 0:OFF , 1 and none:ON )\n"
      end
    end
  end

  #
  # Set the underline property of the font.
  #
  #     Default state:      Underline is off
  #     Default action:     Turn on single underline
  #     Valid args:         0  = No underline
  #                         1  = Single underline
  #                         2  = Double underline
  #                         33 = Single accounting underline
  #                         34 = Double accounting underline
  #
  #     format.set_underline();   # Single underline
  #
  def set_underline(arg = 1)
    begin
      case arg
      when  0  then @underline =  0    # off
      when  1  then @underline =  1    # Single
      when  2  then @underline =  2    # Double
      when 33  then @underline = 33    # Single accounting
      when 34  then @underline = 34    # Double accounting
      else
        raise ArgumentError,
        "\n\n  set_underline(#{arg.inspect})\n    arg must be 0, 1, or none, 2, 33, 34.\n"
        " ( 0:OFF, 1 and none:Single, 2:Double, 33:Single accounting, 34:Double accounting )\n"
      end
    end
  end

  #
  # Set the strikeout property of the font.
  #
  #     Default state:      Strikeout is off
  #     Default action:     Turn strikeout on
  #     Valid args:         0, 1
  #
  def set_font_strikeout(arg = 1)
    begin
      if    arg == 0 then @font_strikeout = 0
      elsif arg == 1 then @font_strikeout = 1
      else
        raise ArgumentError,
        "\n\n  set_font_strikeout(#{arg.inspect})\n    arg must be 0, 1, or none.\n"
        " ( 0:OFF, 1 and none:Strikeout )\n"
      end
    end
  end

  #
  # Set the superscript/subscript property of the font.
  # This format is currently not very useful.
  #
  #     Default state:      Super/Subscript is off
  #     Default action:     Turn Superscript on
  #     Valid args:         0  = Normal
  #                         1  = Superscript
  #                         2  = Subscript
  #
  def set_font_script(arg = 1)
    begin
      if    arg == 0 then @font_script = 0
      elsif arg == 1 then @font_script = 1
      elsif arg == 2 then @font_script = 2
      else
        raise ArgumentError,
        "\n\n  set_font_script(#{arg.inspect})\n    arg must be 0, 1, or none. or 2\n"
        " ( 0:OFF, 1 and none:Superscript, 2:Subscript )\n"
      end
    end
  end

  #
  # Macintosh only.
  #
  #     Default state:      Outline is off
  #     Default action:     Turn outline on
  #     Valid args:         0, 1
  #
  def set_font_outline(arg = 1)
    begin
      if    arg == 0 then @font_outline = 0
      elsif arg == 1 then @font_outline = 1
      else
        raise ArgumentError,
        "\n\n  set_font_outline(#{arg.inspect})\n    arg must be 0, 1, or none.\n"
        " ( 0:OFF, 1 and none:outline on )\n"
      end
    end
  end

  #
  # Macintosh only.
  #
  #     Default state:      Shadow is off
  #     Default action:     Turn shadow on
  #     Valid args:         0, 1
  #
  def set_font_shadow(arg = 1)
    begin
      if    arg == 0 then @font_shadow = 0
      elsif arg == 1 then @font_shadow = 1
      else
        raise ArgumentError,
        "\n\n  set_font_shadow(#{arg.inspect})\n    arg must be 0, 1, or none.\n"
        " ( 0:OFF, 1 and none:shadow on )\n"
      end
    end
  end

  #
  # prevent modification of a cells contents.
  #
  #     Default state:      Cell locking is on
  #     Default action:     Turn locking on
  #     Valid args:         0, 1
  #
  # This property can be used to prevent modification of a cells contents.
  # Following Excel's convention, cell locking is turned on by default.
  # However, it only has an effect if the worksheet has been protected,
  # see the worksheet protect() method.
  #
  #     locked  = workbook.add_format()
  #     locked.set_locked(1) # A non-op
  #
  #     unlocked = workbook.add_format()
  #     locked.set_locked(0)
  #
  #     # Enable worksheet protection
  #     worksheet.protect()
  #
  #     # This cell cannot be edited.
  #     worksheet.write('A1', '=1+2', locked)
  #
  #     # This cell can be edited.
  #     worksheet.write('A2', '=1+2', unlocked)
  #
  # Note: This offers weak protection even with a password, see the note
  # in relation to the protect() method.
  #
  def set_locked(arg = 1)
    begin
      if    arg == 0 then @locked = 0
      elsif arg == 1 then @locked = 1
      else
        raise ArgumentError,
        "\n\n  set_locked(#{arg.inspect})\n    arg must be 0, 1, or none.\n"
        " ( 0:OFF, 1 and none:Lock On )\n"
      end
    end
  end

  #
  # hide a formula while still displaying its result.
  #
  #     Default state:      Formula hiding is off
  #     Default action:     Turn hiding on
  #     Valid args:         0, 1
  #
  # This property is used to hide a formula while still displaying
  # its result. This is generally used to hide complex calculations
  # from end users who are only interested in the result. It only has
  # an effect if the worksheet has been protected,
  # see the worksheet protect() method.
  #
  #     hidden = workbook.add_format
  #     hidden.set_hidden
  #
  #     # Enable worksheet protection
  #     worksheet.protect
  #
  #     # The formula in this cell isn't visible
  #     worksheet.write('A1', '=1+2', hidden)
  #
  # Note: This offers weak protection even with a password,
  #       see the note in relation to the protect() method  .
  #
  def set_hidden(arg = 1)
    begin
      if    arg == 0 then @hidden = 0
      elsif arg == 1 then @hidden = 1
      else
        raise ArgumentError,
        "\n\n  set_hidden(#{arg.inspect})\n    arg must be 0, 1, or none.\n"
        " ( 0:OFF, 1 and none:hiding On )\n"
      end
    end
  end

  #
  # Set cell alignment.
  #
  #     Default state:      Alignment is off
  #     Default action:     Left alignment
  #     Valid args:         'left'              Horizontal
  #                         'center'
  #                         'right'
  #                         'fill'
  #                         'justify'
  #                         'center_across'
  #
  #                         'top'               Vertical
  #                         'vcenter'
  #                         'bottom'
  #                         'vjustify'
  #
  # This method is used to set the horizontal and vertical text alignment
  # within a cell. Vertical and horizontal alignments can be combined.
  #  The method is used as follows:
  #
  #     format = workbook.add_format
  #     format->set_align('center')
  #     format->set_align('vcenter')
  #     worksheet->set_row(0, 30)
  #     worksheet->write(0, 0, 'X', format)
  #
  # Text can be aligned across two or more adjacent cells using
  # the center_across property. However, for genuine merged cells
  # it is better to use the merge_range() worksheet method.
  #
  # The vjustify (vertical justify) option can be used to provide
  # automatic text wrapping in a cell. The height of the cell will be
  # adjusted to accommodate the wrapped text. To specify where the text
  # wraps use the set_text_wrap() method.
  #
  # For further examples see the 'Alignment' worksheet created by formats.rb.
  #
  def set_align(align = 'left')

    return unless align.kind_of?(String)

    location = align.downcase

    case location
    when 'left'             then set_text_h_align(1)
    when 'centre', 'center' then set_text_h_align(2)
    when 'right'            then set_text_h_align(3)
    when 'fill'             then set_text_h_align(4)
    when 'justify'          then set_text_h_align(5)
    when 'center_across', 'centre_across' then set_text_h_align(6)
    when 'merge'            then set_text_h_align(6) # S:WE name
    when 'distributed'      then set_text_h_align(7)
    when 'equal_space'      then set_text_h_align(7) # ParseExcel

    when 'top'              then set_text_v_align(0)
    when 'vcentre'          then set_text_v_align(1)
    when 'vcenter'          then set_text_v_align(1)
    when 'bottom'           then set_text_v_align(2)
    when 'vjustify'         then set_text_v_align(3)
    when 'vdistributed'     then set_text_v_align(4)
    when 'vequal_space'     then set_text_v_align(4) # ParseExcel
    end
  end

  ###############################################################################
  #
  # set_valign()
  #
  # Set vertical cell alignment. This is required by the set_format_properties()
  # method to differentiate between the vertical and horizontal properties.
  #
  def set_valign(alignment)
    set_align(alignment);
  end

  #
  # Implements the Excel5 style "merge".
  #
  #     Default state:      Center across selection is off
  #     Default action:     Turn center across on
  #     Valid args:         1
  #
  # Text can be aligned across two or more adjacent cells using the
  # set_center_across() method. This is an alias for the
  # set_align('center_across') method call.
  #
  # Only one cell should contain the text, the other cells should be blank:
  #
  #     format = workbook.add_format
  #     format.set_center_across
  #
  #     worksheet.write(1, 1, 'Center across selection', format)
  #     worksheet.write_blank(1, 2, format)
  #
  # See also the merge1.pl to merge6.rb programs in the examples directory and
  # the merge_range() method.
  #
  def set_center_across(arg = 1)
    set_text_h_align(6)
  end

  ###############################################################################
  #
  # set_merge()
  #
  # This was the way to implement a merge in Excel5. However it should have been
  # called "center_across" and not "merge".
  # This is now deprecated. Use set_center_across() or better merge_range().
  #
  #
  def set_merge(val=true)
    set_text_h_align(6)
  end

  #
  #    Default state:      Text wrap is off
  #    Default action:     Turn text wrap on
  #    Valid args:         0, 1
  #
  # Here is an example using the text wrap property, the escape
  # character \n is used to indicate the end of line:
  #
  #    format = workbook.add_format()
  #    format.set_text_wrap()
  #    worksheet.write(0, 0, "It's\na bum\nwrap", format)
  #
  def set_text_wrap(arg = 1)
    begin
      if    arg == 0 then @text_wrap = 0
      elsif arg == 1 then @text_wrap = 1
      else
        raise ArgumentError,
        "\n\n  set_text_wrap(#{arg.inspect})\n    arg must be 0, 1, or none.\n"
        " ( 0:OFF, 1 and none:text wrap On )\n"
      end
    end
  end

  #
  # Set the bold property of the font:
  #
  #     Default state:      bold is off
  #     Default action:     Turn bold on
  #     Valid args:         0, 1 [1]
  #
  #     format.set_bold()   # Turn bold on
  #
  # [1] Actually, values in the range 100..1000 are also valid. 400 is normal,
  # 700 is bold and 1000 is very bold indeed. It is probably best to set the
  # value to 1 and use normal bold.
  #
  def set_bold(weight = nil)
    if weight.nil?
      weight = 0x2BC
    elsif !weight.kind_of?(Numeric)
      weight = 0x190
    elsif weight == 1                    # Bold text
      weight = 0x2BC
    elsif weight == 0                    # Normal text
      weight = 0x190
    elsif weight <  0x064                # Lower bound
      weight = 0x190
    elsif weight >  0x3E8                # Upper bound
      weight = 0x190
    else
      weight = weight.to_i
    end

    @bold = weight
  end


  #
  # Set cells borders to the same style
  #
  #     Also applies to:    set_bottom()
  #                         set_top()
  #                         set_left()
  #                         set_right()
  #
  #     Default state:      Border is off
  #     Default action:     Set border type 1
  #     Valid args:         0-13, See below.
  #
  # A cell border is comprised of a border on the bottom, top, left and right.
  # These can be set to the same value using set_border() or individually
  # using the relevant method calls shown above.
  #
  # The following shows the border styles sorted by WriteExcel index number:
  #
  #     Index   Name            Weight   Style
  #     =====   =============   ======   ===========
  #     0       None            0
  #     1       Continuous      1        -----------
  #     2       Continuous      2        -----------
  #     3       Dash            1        - - - - - -
  #     4       Dot             1        . . . . . .
  #     5       Continuous      3        -----------
  #     6       Double          3        ===========
  #     7       Continuous      0        -----------
  #     8       Dash            2        - - - - - -
  #     9       Dash Dot        1        - . - . - .
  #     10      Dash Dot        2        - . - . - .
  #     11      Dash Dot Dot    1        - . . - . .
  #     12      Dash Dot Dot    2        - . . - . .
  #     13      SlantDash Dot   2        / - . / - .
  #
  # The following shows the borders sorted by style:
  #
  #     Name            Weight   Style         Index
  #     =============   ======   ===========   =====
  #     Continuous      0        -----------   7
  #     Continuous      1        -----------   1
  #     Continuous      2        -----------   2
  #     Continuous      3        -----------   5
  #     Dash            1        - - - - - -   3
  #     Dash            2        - - - - - -   8
  #     Dash Dot        1        - . - . - .   9
  #     Dash Dot        2        - . - . - .   10
  #     Dash Dot Dot    1        - . . - . .   11
  #     Dash Dot Dot    2        - . . - . .   12
  #     Dot             1        . . . . . .   4
  #     Double          3        ===========   6
  #     None            0                      0
  #     SlantDash Dot   2        / - . / - .   13
  #
  # The following shows the borders in the order shown in the Excel Dialog.
  #
  #     Index   Style             Index   Style
  #     =====   =====             =====   =====
  #     0       None              12      - . . - . .
  #     7       -----------       13      / - . / - .
  #     4       . . . . . .       10      - . - . - .
  #     11      - . . - . .       8       - - - - - -
  #     9       - . - . - .       2       -----------
  #     3       - - - - - -       5       -----------
  #     1       -----------       6       ===========
  #
  # Examples of the available border styles are shown in the 'Borders' worksheet
  # created by formats.rb.
  #
  def set_border(style)
    set_bottom(style)
    set_top(style)
    set_left(style)
    set_right(style)
  end

  #
  # set bottom border of the cell.
  # see set_border() about style.
  #
  def set_bottom(style)
    @bottom = style
  end

  #
  # set top border of the cell.
  # see set_border() about style.
  #
  def set_top(style)
    @top = style
  end

  #
  # set left border of the cell.
  # see set_border() about style.
  #
  def set_left(style)
    @left = style
  end

  #
  # set right border of the cell.
  # see set_border() about style.
  #
  def set_right(style)
    @right = style
  end

  #
  # Set cells border to the same color
  #
  #     Also applies to:    set_bottom_color()
  #                         set_top_color()
  #                         set_left_color()
  #                         set_right_color()
  #
  #     Default state:      Color is off
  #     Default action:     Undefined
  #     Valid args:         See set_color()
  #
  # Set the colour of the cell borders. A cell border is comprised of a border
  # on the bottom, top, left and right. These can be set to the same colour
  # using set_border_color() or individually using the relevant method calls
  # shown above. Examples of the border styles and colours are shown in the
  # 'Borders' worksheet created by formats.rb.
  #
  def set_border_color(color)
    set_bottom_color(color);
    set_top_color(color);
    set_left_color(color);
    set_right_color(color);
  end

  #
  # set bottom border color of the cell.
  # see set_border_color() about color.
  #
  def set_bottom_color(color)
    @bottom_color = get_color(color)
  end

  #
  # set top border color of the cell.
  # see set_border_color() about color.
  #
  def set_top_color(color)
    @top_color = get_color(color)
  end

  #
  # set left border color of the cell.
  # see set_border_color() about color.
  #
  def set_left_color(color)
    @left_color = get_color(color)
  end

  #
  # set right border color of the cell.
  # see set_border_color() about color.
  #
  def set_right_color(color)
    @right_color = get_color(color)
  end

  #
  # Set the rotation angle of the text. An alignment property.
  #
  #     Default state:      Text rotation is off
  #     Default action:     None
  #     Valid args:         Integers in the range -90 to 90 and 270
  #
  # Set the rotation of the text in a cell. The rotation can be any angle in
  # the range -90 to 90 degrees.
  #
  #     format = workbook.add_format
  #     format.set_rotation(30)
  #     worksheet.write(0, 0, 'This text is rotated', format)
  #
  # The angle 270 is also supported. This indicates text where the letters run
  # from top to bottom.
  #
  def set_rotation(rotation)
    # Argument should be a number
    return unless rotation.kind_of?(Numeric)

    # The arg type can be a double but the Excel dialog only allows integers.
    rotation = rotation.to_i

    #      if (rotation == 270)
    #         rotation = 255
    #      elsif (rotation >= -90 or rotation <= 90)
    #         rotation = -rotation +90 if rotation < 0;
    #      else
    #         # carp "Rotation $rotation outside range: -90 <= angle <= 90";
    #         rotation = 0;
    #      end
    #
    if rotation == 270
      rotation = 255
    elsif rotation >= -90 && rotation <= 90
      rotation = -rotation + 90 if rotation < 0
    else
      rotation = 0
    end

    @rotation = rotation;
  end


  #
  # :call-seq:
  #    set_format_properties( :bold => 1 [, :color => 'red'..] )
  #    set_format_properties( font [, shade, ..])
  #    set_format_properties( :bold => 1, font, ...)
  #      *) font  = { :color => 'red', :bold => 1 }
  #         shade = { :bg_color => 'green', :pattern => 1 }
  #
  # Convert hashes of properties to method calls.
  #
  # The properties of an existing Format object can be also be set by means
  # of set_format_properties():
  #
  #     format = workbook.add_format
  #     format.set_format_properties(:bold => 1, :color => 'red');
  #
  # However, this method is here mainly for legacy reasons. It is preferable
  # to set the properties in the format constructor:
  #
  #     format = workbook.add_format(:bold => 1, :color => 'red');
  #
  def set_format_properties(*properties)
    return if properties.empty?
    properties.each do |property|
      property.each do |key, value|
        # Strip leading "-" from Tk style properties e.g. -color => 'red'.
        key.sub!(/^-/, '') if key.kind_of?(String)

        # Create a sub to set the property.
        if value.kind_of?(String)
          s = "set_#{key}('#{value}')"
        else
          s = "set_#{key}(#{value})"
        end
        eval s
      end
    end
  end

  #
  # :call-seq:
  #    set_font(fontname)
  #
  #     Default state:      Font is Arial
  #     Default action:     None
  #     Valid args:         Any valid font name
  #
  # Specify the font used:
  #
  #     format.set_font('Times New Roman');
  #
  # Excel can only display fonts that are installed on the system that it is
  # running on. Therefore it is best to use the fonts that come as standard
  # such as 'Arial', 'Times New Roman' and 'Courier New'. See also the Fonts
  # worksheet created by formats.rb
  #
  def set_font(fontname)
    @font = fontname
  end

  #
  # This method is used to define the numerical format of a number in Excel.
  #
  #     Default state:      General format
  #     Default action:     Format index 1
  #     Valid args:         See the following table
  #
  # It controls whether a number is displayed as an integer, a floating point
  # number, a date, a currency value or some other user defined format.
  #
  # The numerical format of a cell can be specified by using a format string
  # or an index to one of Excel's built-in formats:
  #
  #     format1 = workbook.add_format
  #     format2 = workbook.add_format
  #     format1.set_num_format('d mmm yyyy')  # Format string
  #     format2.set_num_format(0x0f)          # Format index
  #
  #     worksheet.write(0, 0, 36892.521, format1)       # 1 Jan 2001
  #     worksheet.write(0, 0, 36892.521, format2)       # 1-Jan-01
  #
  # Using format strings you can define very sophisticated formatting of
  # numbers.
  #
  #     format01.set_num_format('0.000')
  #     worksheet.write(0,  0, 3.1415926, format01)     # 3.142
  #
  #     format02.set_num_format('#,##0')
  #     worksheet.write(1,  0, 1234.56,   format02)     # 1,235
  #
  #     format03.set_num_format('#,##0.00')
  #     worksheet.write(2,  0, 1234.56,   format03)     # 1,234.56
  #
  #     format04.set_num_format('0.00')
  #     worksheet.write(3,  0, 49.99,     format04)     # 49.99
  #
  #     # Note you can use other currency symbols such as the pound or yen as well.
  #     # Other currencies may require the use of Unicode.
  #
  #     format07.set_num_format('mm/dd/yy')
  #     worksheet.write(6,  0, 36892.521, format07)     # 01/01/01
  #
  #     format08.set_num_format('mmm d yyyy')
  #     worksheet.write(7,  0, 36892.521, format08)     # Jan 1 2001
  #
  #     format09.set_num_format('d mmmm yyyy')
  #     worksheet.write(8,  0, 36892.521, format09)     # 1 January 2001
  #
  #     format10.set_num_format('dd/mm/yyyy hh:mm AM/PM')
  #     worksheet.write(9,  0, 36892.521, format10)     # 01/01/2001 12:30 AM
  #
  #     format11.set_num_format('0 "dollar and" .00 "cents"')
  #     worksheet.write(10, 0, 1.87,      format11)     # 1 dollar and .87 cents
  #
  #     # Conditional formatting
  #     format12.set_num_format('[Green]General;[Red]-General;General')
  #     worksheet.write(11, 0, 123,       format12)     # > 0 Green
  #     worksheet.write(12, 0, -45,       format12)     # < 0 Red
  #     worksheet.write(13, 0, 0,         format12)     # = 0 Default colour
  #
  #     # Zip code
  #     format13.set_num_format('00000')
  #     worksheet.write(14, 0, '01209',   format13)
  #
  # The number system used for dates is described in "DATES AND TIME IN EXCEL".
  #
  # The colour format should have one of the following values:
  #
  #     [Black] [Blue] [Cyan] [Green] [Magenta] [Red] [White] [Yellow]
  #
  # Alternatively you can specify the colour based on a colour index as follows:
  # [Color n], where n is a standard Excel colour index - 7. See the
  # 'Standard colors' worksheet created by formats.rb.
  #
  # For more information refer to the documentation on formatting in the doc
  # directory of the WriteExcel distro, the Excel on-line help or
  # http://office.microsoft.com/en-gb/assistance/HP051995001033.aspx
  #
  # You should ensure that the format string is valid in Excel prior to using
  # it in WriteExcel.
  #
  # Excel's built-in formats are shown in the following table:
  #
  #     Index   Index   Format String
  #     0       0x00    General
  #     1       0x01    0
  #     2       0x02    0.00
  #     3       0x03    #,##0
  #     4       0x04    #,##0.00
  #     5       0x05    ($#,##0_);($#,##0)
  #     6       0x06    ($#,##0_);[Red]($#,##0)
  #     7       0x07    ($#,##0.00_);($#,##0.00)
  #     8       0x08    ($#,##0.00_);[Red]($#,##0.00)
  #     9       0x09    0%
  #     10      0x0a    0.00%
  #     11      0x0b    0.00E+00
  #     12      0x0c    # ?/?
  #     13      0x0d    # ??/??
  #     14      0x0e    m/d/yy
  #     15      0x0f    d-mmm-yy
  #     16      0x10    d-mmm
  #     17      0x11    mmm-yy
  #     18      0x12    h:mm AM/PM
  #     19      0x13    h:mm:ss AM/PM
  #     20      0x14    h:mm
  #     21      0x15    h:mm:ss
  #     22      0x16    m/d/yy h:mm
  #     ..      ....    ...........
  #     37      0x25    (#,##0_);(#,##0)
  #     38      0x26    (#,##0_);[Red](#,##0)
  #     39      0x27    (#,##0.00_);(#,##0.00)
  #     40      0x28    (#,##0.00_);[Red](#,##0.00)
  #     41      0x29    _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
  #     42      0x2a    _($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)
  #     43      0x2b    _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
  #     44      0x2c    _($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)
  #     45      0x2d    mm:ss
  #     46      0x2e    [h]:mm:ss
  #     47      0x2f    mm:ss.0
  #     48      0x30    ##0.0E+0
  #     49      0x31    @
  #
  # For examples of these formatting codes see the 'Numerical formats' worksheet
  # created by formats.rb.
  #--
  # See also the number_formats1.html and the number_formats2.html documents in
  # the doc directory of the distro.
  #++
  #
  # Note 1. Numeric formats 23 to 36 are not documented by Microsoft and may
  # differ in international versions.
  #
  # Note 2. In Excel 5 the dollar sign appears as a dollar sign. In Excel
  # 97-2000 it appears as the defined local currency symbol.
  #
  # Note 3. The red negative numeric formats display slightly differently in
  # Excel 5 and Excel 97-2000.
  #
  def set_num_format(num_format)
    @num_format = num_format
  end

  #
  # This method can be used to indent text. The argument, which should be an
  # integer, is taken as the level of indentation:
  #
  #     Default state:      Text indentation is off
  #     Default action:     Indent text 1 level
  #     Valid args:         Positive integers
  #
  #     format = workbook.add_format
  #     format.set_indent(2)
  #     worksheet.write(0, 0, 'This text is indented', format)
  #
  # Indentation is a horizontal alignment property. It will override any
  # other horizontal properties but it can be used in conjunction with
  # vertical properties.
  #
  def set_indent(indent = 1)
    @indent = indent
  end

  #
  # This method can be used to shrink text so that it fits in a cell.
  #
  #     Default state:      Text shrinking is off
  #     Default action:     Turn "shrink to fit" on
  #     Valid args:         1
  #
  #     format = workbook.add_format
  #     format.set_shrink
  #     worksheet.write(0, 0, 'Honey, I shrunk the text!', format)
  #
  def set_shrink(arg = 1)
    @shrink = 1
  end

  #
  #     Default state:      Justify last is off
  #     Default action:     Turn justify last on
  #     Valid args:         0, 1
  #
  # Only applies to Far Eastern versions of Excel.
  #
  def set_text_justlast(arg = 1)
    @text_justlast = 1
  end

  #
  #     Default state:      Pattern is off
  #     Default action:     Solid fill is on
  #     Valid args:         0 .. 18
  #
  # Set the background pattern of a cell.
  #
  # Examples of the available patterns are shown in the 'Patterns' worksheet
  # created by formats.rb. However, it is unlikely that you will ever need
  # anything other than Pattern 1 which is a solid fill of the background color.
  #
  def set_pattern(pattern = 1)
    @pattern = pattern
  end

  #
  # The set_bg_color() method can be used to set the background colour of a
  # pattern. Patterns are defined via the set_pattern() method. If a pattern
  # hasn't been defined then a solid fill pattern is used as the default.
  #
  #     Default state:      Color is off
  #     Default action:     Solid fill.
  #     Valid args:         See set_color()
  #
  # Here is an example of how to set up a solid fill in a cell:
  #
  #     format = workbook.add_format
  #
  #     format.set_pattern()  # This is optional when using a solid fill
  #
  #     format.set_bg_color('green')
  #     worksheet.write('A1', 'Ray', format)
  #
  # For further examples see the 'Patterns' worksheet created by formats.rb.
  #
  def set_bg_color(color = 0x41)
    @bg_color = get_color(color)
  end

  #
  # The set_fg_color() method can be used to set the foreground colour
  # of a pattern.
  #
  #     Default state:      Color is off
  #     Default action:     Solid fill.
  #     Valid args:         See set_color()
  #
  # For further examples see the 'Patterns' worksheet created by formats.rb.
  #
  def set_fg_color(color = 0x40)
    @fg_color = get_color(color)
  end

  # Dynamically create set methods that aren't already defined.
  def method_missing(name, *args)
    # -- original perl comment --
    # There are two types of set methods: set_property() and
    # set_property_color(). When a method is AUTOLOADED we store a new anonymous
    # sub in the appropriate slot in the symbol table. The speeds up subsequent
    # calls to the same method.

    method = "#{name}"

    # Check for a valid method names, i.e. "set_xxx_yyy".
    method =~ /set_(\w+)/ or raise "Unknown method: #{method}\n"

    # Match the attribute, i.e. "@xxx_yyy".
    attribute = "@#{$1}"

    # Check that the attribute exists
    # ........
    if method =~ /set\w+color$/    # for "set_property_color" methods
      value = get_color(args[0])
    else                            # for "set_xxx" methods
      value = args[0].nil? ? 1 : args[0]
    end
    if value.kind_of?(String)
      s = "#{attribute} = \"#{value.to_s}\""
    else
      s = "#{attribute} =   #{value.to_s}"
    end
    eval s
  end

end
