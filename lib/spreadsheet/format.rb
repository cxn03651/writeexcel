   ##############################################################################
#
# Format - A class for defining Excel formatting.
#
#
# Used in conjunction with Spreadsheet::WriteExcel
#
# Copyright 2000-2008, John McNamara, jmcnamara@cpan.org
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

   attr_accessor :xf_index, :used_merge
   attr_accessor :bold, :text_wrap, :text_justlast
   attr_accessor :text_h_align, :text_v_align
   attr_accessor :fg_color, :bg_color, :color, :font, :size, :font_outline, :font_shadow
   attr_accessor :align, :border
   attr_reader   :font, :size, :font_strikeout, :font_script, :num_format, :locked, :hidden
   attr_reader   :rotation, :indent, :shrink, :pattern, :bottom, :top, :left, :right
   attr_reader   :bottom_color, :top_color, :left_color, :right_color
   attr_reader   :italic, :underline, :font_strikeout

   ###############################################################################
   #
   # initialize(xf_index=0, properties = nil)
   #    xf_index   : 
   #    properties : Hash of property => value
   #
   # Constructor
   #
   def initialize(xf_index=0, properties = [])
      @type           = 0
      @font_index     = 0
      @font           = 'Arial'
      @size           = 10
      @bold           = 0x0190
      @italic         = 0
      @color          = 0x7FFF
      @underline      = 0
      @font_strikeout      = 0
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

      ## convenience methods
      @border = 0
      @align = 'left'

      set_format_properties(properties)

      @xf_index = xf_index
   end

###############################################################################
#
# Note to porters. The majority of the set_property() methods are created
# dynamically via Perl' AUTOLOAD sub, see below. You may prefer/have to specify
# them explicitly in other implementation languages.
#


###############################################################################
#
# get_font()
#
# Generate an Excel BIFF FONT record.
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
      @diag_border = 1 if (@diag_border !=0 and @diag_type != 0)

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

      # Handle utf8 strings in perl 5.8.
#      if ($] >= 5.008) {
#         require Encode;
#
#         if (Encode::is_utf8($rgch)) {
#            $rgch = Encode::encode("UTF-16BE", $rgch);
#            $encoding = 1;
#         }
#      }
#
      cch = rgch.length;
#
      # Handle Unicode font names.
      if (encoding == 1)
#         croak "Uneven number of bytes in Unicode font name" if cch % 2;
         cch  /= 2 if encoding;
         rgch  = rgch.unpack('n*').pack('v*')
      end

      record     = 0x31;
      length     = 0x10 + rgch.length;
      reserved   = 0x00;

      grbit      = 0x00;
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

   ###############################################################################
   #
   # set_size(size)
   #
   #    Default state:      Font size is 10
   #    Default action:     Set font size to 1
   #    Valid args:         Integer values from 1 to as big as your screen.
   #                        Set the font size. Excel adjusts the height of a row
   #                        to accommodate the
   #
   # largest font size in the row. You can also explicitly specify the height
   # of a row using the set_row() worksheet method.Set cell alignment.
   #
   def set_size(size = 1)
      if size.kind_of?(Numeric) && size >= 1
         @size = size.to_i
      end
   end

   ###############################################################################
   #
   # set_color(color)
   #
   #    Default state:      Excels default color, usually black
   #    Default action:     Set the default color
   #    Valid args:         Integers from 8..63 or the following strings:
   #                        'black', 'blue', 'brown', 'cyan', 'gray'
   #                        'green', 'lime', 'magenta', 'navy', 'orange'
   #                        'pink', 'purple', 'red', 'silver', 'white', 'yellow'
   #
   # Set the font colour. The set_color() method is used as follows:
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
   
   ###############################################################################
   #
   # set_italic()
   #
   #    Default state:      Italic is off
   #    Default action:     Turn italic on
   #    Valid args:         0, 1
   #
   #  Set the italic property of the font:
   #
   #    format.set_italic()  # Turn italic on
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

   ###############################################################################
   #
   # set_underline()
   #
   #    Default state:      Underline is off
   #    Default action:     Turn on single underline
   #    Valid args:         0  = No underline
   #                        1  = Single underline
   #                        2  = Double underline
   #                        33 = Single accounting underline
   #                        34 = Double accounting underline
   #
   #  Set the underline property of the font.
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

   ###############################################################################
   #
   # set_font_strikeout()
   # 
   #     Default state:      Strikeout is off
   #     Default action:     Turn strikeout on
   #     Valid args:         0, 1
   # 
   # Set the strikeout property of the font.
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

   ###############################################################################
   #
   # set_font_script()
   # 
   #     Default state:      Super/Subscript is off
   #     Default action:     Turn Superscript on
   #     Valid args:         0  = Normal
   #                         1  = Superscript
   #                         2  = Subscript
   # 
   # Set the superscript/subscript property of the font. 
   # This format is currently not very useful.
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

   ###############################################################################
   #
   # set_font_outline()
   # 
   #     Default state:      Outline is off
   #     Default action:     Turn outline on
   #     Valid args:         0, 1
   # 
   # Macintosh only.
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
   
   ###############################################################################
   #
   # set_font_shadow()
   # 
   #     Default state:      Shadow is off
   #     Default action:     Turn shadow on
   #     Valid args:         0, 1
   # 
   # Macintosh only.
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
   
   ###############################################################################
   #
   # set_locked()
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

   ###############################################################################
   #
   # set_align()
   #
   # Set cell alignment.
   #
   def set_align(location = nil)
   
      return if location.nil?           # No default
      return if location =~ /\d/;       # Ignore numbers

      location.downcase!

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

   ###############################################################################
   #
   # set_center_across()
   #
   # Implements the Excel5 style "merge".
   #
   def set_center_across
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
   def set_merge
      set_text_h_align(6);
   end

   ###############################################################################
   #
   # set_bold()
   #
   #    Default state:      bold is off  (internal value = 400)
   #    Default action:     Turn bold on
   #    Valid args:         0, 1 [1]
   #
   # Set the bold property of the font:
   #
   #    format.set_bold()   # Turn bold on
   #
   #[1] Actually, values in the range 100..1000 are also valid.
   #    400 is normal, 700 is bold and 1000 is very bold indeed.
   #    It is probably best to set the value to 1 and use normal bold.
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


   ###############################################################################
   #
   # set_border($style)
   #
   # Set cells borders to the same style
   #
   def set_border(style)
      set_bottom(style)
      set_top(style)
      set_left(style)
      set_right(style)
   end


   ###############################################################################
   #
   # set_border_color($color)
   #
   # Set cells border to the same color
   #
   def set_border_color(color)
      set_bottom_color(color);
      set_top_color(color);
      set_left_color(color);
      set_right_color(color);
   end

   ###############################################################################
   #
   # set_rotation($angle)
   #
   # Set the rotation angle of the text. An alignment property.
   #
   def set_rotation(rotation)
      # Argument should be a number
      return if !(rotation =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/)

      # The arg type can be a double but the Excel dialog only allows integers.
      rotation = rotation.to_i

      if (rotation == 270)
         rotation = 255
      elsif (rotation >= -90 or rotation <= 90)
         rotation = -rotation +90 if rotation < 0;
      else
         # carp "Rotation $rotation outside range: -90 <= angle <= 90";
         rotation = 0;
      end

      @rotation = rotation;
   end


   ###############################################################################
   #
   # set_format_properties(*properties)
   #    properties : Hash of properies
   #  ex)   font  = { :color => 'red', :bold => 1 }
   #        shade = { :bg_color => 'green', :pattern => 1 }
   #     1) set_format_properties( :bold => 1 [, :color => 'red'..] )
   #     2) set_format_properties( font [, shade, ..])
   #     3) set_format_properties( :bold => 1, font, ...)
   #
   # Convert hashes of properties to method calls.
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

# Renamed rarely used set_properties() to set_format_properties() to avoid
# confusion with Workbook method of the same name. The following acts as an
# alias for any code that uses the old name.
# *set_properties = *set_format_properties;

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
