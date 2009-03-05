require 'format'

class MaxSizeError < StandardError; end

class Worksheet < BIFFWriter

   RowMax = 65536
   ColMax = 256
   StrMax = 0
   Buffer = 4096

   attr_reader :name, :encoding, :xf_index, :index, :type, :images_array
   attr_reader :filter_area, :filter_count
   attr_reader :title_rowmin, :title_rowmax, :title_colmin, :title_colmax
   attr_reader :print_rowmin, :print_rowmax, :print_colmin, :print_colmax
   attr_accessor :index, :colinfo, :selection, :offset, :selected, :hidden, :active
   attr_accessor :object_ids

   ###############################################################################
   #
   # new()
   #
   # Constructor. Creates a new Worksheet object from a BIFFwriter object
   #
   def initialize(*args)
      super

      @name                = args[0]
      @index               = args[1]
      @encoding            = args[2]
      @active_sheet        = args[3]
      @first_sheet         = args[4]
      @url_format          = args[5]
      @parser              = args[6]
      @tempdir             = args[7]
      @str_total           = args[8]  || 0
      @str_unique          = args[9]  || 0
      @str_table           = args[10] || {}
      @v1904               = args[11]
      @compatibility       = args[12]

      @table               = []
      @row_data            = {}

      @type                = 0x0000
      @ext_sheets          = []
      @using_tmpfile       = 0    # _initialize not coverted yet.
      @filehandle          = ""
      @fileclosed          = 0
      @offset              = 0
      @xls_rowmax          = RowMax
      @xls_colmax          = ColMax
      @xls_strmax          = StrMax
      @dim_rowmin          = nil
      @dim_rowmax          = nil
      @dim_colmin          = nil
      @dim_colmax          = nil
      @colinfo             = []
      @selection           = [0, 0]
      @panes               = []
      @active_pane         = 3
      @frozen              = 0
      @frozen_no_split     = 1
      @selected            = 0
      @hidden              = 0
      @active              = 0
      @tab_color           = 0

      @first_row           = 0
      @first_col           = 0
      @display_formulas    = 0
      @display_headers     = 1
      @display_zeros       = 1
      @display_arabic      = 0

      @paper_size          = 0x0
      @orientation         = 0x1
      @header              = ''
      @footer              = ''
      @header_encoding     = 0
      @footer_encoding     = 0
      @hcenter             = 0
      @vcenter             = 0
      @margin_header       = 0.50
      @margin_footer       = 0.50
      @margin_left         = 0.75
      @margin_right        = 0.75
      @margin_top          = 1.00
      @margin_bottom       = 1.00

      @title_rowmin        = nil
      @title_rowmax        = nil
      @title_colmin        = nil
      @title_colmax        = nil
      @print_rowmin        = nil
      @print_rowmax        = nil
      @print_colmin        = nil
      @print_colmax        = nil

      @print_gridlines     = 1
      @screen_gridlines    = 1
      @print_headers       = 0

      @page_order          = 0
      @black_white         = 0
      @draft_quality       = 0
      @print_comments      = 0
      @page_start          = 1
      @custom_start        = 0

      @fit_page            = 0
      @fit_width           = 0
      @fit_height          = 0

      @hbreaks             = []
      @vbreaks             = []

      @protect             = 0
      @password            = nil

      @col_sizes           = {}
      @row_sizes           = {}

      @col_formats         = {}
      @row_formats         = {}

      @zoom                = 100
      @print_scale         = 100
      @page_view           = 0

      @leading_zeros       = 0

      @outline_row_level   = 0
      @outline_style       = 0
      @outline_below       = 1
      @outline_right       = 1
      @outline_on          = 1

      @write_match         = []

      @object_ids          = []
      @images              = {}
      @charts              = {}
      @comments            = {}
      @comments_author     = ''
      @comments_author_enc = 0
      @comments_visible    = 0

      @filter_area         = []
      @filter_count        = 0
      @filter_on           = 0
      @filter_cols         = []

      @writing_url         = 0

      @db_indices          = []

      @validations         = []
      _initialize
   end

   def _initialize
      basename = 'spreadsheetwriteexcel'

      begin
         if !@tempdir.nil?
            if @tempdir == ''
               fh = Tempfile.new(basename)
            else
               fh = Tempfile.new(basename, @tempdir)
            end
         end
      # failed. store temporary data in memory.
      rescue
         @using_tmpfile = 0
      # if the temp file creation was successful
      else
         @filehandle = fh
      end
   end
   
   ###############################################################################
   #
   # _close()
   #
   # Add data to the beginning of the workbook (note the reverse order)
   # and to the end of the workbook.
   #
   def close(sheetnames)
       num_sheets = sheetnames.size
   
       ################################################
       # Prepend in reverse order!!
       #
   
       # Prepend the sheet dimensions
       store_dimensions
   
       # Prepend the autofilter filters.
       store_autofilters
   
       # Prepend the sheet autofilter info.
       store_autofilterinfo
   
       # Prepend the sheet filtermode record.
       store_filtermode
   
       # Prepend the COLINFO records if they exist
       if @colinfo
           while (@colinfo)
               arrayref = @colinfo.pop
               store_colinfo(arrayref)
           end
       end
   
       # Prepend the DEFCOLWIDTH record
       store_defcol
   
       # Prepend the sheet password
       store_password
   
       # Prepend the sheet protection
       store_protect
       store_obj_protect
   
       # Prepend the page setup
       store_setup
   
       # Prepend the bottom margin
       store_margin_bottom
   
       # Prepend the top margin
       store_margin_top
   
       # Prepend the right margin
       store_margin_right
   
       # Prepend the left margin
       store_margin_left
   
       # Prepend the page vertical centering
       store_vcenter
   
       # Prepend the page horizontal centering
       store_hcenter
   
       # Prepend the page footer
       store_footer
   
       # Prepend the page header
       store_header
   
       # Prepend the vertical page breaks
       store_vbreak
   
       # Prepend the horizontal page breaks
       store_hbreak
   
       # Prepend WSBOOL
       store_wsbool
   
       # Prepend the default row height.
       store_defrow
   
       # Prepend GUTS
       store_guts
   
       # Prepend GRIDSET
       store_gridset
   
       # Prepend PRINTGRIDLINES
       store_print_gridlines
   
       # Prepend PRINTHEADERS
       store_print_headers
   
       #
       # End of prepend. Read upwards from here.
       ################################################
   
       # Append
       store_table
       store_images
       store_charts
       store_filters
       store_comments
       store_window2
       store_page_view
       store_zoom
       store_panes(@panes) if !@panes.nil? && @panes != 0
       store_selection(@selection)
       store_validation_count
       store_validations
       store_tab_color
       store_eof
   
       # Prepend the BOF and INDEX records
       store_index
       store_bof(0x0010)
   end

   ###############################################################################
   #
   # _compatibility_mode()
   #
   # Set the compatibility mode.
   #
   # See the explanation in Workbook::compatibility_mode(). This private method
   # is mainly used for test purposes.
   #
   def compatibility_mode(compatibility = 1)
      @compatibility = compatibility
   end

   ###############################################################################
   #
   # get_data().
   #
   # Retrieves data from memory in one chunk, or from disk in $buffer
   # sized chunks.
   #
   def get_data
      buffer = 4096

      # Return data stored in memory
      unless @data.nil?
         tmp   = @data
         @data = nil
         fh         = @filehandle
         seek(fh, 0, 0) if @using_tmpfile != 0
         return @data
      end

      # Return data stored on disk
      if @using_tmpfile != 0
         return tmp if read(@filehandle, tmp, buffer)
      end

      # No data to return
      return nil
   end


   ###############################################################################
   #
   # select()
   #
   # Set this worksheet as a selected worksheet, i.e. the worksheet has its tab
   # highlighted.
   #
   def select
      @hidden         = 0  # Selected worksheet can't be hidden.
      @selected       = 1
   end


   ###############################################################################
   #
   # activate()
   #
   # Set this worksheet as the active worksheet, i.e. the worksheet that is
   # displayed when the workbook is opened. Also set it as selected.
   #
   def activate
      @hidden      = 0  # Active worksheet can't be hidden.
      @selected    = 1
      @activesheet = @index
   end


   ###############################################################################
   #
   # hide()
   #
   # Hide this worksheet.
   #
   def hide
      @hidden         = 1

      # A hidden worksheet shouldn't be active or selected.
      @selected       = 0
      @activesheet = 0
      @firstsheet  = 0
   end


   ###############################################################################
   #
   # set_first_sheet()
   #
   # Set this worksheet as the first visible sheet. This is necessary
   # when there are a large number of worksheets and the activated
   # worksheet is not visible on the screen.
   #
   def set_first_sheet
      @hidden      = 0  # Active worksheet can't be hidden.
      @firstsheet  = @index
   end


   ###############################################################################
   #
   # protect($password)
   #
   # Set the worksheet protection flag to prevent accidental modification and to
   # hide formulas if the locked and hidden format properties have been set.
   #
   def protect(password = nil)
      @protect   = 1
      @password  = encode_password(password) unless password.nil?
   end

   ###############################################################################
   #
   # set_column($firstcol, $lastcol, $width, $format, $hidden, $level)
   #
   # Set the width of a single column or a range of columns.
   # See also: _store_colinfo
   #
   def set_column(*args)
      data = args
      cell = data[0]

      # Check for a cell reference in A1 notation and substitute row and column
      if cell =~ /^\D/
         data = substitute_cellref(*args)

         # Returned values $row1 and $row2 aren't required here. Remove them.
         data.shift        # $row1
         data.delete_at(1) # $row2
      end

      return if data.size < 3  # Ensure at least $firstcol, $lastcol and $width
      return if data[0].nil?   # Columns must be defined.
      return if data[1].nil?

      # Assume second column is the same as first if 0. Avoids KB918419 bug.
      data[1] = data[0] if data[1] == 0

      # Ensure 2nd col is larger than first. Also for KB918419 bug.
      data[0], data[1] = data[1], data[0] if data[0] > data[1]

      # Limit columns to Excel max of 255.
      data[0] = ColMax - 1 if data[0] > ColMax - 1
      data[1] = ColMax - 1 if data[1] > ColMax - 1

      @colinfo.push(data)

      # Store the col sizes for use when calculating image vertices taking
      # hidden columns into account. Also store the column formats.
      #
      width  = data[4].nil? || data[4] == 0 ? 0 : data[2]  # Set width to zero if col is hidden
      width  ||= 0                    # Ensure width isn't undef.
      format = data[3]
      firstcol, lastcol = data

      (firstcol .. lastcol).each do |col|
         @col_sizes[col]   = width
         @col_formats[col] = format unless format.nil?
      end
   end

   ###############################################################################
   #
   # set_selection()
   #
   # Set which cell or cells are selected in a worksheet: see also the
   # sub _store_selection
   #
   def set_selection(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end
      @selection = args
   end


   ###############################################################################
   #
   # freeze_panes()
   #
   # Set panes and mark them as frozen. See also _store_panes().
   #
   def freeze_panes(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end
      # Extra flag indicated a split and freeze.
      @frozen_no_split = 0 if args[4]

      @frozen = 1
      @panes  = args
   end


   ###############################################################################
   #
   # split_panes()
   #
   # Set panes and mark them as split. See also _store_panes().
   #
   def split_panes(*args)
      @frozen            = 0
      @frozen_no_split   = 0
      @panes             = args
   end

   # Older method name for backwards compatibility.
   # *thaw_panes = *split_panes;

   ###############################################################################
   #
   # set_portrait()
   #
   # Set the page orientation as portrait.
   #
   def set_portrait
      @orientation = 1
   end


   ###############################################################################
   #
   # set_landscape()
   #
   # Set the page orientation as landscape.
   #
   def set_landscape
      @orientation = 0
   end


   ###############################################################################
   #
   # set_page_view()
   #
   # Set the page view mode for Mac Excel.
   #
   def set_page_view(val = nil)
      @page_view = val.nil? ? 1 : val
   end


   ###############################################################################
   #
   # set_tab_color()
   #
   # Set the colour of the worksheet colour.
   #
   def set_tab_color(colour)
      color = Format._get_color(colour)
      color = 0 if color == 0x7FFF # Default color.
      @tab_color = color
   end

   ###############################################################################
   #
   # set_paper()
   #
   # Set the paper type. Ex. 1 = US Letter, 9 = A4
   #
   def set_paper(paper_size = 0)
      @paper_size = paper_size
   end

   ###############################################################################
   #
   # set_header()
   #
   # Set the page header caption and optional margin.
   #
   def set_header(string = '', margin = 0.50, encoding = 0)
      limit    = encoding != 0 ? 255 *2 : 255

      if string.length >= limit
         #           carp 'Header string must be less than 255 characters';
         return
      end

      @header          = string
      @margin_header   = margin
      @header_encoding = encoding
   end


   ###############################################################################
   #
   # set_footer()
   #
   # Set the page footer caption and optional margin.
   #
   def set_footer(string = '', margin = 0.50, encoding = 0)
      limit    = encoding != 0 ? 255 *2 : 255

      if string.length >= limit
         #           carp 'Header string must be less than 255 characters';
         return
      end

      @footer          = string
      @margin_footer   = margin
      @footer_encoding = encoding
   end

   ###############################################################################
   #
   # center_horizontally()
   #
   # Center the page horizontally.
   #
   def center_horizontally(hcenter = nil)
      if hcenter.nil?
         @hcenter = 1
      else
         @hcenter = hcenter
      end
   end

   ###############################################################################
   #
   # center_vertically()
   #
   # Center the page horinzontally.
   #
   def center_vertically(vcenter = nil)
      if vcenter.nil?
         @vcenter = 1
      else
         @vcenter = vcenter
      end
   end

   ###############################################################################
   #
   # set_margins()
   #
   # Set all the page margins to the same value in inches.
   #
   def set_margins(margin)
      set_margin_left(margin)
      set_margin_right(margin)
      set_margin_top(margin)
      set_margin_bottom(margin)
   end

   ###############################################################################
   #
   # set_margins_LR()
   #
   # Set the left and right margins to the same value in inches.
   #
   def set_margins_LR(margin)
      set_margin_left(margin)
      set_margin_right(margin)
   end

   ###############################################################################
   #
   # set_margins_TB()
   #
   # Set the top and bottom margins to the same value in inches.
   #
   def set_margins_TB(margin)
      set_margin_top(margin)
      set_margin_bottom(margin)
   end


   ###############################################################################
   #
   # set_margin_left()
   #
   # Set the left margin in inches.
   #
   def set_margin_left(margin = 0.75)
      @margin_left = margin
   end


   ###############################################################################
   #
   # set_margin_right()
   #
   # Set the right margin in inches.
   #
   def set_margin_right(margin = 0.75)
      @margin_right = margin
   end

   ###############################################################################
   #
   # set_margin_top()
   #
   # Set the top margin in inches.
   #
   def set_margin_top(margin = 1.00)
      @margin_top = margin
   end

   ###############################################################################
   #
   # set_margin_bottom()
   #
   # Set the bottom margin in inches.
   #
   def set_margin_bottom(margin = 1.00)
      @margin_bottom = margin
   end

   ###############################################################################
   #
   # repeat_rows($first_row, $last_row)
   #
   # Set the rows to repeat at the top of each printed page. See also the
   # _store_name_xxxx() methods in Workbook.pm.
   #
   def repeat_rows(first_row, last_row = nil)
      @title_rowmin  = first_row
      @title_rowmax  = last_row || first_row # Second row is optional
   end

   ###############################################################################
   #
   # repeat_columns($first_col, $last_col)
   #
   # Set the columns to repeat at the left hand side of each printed page.
   # See also the _store_names() methods in Workbook.pm.
   #
   def repeat_columns(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args =~ /^\D/
         args = substitute_cellref(*args)

         # Returned values $row1 and $row2 aren't required here. Remove them.
         args.shift        # $row1
         args.delete_at(1) # $row2
      end
   
      @title_colmin  = args[0]
      @title_colmax  = args[1] || args[0] # Second col is optional
   end

   ###############################################################################
   #
   # print_area($first_row, $first_col, $last_row, $last_col)
   #
   # Set the area of each worksheet that will be printed. See also the
   # _store_names() methods in Workbook.pm.
   #
   def print_area(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args =~ /^\D/
         args = substitute_cellref(*args)
      end

      return if args.size != 4 # Require 4 parameters

      @print_rowmin, @print_colmin, @print_rowmax, @print_colmax = args
   end

   ###############################################################################
   #
   # autofilter($first_row, $first_col, $last_row, $last_col)
   #
   # Set the autofilter area in the worksheet.
   #
   def autofilter(*args)
     # Check for a cell reference in A1 notation and substitute row and column
     if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end

     return if args.size != 4 # Require 4 parameters

     row1, col1, row2, col2 = args

      # Reverse max and min values if necessary.
      if row2 < row1
         tmp  = row1
         row1 = row2
         row2 = tmp
      end
      if col2 < col1
         tmp  = col1
         col1 = col2
         col2 = col1
      end

      # Store the Autofilter information
      @filter_area = [row1, row2, col1, col2]
      @filter_count = 1 + col2 -col1
   end

   ###############################################################################
   #
   # filter_column($column, $criteria, ...)
   #
   # Set the column filter criteria.
   #
   def filter_column(col, expression)
      raise "Must call autofilter() before filter_column()" if @filter_count == 0
#      raise "Incorrect number of arguments to filter_column()" unless @_ == 2

      # Check for a column reference in A1 notation and substitute.
      if col =~ /^\D/
         # Convert col ref to a cell ref and then to a col number.
         no_use, col = substitute_cellref(col + '1')
      end
      col_first = @filter_area[2]
      col_last  = @filter_area[3]

      # Reject column if it is outside filter range.
      if (col < col_first or col > col_last)
         raise "Column '#{col}' outside autofilter() column range " +
               "(#{col_first} .. #{col_last})";
      end

      tokens = extract_filter_tokens(expression)

      unless (tokens.size == 3 or tokens.size == 7)
         raise "Incorrect number of tokens in expression '#{expression}'"
      end


      tokens = parse_filter_expression(expression, tokens)

      @filter_cols[col] = Array.new(tokens)
      @filter_on        = 1
   end

   ###############################################################################
   #
   # _extract_filter_tokens($expression)
   #
   # Extract the tokens from the filter expression. The tokens are mainly non-
   # whitespace groups. The only tricky part is to extract string tokens that
   # contain whitespace and/or quoted double quotes (Excel's escaped quotes).
   #
   # Examples: 'x <  2000'
   #           'x >  2000 and x <  5000'
   #           'x = "foo"'
   #           'x = "foo bar"'
   #           'x = "foo "" bar"'
   #
   def extract_filter_tokens(expression = nil)
      return unless expression

      #  @tokens = ($expression  =~ /"(?:[^"]|"")*"|\S+/g); #"

      tokens = []
      str = expression
      while str =~ /"(?:[^"]|"")*"|\S+/
         tokens << $&
         str = $~.post_match
      end

      # Remove leading and trailing quotes and unescape other quotes
      tokens.map! do |token|
         token.sub!(/^"/, '')
         token.sub!(/"$/, '')
         token.gsub!(/""/, '"')
      end

      return tokens
   end

   ###############################################################################
   #
   # _parse_filter_expression(expression, @token)
   #
   # Converts the tokens of a possibly conditional expression into 1 or 2
   # sub expressions for further parsing.
   #
   # Examples:
   #          ('x', '==', 2000) -> exp1
   #          ('x', '>',  2000, 'and', 'x', '<', 5000) -> exp1 and exp2
   #
   def parse_filter_expression(expression, tokens)
      # The number of tokens will be either 3 (for 1 expression)
      # or 7 (for 2  expressions).
      #
      if (tokens.size == 7)
         conditional = tokens[3]
         if conditional =~ /^(and|&&)$/
            conditional = 0
         elsif conditional =~ /^(or|\|\|)$/
            conditional = 1
         else
            raise "Token '#{conditional}' is not a valid conditional " +
                  "in filter expression '#{expression}'"
         end
         expression_1 = parse_filter_tokens(expression, tokens[0..2])
         expression_2 = parse_filter_tokens(expression, tokens[4..6])
         return [expression_1, conditional, expression_2]
      else
         return parse_filter_tokens(expression, tokens)
      end
   end

   ###############################################################################
   #
   # _parse_filter_tokens(@token)  # (@expression, @token)
   #
   # Parse the 3 tokens of a filter expression and return the operator and token.
   #
   def parse_filter_tokens(expression, tokens)
      operators = {
         '==' => 2,
         '='  => 2,
         '=~' => 2,
         'eq' => 2,

         '!=' => 5,
         '!~' => 5,
         'ne' => 5,
         '<>' => 5,

         '<'  => 1,
         '<=' => 3,
         '>'  => 4,
         '>=' => 6,
      }

      operator = operators[tokens[1]]
      token    = tokens[2]

      # Special handling of "Top" filter expressions.
      if tokens[0] =~ /^top|bottom$/i
         value = tokens[1]
         if (value =~ /\D/ or value < 1 or value > 500)
            raise "The value '#{value}' in expression '#{expression}' " +
                  "must be in the range 1 to 500"
         end
         token.downcase!
         if (token != 'items' and token != '%')
            raise "The type '#{token}' in expression '#{expression}' " +
                  "must be either 'items' or '%'"
         end

         if (tokens[0] =~ /^top$/i)
            operator = 30
         else
            operator = 32
         end

         if (tokens[2] == '%')
            operator = operator + 1
         end

         token    = value
      end

      if (not operator and tokens[0])
         raise "Token '#{tokens[1]}' is not a valid operator " +
               "in filter expression '#{expression}'"
      end

      # Special handling for Blanks/NonBlanks.
      if (token =~ /^blanks|nonblanks$/i)
         # Only allow Equals or NotEqual in this context.
         if (operator != 2 and operator != 5)
            raise "The operator '#{tokens[1]}' in expression '#{expression}' " +
                  "is not valid in relation to Blanks/NonBlanks'"
         end

         token.downcase!
         
         # The operator should always be 2 (=) to flag a "simple" equality in
         # the binary record. Therefore we convert <> to =.
         if (token == 'blanks')
            if (operator == 5)
               operator = 2
               token    = 'nonblanks'
            end
         else
            if (operator == 5)
               operator = 2
               token    = 'blanks'
            end
         end
      end

      # if the string token contains an Excel match character then change the
      # operator type to indicate a non "simple" equality.
      if (operator == 2 and token =~ /[*?]/)
         operator = 22
      end

      return [operator, token]
   end

   ###############################################################################
   #
   # hide_gridlines()
   #
   # Set the option to hide gridlines on the screen and the printed page.
   # There are two ways of doing this in the Excel BIFF format: The first is by
   # setting the DspGrid field of the WINDOW2 record, this turns off the screen
   # and subsequently the print gridline. The second method is to via the
   # PRINTGRIDLINES and GRIDSET records, this turns off the printed gridlines
   # only. The first method is probably sufficient for most cases. The second
   # method is supported for backwards compatibility. Porters take note.
   #
   def hide_gridlines(option = 1)
      if option == 0
         @print_gridlines  = 1  # 1 = display, 0 = hide
         @screen_gridlines = 1
      elsif option == 1
         @print_gridlines  = 0
         @screen_gridlines = 1
      else
         @print_gridlines  = 0
         @screen_gridlines = 0
      end
   end

   ###############################################################################
   #
   # print_row_col_headers()
   #
   # Set the option to print the row and column headers on the printed page.
   # See also the _store_print_headers() method below.
   #
   def print_row_col_headers(option = nil)
      if option.nil?
         @print_headers = 1
      else
         @print_headers = option
      end
   end

   ###############################################################################
   #
   # fit_to_pages($width, $height)
   #
   # Store the vertical and horizontal number of pages that will define the
   # maximum area printed. See also _store_setup() and _store_wsbool() below.
   #
   def fit_to_pages(width = 0, height = 0)
      @fit_page      = 1
      @fit_width     = width
      @fit_height    = height
   end

   ###############################################################################
   #
   # set_h_pagebreaks(@breaks)
   #
   # Store the horizontal page breaks on a worksheet.
   #
   def set_h_pagebreaks(breaks)
      @hbreaks.push(breaks)
   end
   
   ###############################################################################
   #
   # set_v_pagebreaks(@breaks)
   #
   # Store the vertical page breaks on a worksheet.
   #
   def set_v_pagebreaks(breaks)
      @vbreaks.push(breaks)
   end
   
   ###############################################################################
   #
   # set_zoom($scale)
   #
   # Set the worksheet zoom factor.
   #
   def set_zoom(scale = 100)
      # Confine the scale to Excel's range
      if scale < 10 or scale > 400
         #           carp "Zoom factor $scale outside range: 10 <= zoom <= 400";
         scale = 100
      end

      @zoom = scale.to_i
   end
   
   ###############################################################################
   #
   # set_print_scale($scale)
   #
   # Set the scale factor for the printed page.
   #
   def set_print_scale(scale = 100)
      # Confine the scale to Excel's range
      if scale < 10 or scale > 400
         #           carp "Print scale $scale outside range: 10 <= zoom <= 400";
         scale = 100
      end

      # Turn off "fit to page" option
      @fit_page    = 0
   
      @print_scale = scale.to_i
   end
   
   ###############################################################################
   #
   # keep_leading_zeros()
   #
   # Causes the write() method to treat integers with a leading zero as a string.
   # This ensures that any leading zeros such, as in zip codes, are maintained.
   #
   def keep_leading_zeros(val = nil)
      if val.nil?
         @leading_zeros = 1
      else
         @leading_zeros = val
      end
   end
   
   ###############################################################################
   #
   # show_comments()
   #
   # Make any comments in the worksheet visible.
   #
   def show_comments(val = nil)
      @comments_visible = val.nil? ? 1 : val
   end
   
   ###############################################################################
   #
   # set_comments_author()
   #
   # Set the default author of the cell comments.
   #
   def set_comments_author(author = '', author_enc = 0)
      @comments_author     = author
      @comments_author_enc = author_enc
   end
   
   ###############################################################################
   #
   # right_to_left()
   #
   # Display the worksheet right to left for some eastern versions of Excel.
   #
   def right_to_left(val = nil)
      @display_arabic = val.nil? ? 1 : val
   end
   
   ###############################################################################
   #
   # hide_zero()
   #
   # Hide cell zero values.
   #
   def hide_zero(val = nil)
      @display_zeros = val.nil? ? 0 : !val
   end
   
   ###############################################################################
   #
   # print_across()
   #
   # Set the order in which pages are printed.
   #
   def print_across(val = nil)
      @page_order = val.nil? ? 1 : val
   end
   
   ###############################################################################
   #
   # set_start_page()
   #
   # Set the start page number.
   #
   def set_start_page(start_page = nil)
      return if start_page.nil?
   
      @page_start    = start_page
      @custom_start  = 1
   end

   ###############################################################################
   #
   # set_first_row_column()
   #
   # Set the topmost and leftmost visible row and column.
   # TODO: Document this when tested fully for interaction with panes.
   #
   def set_first_row_column(row = 0, col = 0)
      row = RowMax - 1  if row > RowMax - 1
      col = ColMax - 1  if col > ColMax - 1
   
      @first_row = row
      @first_col = col
   end

   ###############################################################################
   #
   # add_write_handler($re, $code_ref)
   #
   # Allow the user to add their own matches and handlers to the write() method.
   #
   def add_write_handler(regexp, code_ref)
      #       return unless ref $_[1] eq 'CODE';

      @write_match.push([regexp, code_ref])
   end

   ###############################################################################
   #
   # write($row, $col, $token, $format)
   #
   # Parse $token and call appropriate write method. $row and $column are zero
   # indexed. $format is optional.
   #
   # The write_url() methods have a flag to prevent recursion when writing a
   # string that looks like a url.
   #
   # Returns: return value of called subroutine
   #
   def write(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end

      token = args[2]

      # Handle undefs as blanks
      token = '' if token.nil?

      # First try user defined matches.
      @write_match.each do |aref|
         re  = aref[0]
         sub = aref[1]

         if token =~ Regexp.new(re)
            match = eval("#{sub} self, args")
            return match unless match.nil?
         end
      end

      # Match an array ref.
      if token.kind_of?(Array)
         return write_row(*args)
         # Match integer with leading zero(s)
      elsif @leading_zeros != 0 and token =~ /^0\d+$/
         return write_string(*args)
         # Match number
      elsif token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/
         return write_number(*args)
         # Match http, https or ftp URL
      elsif token =~ %r|^[fh]tt?ps?://|    and @writing_url == 0
         return write_url(*args)
         # Match mailto:
      elsif token =~ %r|^mailto:|          and @writing_url == 0
         return write_url(*args)
         # Match internal or external sheet link
      elsif token =~ %r!^(?:in|ex)ternal:! and @writing_url == 0
         return write_url(*args)
         # Match formula
      elsif token =~ /^=/
         return write_formula(*args)
         # Match blank
      elsif token == ''
         args.delete_at(2)     # remove the empty string from the parameter list
         return write_blank(*args)
      else
         return write_string(*args)
      end
   end


   ###############################################################################
   #
   # write_row($row, $col, $array_ref, $format)
   #
   # Write a row of data starting from ($row, $col). Call write_col() if any of
   # the elements of the array ref are in turn array refs. This allows the writing
   # of 1D or 2D arrays of data in one go.
   #
   # Returns: the first encountered error value or zero for no errors
   #
   def write_row(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end

      # Catch non array refs passed by user.
      unless args[2].kind_of?(Array)
         raise "Not an array ref in call to write_row() #{$!}";
      end

      row, col, tokens, *options = args
      error   = 0
      unless tokens.nil?
         tokens.each do |token|
            # Check for nested arrays
            if token.kind_of?(Array)
               ret = write_col(row, col, token, options)
            else
               ret = write(row, col, token, options)
            end

            # Return only the first error encountered, if any.
            error ||= ret
            col = col + 1
         end
      end
      return error
   end


   ###############################################################################
   #
   # write_col($row, $col, $array_ref, $format)
   #
   # Write a column of data starting from ($row, $col). Call write_row() if any of
   # the elements of the array ref are in turn array refs. This allows the writing
   # of 1D or 2D arrays of data in one go.
   #
   # Returns: the first encountered error value or zero for no errors
   #
   def write_col(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end

      # Catch non array refs passed by user.
      unless args[2].kind_of?(Array)
         raise "Not an array ref in call to write_row()";
      end

      row, col, tokens, *options = args
      error   = 0
      unless tokens.nil?
         tokens.each do |token|
            # write() will deal with any nested arrays
            ret = write(row, col, token, options)

            # Return only the first error encountered, if any.
            error ||= ret
            col = col + 1
         end
      end
      return error
   end


   ###############################################################################
   #
   # write_comment($row, $col, $comment)
   #
   # Write a comment to the specified row and column (zero indexed).
   #
   # Returns  0 : normal termination
   #         -1 : insufficient number of arguments
   #         -2 : row or column out of range
   #
   def write_comment(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end

      return -1 if args.size < 3   # Check the number of args


      row = args[0]
      col = args[1]

      # Check for pairs of optional arguments, i.e. an odd number of args.
      raise "Uneven number of additional arguments" if args.size % 2 == 0

      # Check that row and col are valid and store max and min values
      return -2 if check_dimensions(row, col) != 0

      # We have to avoid duplicate comments in cells or else Excel will complain.
      @comments[row] = { col => comment_params(*args) }
   end

   ###############################################################################
   #
   # _xf_record_index()
   #
   # Returns an index to the XF record in the workbook.
   #
   # Note: this is a function, not a method.
   #
   def xf_record_index(row, col, xf=nil)
      if xf.kind_of?(Format)
         return xf.xf_index
      elsif @row_formats.has_key?(row) 
         return @row_formats[row].xf_index
      elsif @col_formats.has_key?(col)
         return @col_formats[col].xf_index
      else
         return 0x0F
      end
   end
   
   ###############################################################################
   #
   # _substitute_cellref()
   #
   # Substitute an Excel cell reference in A1 notation for  zero based row and
   # column values in an argument list.
   #
   # Ex: ("A4", "Hello") is converted to (3, 0, "Hello").
   #
   def substitute_cellref(cell, *args)
      cell.upcase!

      # Convert a column range: 'A:A' or 'B:G'.
      # A range such as A:A is equivalent to A1:65536, so add rows as required
      if cell =~ /\$?([A-I]?[A-Z]):\$?([A-I]?[A-Z])/
         row1, col1 =  cell_to_rowcol($1 +'1')
         row2, col2 =  cell_to_rowcol($2 +'65536')
         return [row1, col1, row2, col2, *args]
      end

      # Convert a cell range: 'A1:B7'
      if cell =~ /\$?([A-I]?[A-Z]\$?\d+):\$?([A-I]?[A-Z]\$?\d+)/
         row1, col1 =  cell_to_rowcol($1)
         row2, col2 =  cell_to_rowcol($2)
         return [row1, col1, row2, col2, *args]
      end

      # Convert a cell reference: 'A1' or 'AD2000'
      if (cell =~ /\$?([A-I]?[A-Z]\$?\d+)/)
         row1, col1 =  cell_to_rowcol($1)
         return [row1, col1, *args]

      end

      raise("Unknown cell reference #{cell}")
   end

   ###############################################################################
   #
   # _cell_to_rowcol($cell_ref)
   #
   # Convert an Excel cell reference in A1 notation to a zero based row and column
   # reference; converts C1 to (0, 2).
   #
   # Returns: row, column
   #
   def cell_to_rowcol(cell)
      cell =~ /\$?([A-I]?[A-Z])\$?(\d+)/
      col     = $1
      row     = $2.to_i

      # Convert base26 column string to number
      # All your Base are belong to us.
      chars = col.split(//)
      expn  = 0
      col   = 0

      while (chars.size > 0)
         char = chars.pop   # LS char first
         col = col + (char[0] - 65 +1) * (26**expn)
             #####  ord(char) - ord('A')  in perl  ####
         expn = expn + 1
      end

      # Convert 1-index to zero-index
      row -= 1
      col -= 1

      return [row, col]
   end

   ###############################################################################
   #
   # write_number($row, $col, $num, $format)
   #
   # Write a double to the specified row and column (zero indexed).
   # An integer can be written as a double. Excel will display an
   # integer. $format is optional.
   #
   # Returns  0 : normal termination
   #         -1 : insufficient number of arguments
   #         -2 : row or column out of range
   #
   def write_number(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end

      return -1 if (args.size < 3)                # Check the number of args

       record  = 0x0203                        # Record identifier
       length  = 0x000E                        # Number of bytes to follow
   
       row     = args[0]                         # Zero indexed row
       col     = args[1]                         # Zero indexed column
       num     = args[2]
       xf      = xf_record_index(row, col, args[3]) # The cell format
   
       # Check that row and col are valid and store max and min values
       return -2 if check_dimensions(row, col) != 0
   
       header = [record, length].pack('vv')
       data   = [row, col, xf].pack('vvv')
       xl_double = [num].pack("d")
   
       xl_double.reverse! if @byte_order != 0
   
       # Store the data or write immediately depending on the compatibility mode.
       if @compatibility != 0
          tmp = []
          tmp[col] = header + data + xl_double
          @table[row] = tmp
       else
           append(header, data, xl_double)
       end
   
       return 0
   end

   ###############################################################################
   #
   # write_string ($row, $col, $string, $format)
   #
   # Write a string to the specified row and column (zero indexed).
   # NOTE: there is an Excel 5 defined limit of 255 characters.
   # $format is optional.
   # Returns  0 : normal termination
   #         -1 : insufficient number of arguments
   #         -2 : row or column out of range
   #         -3 : long string truncated to 255 chars
   #
   def write_string(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end

      return -1 if (args.size < 3)                # Check the number of args

      record      = 0x00FD                        # Record identifier
      length      = 0x000A                        # Bytes to follow

      row         = args[0]                       # Zero indexed row
      col         = args[1]                       # Zero indexed column
      str         = args[2].to_s
      strlen      = str.length
      xf          = xf_record_index(row, col, args[3])   # The cell format
      encoding    = 0x0
      str_error   = 0

      # Check that row and col are valid and store max and min values
      return -2 if check_dimensions(row, col) != 0

      # Limit the string to the max number of chars.
      if (strlen > 32767)
         str       = substr(str, 0, 32767)
         str_error = -3
      end

      # Prepend the string with the type.
      str_header  = [str.length, encoding].pack('vC')
      str         = str_header + str

      if @str_table[str].nil?
         @str_table[str] = @str_unique
         @str_unique += 1
      end

      @str_total += 1

      header = [record, length].pack('vv')
      data   = [row, col, xf, @str_table[str]].pack('vvvV')

      # Store the data or write immediately depending on the compatibility mode.
      if @compatibility != 0
         tmp = []
         tmp[col] = header + data
         @table[row] = tmp
      else
         append(header, data)
      end

      return str_error
   end

   ###############################################################################
   #
   # write_blank($row, $col, $format)
   #
   # Write a blank cell to the specified row and column (zero indexed).
   # A blank cell is used to specify formatting without adding a string
   # or a number.
   #
   # A blank cell without a format serves no purpose. Therefore, we don't write
   # a BLANK record unless a format is specified. This is mainly an optimisation
   # for the write_row() and write_col() methods.
   #
   # Returns  0 : normal termination (including no format)
   #         -1 : insufficient number of arguments
   #         -2 : row or column out of range
   #
   def write_blank(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end

      # Check the number of args
      return -1 if args.size < 2

      # Don't write a blank cell unless it has a format
      return 0 if args[2].nil?

      record  = 0x0201                        # Record identifier
      length  = 0x0006                        # Number of bytes to follow

      row     = args[0]                       # Zero indexed row
      col     = args[1]                       # Zero indexed column
      xf      = xf_record_index(row, col, args[2])   # The cell format

      # Check that row and col are valid and store max and min values
      return -2 if check_dimensions(row, col) != 0

      header    = [record, length].pack('vv')
      data      = [row, col, xf].pack('vvv')
       
      # Store the data or write immediately depending    on the compatibility mode.
      if @compatibility != 0
         tmp = []
         tmp[col] = header + data
         @table[row] = tmp
      else
         append(header, data)
      end

      return 0
   end

   ###############################################################################
   #
   # write_url($row, $col, $url, $string, $format)
   #
   # Write a hyperlink. This is comprised of two elements: the visible label and
   # the invisible link. The visible label is the same as the link unless an
   # alternative string is specified.
   #
   # The parameters $string and $format are optional and their order is
   # interchangeable for backward compatibility reasons.
   #
   # The hyperlink can be to a http, ftp, mail, internal sheet, or external
   # directory url.
   #
   # Returns  0 : normal termination
   #         -1 : insufficient number of arguments
   #         -2 : row or column out of range
   #         -3 : long string truncated to 255 chars
   #
   def write_url(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end

      # Check the number of args
      return -1 if args.size < 3

      # Add start row and col to arg list
      return write_url_range(args[0], args[1], args)
   end

   ###############################################################################
   #
   # write_url_range($row1, $col1, $row2, $col2, $url, $string, $format)
   #
   # This is the more general form of write_url(). It allows a hyperlink to be
   # written to a range of cells. This function also decides the type of hyperlink
   # to be written. These are either, Web (http, ftp, mailto), Internal
   # (Sheet1!A1) or external ('c:\temp\foo.xls#Sheet1!A1').
   #
   # See also write_url() above for a general description and return values.
   #
   def write_url_range(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end

      # Check the number of args
      return -1 if args.size < 5

      # Reverse the order of $string and $format if necessary. We work on a copy
      # in order to protect the callers args. We don't use "local @_" in case of
      # perl50005 threads.
      #
      #       my @args = @_;
      #
      #       ($args[5], $args[6]) = ($args[6], $args[5]) if ref $args[5];
      #
      url = args[4]

      # Check for internal/external sheet links or default to web link
      return write_url_internal(*args) if url =~ /^internal:/
      return write_url_external(*args) if url =~ /^external:/
      return write_url_web(*args)
   end

   ###############################################################################
   #
   # _write_url_web($row1, $col1, $row2, $col2, $url, $string, $format)
   #    row1        = $_[0];                        # Start row
   #    col1        = $_[1];                        # Start column
   #    row2        = $_[2];                        # End row
   #    col2        = $_[3];                        # End column
   #    url         = $_[4];                        # URL string
   #    str         = $_[5];                        # Alternative label
   #
   # Used to write http, ftp and mailto hyperlinks.
   # The link type ($options) is 0x03 is the same as absolute dir ref without
   # sheet. However it is differentiated by the $unknown2 data stream.
   #
   # See also write_url() above for a general description and return values.
   #
   def _write_url_web(row1, col1, row2, col2, url, str = nil, format = nil)
      record = 0x01B8                       # Record identifier
      length = 0x00000                      # Bytes to follow

      xf     = format || @url_format        # The cell format

      # Write the visible label but protect against url recursion in write().
      str          = url if str.nil?
      @writing_url = 1
      error        = write(row1, col1, str, xf)
      @writing_url = 0
      return error if error == -2

      # Pack the undocumented parts of the hyperlink stream
      unknown1    = ["D0C9EA79F9BACE118C8200AA004BA90B02000000"].pack("H*")
      unknown2    = ["E0C9EA79F9BACE118C8200AA004BA90B"].pack("H*")

      # Pack the option flags
      options     = [0x03].pack("V")

      # Convert URL to a null terminated wchar string
      url         = url.split('').join("\0")
      url         = url + "\0\0\0"

      # Pack the length of the URL
      url_len     = [url.length].pack("V")

      # Calculate the data length
      length         = 0x34 + url.length

      # Pack the header data
      header      = [record, length].pack("vv")
      data        = [row1, row2, col1, col2].pack("vvvv")

      # Write the packed data
      append( header, data,unknown1,options,unknown2,url_len,url)

      return error
   end


   ###############################################################################
   #
   # _write_url_internal($row1, $col1, $row2, $col2, $url, $string, $format)
   #    row1        = $_[0];                        # Start row
   #    col1        = $_[1];                        # Start column
   #    row2        = $_[2];                        # End row
   #    col2        = $_[3];                        # End column
   #    url         = $_[4];                        # URL string
   #    str         = $_[5];                        # Alternative label
   #
   # Used to write internal reference hyperlinks such as "Sheet1!A1".
   #
   # See also write_url() above for a general description and return values.
   #
   def _write_url_internal(row1, col1, row2, col2, url, str = nil, format = nil)
      record = 0x01B8                       # Record identifier
      length = 0x00000                      # Bytes to follow

      xf     = format || @url_format        # The cell format

      # Strip URL type
      url.sub!(/^internal:/, '')

      # Write the visible label but protect against url recursion in write().
      str          = url if str.nil?
      @writing_url = 1
      error        = write(row1, col1, str, xf)
      @writing_url = 0
      return error if error == -2

      # Pack the undocumented parts of the hyperlink stream
      unknown1    = ["D0C9EA79F9BACE118C8200AA004BA90B02000000"].pack("H*")

      # Pack the option flags
      options     = [0x08].pack("V")

      # URL encoding.
      encoding    = 0

      # Convert an Ascii URL type and to a null terminated wchar string.
      if encoding == 0
         url = url + "\0"
         url = url.unpack('c*').pack('v*')
      end

      # Pack the length of the URL as chars (not wchars)
      url_len     = [(url.length/2).to_i].pack("V")

      # Calculate the data length
      length         = 0x24 + url.length

      # Pack the header data
      header      = [record, length].pack("vv")
      data        = [row1, row2, col1, col2].pack("vvvv")

      # Write the packed data
      append( header, data, unknown1, options, url_len, url)

      return error
   end

   ###############################################################################
   #
   # _write_url_external($row1, $col1, $row2, $col2, $url, $string, $format)
   #
   # Write links to external directory names such as 'c:\foo.xls',
   # c:\foo.xls#Sheet1!A1', '../../foo.xls'. and '../../foo.xls#Sheet1!A1'.
   #
   # Note: Excel writes some relative links with the $dir_long string. We ignore
   # these cases for the sake of simpler code.
   #
   # See also write_url() above for a general description and return values.
   #
   def _write_url_external(row1, col1, row2, col2, url, str = nil, format = nil)
      # Network drives are different. We will handle them separately
      # MS/Novell network drives and shares start with \\
      if url =~ /^external:\\\\/
         return write_url_external_net(row1, col1, row2, col2, url, str, format)
      end

      record      = 0x01B8                       # Record identifier
      length      = 0x00000                      # Bytes to follow

      xf     = format || @url_format        # The cell format

      # Strip URL type and change Unix dir separator to Dos style (if needed)
      #
      url.sub!(/^external:/, '')
      url.gsub!(%r|/|, '\\')


      # Write the visible label but protect against url recursion in write().
      str = url.sub!(/\#/, ' - ') if str.nil?
      @writing_url = 1
      error        = write(row1, col1, str, xf)
      @writing_url = 0
      return error if error == -2

      # Determine if the link is relative or absolute:
      # Absolute if link starts with DOS drive specifier like C:
      # Otherwise default to 0x00 for relative link.
      #
      absolute    = 0x00
      absolute    = 0x02  if url =~ /^[A-Za-z]:/

      # Determine if the link contains a sheet reference and change some of the
      # parameters accordingly.
      # Split the dir name and sheet name (if it exists)
      #
      dir_long , sheet = url.split(/\#/)
      link_type        = 0x01 | absolute

      unless sheet.nil?
         link_type |= 0x08
         sheet_len  = [sheet.length + 0x01].pack("V")
         sheet      = sheet.split('').join("\0") + "\0\0\0"
      else
         sheet_len   = ''
         sheet       = ''
      end

      # Pack the link type
      link_type      = link_type.pack("V")

      # Calculate the up-level dir count e.g. (..\..\..\ == 3)
      up_count    = 0
      while dir_long.sub!(/^\.\.\\/, '')
         up_count = up_count + 1
      end
      up_count    = [up_count].pack("v")

      # Store the short dos dir name (null terminated)
      dir_short   = dir_long + "\0"

      # Store the long dir name as a wchar string (non-null terminated)
      dir_long = dir_long.split('').join("\0") + "\0"

      # Pack the lengths of the dir strings
      dir_short_len = [dir_short.length].pack("V")
      dir_long_len  = [dir_long.length].pack("V")
      stream_len    = [dir_long.length + 0x06].pack("V")

      # Pack the undocumented parts of the hyperlink stream
      unknown1 = ['D0C9EA79F9BACE118C8200AA004BA90B02000000'].pack("H*")
      unknown2 = ['0303000000000000C000000000000046'].pack("H*")
      unknown3 = ['FFFFADDE000000000000000000000000000000000000000'].pack("H*")
      unknown4 = [0x03].pack("v")

      # Pack the main data stream
      data        = [row1, row2, col1, col2].pack("vvvv") +
        unknown1     +
        link_type    +
        unknown2     +
        up_count     +
        dir_short_len+
        dir_short    +
        unknown3     +
        stream_len   +
        dir_long_len +
        unknown4     +
        dir_long     +
        sheet_len    +
        sheet

      # Pack the header data
      length      = data.length
      header      = [record, length].pack("vv")

      # Write the packed data
      append(header, data)

      return error
   end

   ###############################################################################
   #
   # _write_url_external_net($row1, $col1, $row2, $col2, $url, $string, $format)
   #
   # Write links to external MS/Novell network drives and shares such as
   # '//NETWORK/share/foo.xls' and '//NETWORK/share/foo.xls#Sheet1!A1'.
   #
   # See also write_url() above for a general description and return values.
   #
   def _write_url_external_net(row1, col1, row2, col2, url, str, format)
      record      = 0x01B8                       # Record identifier
      length      = 0x00000                      # Bytes to follow

      xf          = format || @url_format  # The cell format

      # Strip URL type and change Unix dir separator to Dos style (if needed)
      #
      url.sub!(/^external:/, '')
      url.gsub!(%r|/|, '\\')

      # Write the visible label but protect against url recursion in write().
      str = url.sub!(/\#/, ' - ') if str.nil?
      @writing_url = 1
      error        = write(row1, col1, str, xf)
      @writing_url = 0
      return error if error == -2

      # Determine if the link contains a sheet reference and change some of the
      # parameters accordingly.
      # Split the dir name and sheet name (if it exists)
      #
      dir_long , sheet = url.split(/\#/)
      link_type        = 0x0103  # Always absolute

      unless sheet.nil?
         link_type |= 0x08
         sheet_len  = [sheet.length + 0x01].pack("V")
         sheet      = sheet.split('').join("\0") + "\0\0\0"
      else
         sheet_len   = ''
         sheet       = ''
      end

      # Pack the link type
      link_type      = [link_type].pack("V")


      # Make the string null terminated
      dir_long       = dir_long + "\0"

      # Pack the lengths of the dir string
      dir_long_len  = [dir_long.length].pack("V")

      # Store the long dir name as a wchar string (non-null terminated)
      dir_long = dir_long.split('').join("\0") + "\0"

      # Pack the undocumented part of the hyperlink stream
      unknown1    = ['D0C9EA79F9BACE118C8200AA004BA90B02000000'].pack("H*")

      # Pack the main data stream
        data         = [row1, row2, col1, col2].pack("vvvv") +
        unknown1     +
        link_type    +
        dir_long_len +
        dir_long     +
        sheet_len    +
        sheet

      # Pack the header data
      length      = data.length
      header      = [record, length].pack("vv")

      # Write the packed data
      append(header, data)

      return error
   end

   ###############################################################################
   #
   # write_date_time ($row, $col, $string, $format)
   #
   # Write a datetime string in ISO8601 "yyyy-mm-ddThh:mm:ss.ss" format as a
   # number representing an Excel date. $format is optional.
   #
   # Returns  0 : normal termination
   #         -1 : insufficient number of arguments
   #         -2 : row or column out of range
   #         -3 : Invalid date_time, written as string
   #
   def write_date_time(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end
   
      return -1 if (args.size < 3)                 # Check the number of args
   
      row       = args[0]                           # Zero indexed row
      col       = args[1]                           # Zero indexed column
      str       = args[2]
   
      # Check that row and col are valid and store max and min values
      return -2 if check_dimensions(row, col) != 0
   
      error     = 0
      date_time = convert_date_time(str)
   
      unless date_time.nil?
         error = write_number(row, col, date_time, args[3])
      else
         # The date isn't valid so write it as a string.
         write_string(row, col, str, args[3])
         error = -3
      end
      return error
   end
   
   ###############################################################################
   #
   # convert_date_time($date_time_string)
   #
   # The function takes a date and time in ISO8601 "yyyy-mm-ddThh:mm:ss.ss" format
   # and converts it to a decimal number representing a valid Excel date.
   #
   # Dates and times in Excel are represented by real numbers. The integer part of
   # the number stores the number of days since the epoch and the fractional part
   # stores the percentage of the day in seconds. The epoch can be either 1900 or
   # 1904.
   #
   # Parameter: Date and time string in one of the following formats:
   #               yyyy-mm-ddThh:mm:ss.ss  # Standard
   #               yyyy-mm-ddT             # Date only
   #                         Thh:mm:ss.ss  # Time only
   #
   # Returns:
   #            A decimal number representing a valid Excel date, or
   #            undef if the date is invalid.
   #
   def convert_date_time(date_time_string)
      date_time = date_time_string
   
      days      = 0 # Number of days since epoch
      seconds   = 0 # Time expressed as fraction of 24h hours in seconds
   
      # Strip leading and trailing whitespace.
      date_time.sub!(/^\s+/, '')
      date_time.sub!(/\s+$/, '')
   
      # Check for invalid date char.
      return nil if date_time =~ /[^0-9T:\-\.Z]/
   
      # Check for "T" after date or before time.
      return nil unless date_time =~ /\dT|T\d/
   
      # Strip trailing Z in ISO8601 date.
      date_time.sub!(/Z$/, '')
   
      # Split into date and time.
      date, time = date_time.split(/T/)
   
      # We allow the time portion of the input DateTime to be optional.
      if time != ''
         # Match hh:mm:ss.sss+ where the seconds are optional
         if time =~ /^(\d\d):(\d\d)(:(\d\d(\.\d+)?))?/
            hour   = $1
            min    = $2
            sec    = $4 || 0
         else
            return nil # Not a valid time format.
         end
   
         # Some boundary checks
         return nil if hour >= 24
         return nil if min  >= 60
         return nil if sec  >= 60
   
         # Excel expresses seconds as a fraction of the number in 24 hours.
         seconds = (hour *60*60 + min *60 + sec) / (24 *60 *60)
      end
   
      # We allow the date portion of the input DateTime to be optional.
      return seconds if date == ''
   
      # Match date as yyyy-mm-dd.
      if date =~ /^(\d\d\d\d)-(\d\d)-(\d\d)$/
         year   = $1
         month  = $2
         day    = $3
      else
         return nil  # Not a valid date format.
      end
   
      # Set the epoch as 1900 or 1904. Defaults to 1900.
      date_1904 = @v1904
   
      # Special cases for Excel.
      if !date_1904
         return      seconds if date == '1899-12-31' # Excel 1900 epoch
         return      seconds if date == '1900-01-00' # Excel 1900 epoch
         return 60 + seconds if date == '1900-02-29' # Excel false leapday
      end
   
   
      # We calculate the date by calculating the number of days since the epoch
      # and adjust for the number of leap days. We calculate the number of leap
      # days by normalising the year in relation to the epoch. Thus the year 2000
      # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
      #
      epoch   = date_1904 ? 1904 : 1900
      offset  = date_1904 ?    4 :    0
      norm    = 300
      range   = year -epoch
   
      # Set month days and check for leap year.
      mdays   = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
      leap    = 0
      leap    = 1  if year % 4 == 0 and year % 100 or year % 400 == 0
      mdays[1]   = 29 if leap
   
      # Some boundary checks
      return nil if year  < epoch or year  > 9999
      return nil if month < 1     or month > 12
      return nil if day   < 1     or day   > mdays[month -1]
   
      # Accumulate the number of days since the epoch.
      days = day                                     # Add days for current month
      (0 .. month-2).each do |m|
         days = days + mdays[m]                      # Add days for past months
      end
      days = days + range *365                       # Add days for past years
      days = days + ((range)                /  4)    # Add leapdays
      days = days - ((range + offset)       /100)    # Subtract 100 year leapdays
      days = days + ((range + offset + norm)/400)    # Add 400 year leapdays
      days = days - lea                              # Already counted above
   
      # Adjust for Excel erroneously treating 1900 as a leap year.
      days = days + 1 if date_1904 == 0 and days > 59
   
      return days + seconds
   end

   ###############################################################################
   #
   # set_row($row, $height, $format, $hidden, $level, collapsed)
   #          row       : Row Number
   #          height    : Format object
   #          format    : Format object
   #          hidden    : Hidden flag
   #          level     : Outline level
   #          collapsed : Collapsed row
   # This method is used to set the height and XF format for a row.
   # Writes the  BIFF record ROW.
   #
   def set_row(row, height = nil, format = nil, hidden = 0, level = 0, collapsed = 0)
      record      = 0x0208               # Record identifier
      length      = 0x0010               # Number of bytes to follow

      colMic      = 0x0000               # First defined column
      colMac      = 0x0000               # Last defined column
      # miyRw;                           # Row height
      irwMac      = 0x0000               # Used by Excel to optimise loading
      reserved    = 0x0000               # Reserved
      grbit       = 0x0000               # Option flags
      # ixfe;                            # XF index

      return if row.nil?

      # Check that row and col are valid and store max and min values
      return -2 if check_dimensions(row, 0, 0, 1) != 0
   
      # Check for a format object
      if format.kind_of?(Format)
         ixfe = format.get_xf_index
      else
         ixfe = 0x0F
      end
   
      # Set the row height in units of 1/20 of a point. Note, some heights may
      # not be obtained exactly due to rounding in Excel.
      #
      unless height.nil?
         miyRw = height *20
      else
         miyRw = 0xff # The default row height
         height = 0
      end

      # Set the limits for the outline levels (0 <= x <= 7).
      level = 0 if level < 0
      level = 7 if level > 7
   
      @outline_row_level = level if level > @outline_row_level
   
      # Set the options flags.
      # 0x10: The fCollapsed flag indicates that the row contains the "+"
      #       when an outline group is collapsed.
      # 0x20: The fDyZero height flag indicates a collapsed or hidden row.
      # 0x40: The fUnsynced flag is used to show that the font and row heights
      #       are not compatible. This is usually the case for WriteExcel.
      # 0x80: The fGhostDirty flag indicates that the row has been formatted.
      #
      grbit |= level
      grbit |= 0x0010 if collapsed != 0
      grbit |= 0x0020 if hidden    != 0
      grbit |= 0x0040
      grbit |= 0x0080 unless format.nil?
      grbit |= 0x0100
   
      header = [record, length].pack("vv")
      data   = [row, colMic, colMac, miyRw, irwMac, reserved, grbit, ixfe].pack("vvvvvvvv")

      # Store the data or write immediately depending on the compatibility mode.
      if @compatibility != 0
         @row_data[row] = header + data
      else
         append(header, data)
      end

      # Store the row sizes for use when calculating image vertices.
      # Also store the column formats.
      @row_sizes[row]   = height
      @row_formats[row] = format unless format.nil?
   end

   ###############################################################################
   #
   # _write_row_default()
   #        row    : Row Number
   #        colMic : First defined column
   #        colMac : Last defined column
   #
   # Write a default row record, in compatibility mode, for rows that don't have
   # user specified values..
   #
   def _write_row_default(row, colMic, colMac)

      record      = 0x0208               # Record identifier
      length      = 0x0010               # Number of bytes to follow

      miyRw       = 0xFF                 # Row height
      irwMac      = 0x0000               # Used by Excel to optimise loading
      reserved    = 0x0000               # Reserved
      grbit       = 0x0100               # Option flags
      ixfe        = 0x0F                 # XF index

      header = [record, length].pack("vv")
      data   = [row, colMic, colMac, miyRw, irwMac, reserved, grbit, ixfe].pack("vvvvvvvv")

      append(header, data)
   end

   ###############################################################################
   #
   # _check_dimensions($row, $col, $ignore_row, $ignore_col)
   #
   # Check that $row and $col are valid and store max and min values for use in
   # DIMENSIONS record. See, _store_dimensions().
   #
   # The $ignore_row/$ignore_col flags is used to indicate that we wish to
   # perform the dimension check without storing the value.
   #
   # The ignore flags are use by set_row() and data_validate.
   #
   def check_dimensions(row, col, ignore_row = 0, ignore_col = 0)
      return -2 if row.nil?
      return -2 if row >= @xls_rowmax

      return -2 if col.nil?
      return -2 if col >= @xls_colmax

      if ignore_row == 0
         if @dim_rowmin.nil? or row < @dim_rowmin
            @dim_rowmin = row
         end

         if @dim_rowmax.nil? or row > @dim_rowmax
            @dim_rowmax = row
         end
      end

      if ignore_col == 0
         if @dim_colmin.nil? or col < @dim_colmin
            @dim_colmin = col
         end

         if @dim_colmax.nil? or col > @dim_colmax
            @dim_colmax =col
         end
      end

      return 0
   end

   ###############################################################################
   #
   # _store_dimensions()
   #
   # Writes Excel DIMENSIONS to define the area in which there is cell data.
   #
   # Notes:
   #   Excel stores the max row/col as row/col +1.
   #   Max and min values of 0 are used to indicate that no cell data.
   #   We set the undef member data to 0 since it is used by _store_table().
   #   Inserting images or charts doesn't change the DIMENSION data.
   #
   def store_dimensions
      record    = 0x0200         # Record identifier
      length    = 0x000E         # Number of bytes to follow
      reserved  = 0x0000         # Reserved by Excel

      row_min = @dim_rowmin.nil? ? 0 : @dim_rowmin
      row_max = @dim_rowmax.nil? ? 0 : @dim_rowmax + 1
      col_min = @dim_colmin.nil? ? 0 : @dim_colmin
      col_max = @dim_colmax.nil? ? 0 : @dim_colmax + 1

      # Set member data to the new max/min value for use by _store_table().
      @dim_rowmin = row_min
      @dim_rowmax = row_max
      @dim_colmin = col_min
      @dim_colmax = col_max
  
      header = [record, length].pack("vv")
      fields = [row_min, row_max, col_min, col_max, reserved]
      data   = fields.pack("VVvvv")
     
      return prepend(header, data)
   end

   ###############################################################################
   #
   # _store_window2()
   #
   # Write BIFF record Window2.
   #
   def _store_window2
      record         = 0x023E     # Record identifier
      length         = 0x0012     # Number of bytes to follow

      grbit          = 0x00B6     # Option flags
      rwTop          = @first_row   # Top visible row
      colLeft        = @first_col   # Leftmost visible column
      rgbHdr         = 0x00000040            # Row/col heading, grid color

      wScaleSLV      = 0x0000                # Zoom in page break preview
      wScaleNormal   = 0x0000                # Zoom in normal view
      reserved       = 0x00000000


      # The options flags that comprise $grbit
      fDspFmla       = @display_formulas # 0 - bit
      fDspGrid       = @screen_gridlines # 1
      fDspRwCol      = @display_headers  # 2
      fFrozen        = @frozen           # 3
      fDspZeros      = @display_zeros    # 4
      fDefaultHdr    = 1                 # 5
      fArabic        = @display_arabic   # 6
      fDspGuts       = @outline_on       # 7
      fFrozenNoSplit = @frozen_no_split  # 0 - bit
      fSelected      = @selected         # 1
      fPaged         = @active           # 2
      fBreakPreview  = 0                # 3

      grbit             = fDspFmla
      grbit            |= fDspGrid       << 1
      grbit            |= fDspRwCol      << 2
      grbit            |= fFrozen        << 3
      grbit            |= fDspZeros      << 4
      grbit            |= fDefaultHdr    << 5
      grbit            |= fArabic        << 6
      grbit            |= fDspGuts       << 7
      grbit            |= fFrozenNoSplit << 8
      grbit            |= fSelected      << 9
      grbit            |= fPaged         << 10
      grbit            |= fBreakPreview  << 11

      header = [record, length].pack("vv")
      data    =[grbit, rwTop, colLeft, rgbHdr, wScaleSLV, wScaleNormal, reserved].pack("vvvVvvV")

      append(header, data)
   end

   ###############################################################################
   #
   # _store_page_view()
   #
   # Set page view mode. Only applicable to Mac Excel.
   #
   def store_page_view
      return if @page_view == 0
      data    = ['C8081100C808000000000040000000000900000000'].pack("H*")
      append(data)
   end
   
   ###############################################################################
   #
   # _store_tab_color()
   #
   # Write the Tab Color BIFF record.
   #
   def store_tab_color
      color   = @tab_color
   
      return if color == 0
   
      record  = 0x0862      # Record identifier
      length  = 0x0014      # Number of bytes to follow
   
      zero    = 0x0000
      unknown = 0x0014
   
      header = [record, length].pack("vv")
      data   = [record, zero, zero, zero, zero,
         zero, unknown, zero, color, zero].pack("vvvvvvvvvv")
   
      append(header, data)
   end
   
   ###############################################################################
   #
   # _store_defrow()
   #
   # Write BIFF record DEFROWHEIGHT.
   #
   def store_defrow
      record   = 0x0225      # Record identifier
      length   = 0x0004      # Number of bytes to follow
   
      grbit    = 0x0000      # Options.
      height   = 0x00FF      # Default row height
   
      header = [record, length].pack("vv")
      data   = [grbit,  height].pack("vv")
   
      prepend(header, data)
   end
   
   ###############################################################################
   #
   # _store_defcol()
   #
   # Write BIFF record DEFCOLWIDTH.
   #
   def store_defcol
      record   = 0x0055      # Record identifier
      length   = 0x0002      # Number of bytes to follow
   
      colwidth = 0x0008      # Default column width
   
      header   = pack("vv", record, length)
      data     = pack("v",  colwidth)
   
      prepend(header, data)
   end

   ###############################################################################
   #
   # _store_colinfo($firstcol, $lastcol, $width, $format, $hidden)
   #
   #   firstcol : First formatted column
   #   lastcol  : Last formatted column
   #   width    : Col width in user units, 8.43 is default
   #   format   : format object
   #   hidden   : hidden flag
   #   
   # Write BIFF record COLINFO to define column widths
   #
   # Note: The SDK says the record length is 0x0B but Excel writes a 0x0C
   # length record.
   #
   def store_colinfo(firstcol=0, lastcol=0, width=8.43, format=nil, hidden=0, level=0, collapsed=0)
      record   = 0x007D          # Record identifier
      length   = 0x000B          # Number of bytes to follow

      # Excel rounds the column width to the nearest pixel. Therefore we first
      # convert to pixels and then to the internal units. The pixel to users-units
      # relationship is different for values less than 1.
      #
      if width < 1
         pixels = width *12
      else
         pixels = width *7 +5
      end
      pixels = pixels.to_i
      
      coldx    = (pixels *256/7).to_i   # Col width in internal units
      grbit    = 0x0000               # Option flags
      reserved = 0x00                 # Reserved

      # Check for a format object
      if !format.nil? && format.kind_of?(Format)
         ixfe = format.get_xf_index
      else
         ixfe = 0x0F
      end

      # Set the limits for the outline levels (0 <= x <= 7).
      level = 0 if level < 0
      level = 7 if level > 7


      # Set the options flags. (See set_row() for more details).
      grbit |= 0x0001 if hidden != 0
      grbit |= level << 8
      grbit |= 0x1000 if collapsed != 0

      header = [record, length].pack("vv")
      data   = [firstcol, lastcol, coldx,
         ixfe, grbit, reserved].pack("vvvvvC")
   
      prepend(header, data)
   end

   ###############################################################################
   #
   # _store_filtermode()
   #
   # Write BIFF record FILTERMODE to indicate that the worksheet contains
   # AUTOFILTER record, ie. autofilters with a filter set.
   #
   def store_filtermode
      # Only write the record if the worksheet contains a filtered autofilter.
      return '' if @filter_on == 0

      record      = 0x009B      # Record identifier
      length      = 0x0000      # Number of bytes to follow

      header = [record, length].pack('vv')

      prepend(header)
   end


   ###############################################################################
   #
   # _store_autofilterinfo()
   #
   # Write BIFF record AUTOFILTERINFO.
   #
   def store_autofilterinfo
      # Only write the record if the worksheet contains an autofilter.
      return '' if @filter_count == 0

      record      = 0x009D      # Record identifier
      length      = 0x0002      # Number of bytes to follow
      num_filters = @filter_count

      header = [record, length].pack('vv')
      data   = [num_filters].pack('v')

      prepend(header, data)
   end


   ###############################################################################
   #
   # _store_selection($first_row, $first_col, $last_row, $last_col)
   #
   # Write BIFF record SELECTION.
   #
   def store_selection(first_row=0, first_col=0, last_row = nil, last_col =nil)
      record   = 0x001D                  # Record identifier
      length   = 0x000F                  # Number of bytes to follow

      pnn      = @active_pane   # Pane position
      rwAct    = first_row                   # Active row
      colAct   = first_col                   # Active column
      irefAct  = 0                       # Active cell ref
      cref     = 1                       # Number of refs

      rwFirst  = first_row                   # First row in reference
      colFirst = first_col                   # First col in reference
      rwLast   = last_row || rwFirst       # Last  row in reference
      colLast  = last_col || colFirst      # Last  col in reference

      # Swap last row/col for first row/col as necessary
      if rwFirst > rwLast
         tmp = rwFirst
         rwFirst = rwLast
         rwLast = tmp
      end

      if colFirst > colLast
         tmp = colFirst
         colFirst = colLast
         colLast = tmp
      end

      header = [record, length].pack('vv')
      data = [pnn, rwAct, colAct, irefAct, cref,
         rwFirst, rwLast, colFirst, colLast].pack('CvvvvvvCC')

      append(header, data)
   end


   ###############################################################################
   #
   # _store_externcount($count)
   #
   # Write BIFF record EXTERNCOUNT to indicate the number of external sheet
   # references in a worksheet.
   #
   # Excel only stores references to external sheets that are used in formulas.
   # For simplicity we store references to all the sheets in the workbook
   # regardless of whether they are used or not. This reduces the overall
   # complexity and eliminates the need for a two way dialogue between the formula
   # parser the worksheet objects.
   #
   def store_externcount(count)
      record   = 0x0016          # Record identifier
      length   = 0x0002          # Number of bytes to follow

      cxals    = count           # Number of external references

      header = [record, length].pack('vv')
      data   = [cxals].pack('v')

      prepend(header, data)
   end


   ###############################################################################
   #
   # _store_externsheet($sheetname)
   #    sheetname  : Worksheet name
   #
   # Writes the Excel BIFF EXTERNSHEET record. These references are used by
   # formulas. A formula references a sheet name via an index. Since we store a
   # reference to all of the external worksheets the EXTERNSHEET index is the same
   # as the worksheet index.
   #
   def store_externsheet(sheetname)
      record    = 0x0017         # Record identifier
      # length;                     # Number of bytes to follow

      # cch                        # Length of sheet name
      # rgch                       # Filename encoding

      # References to the current sheet are encoded differently to references to
      # external sheets.
      #
      if @name == sheetname
         sheetname = ''
         length    = 0x02  # The following 2 bytes
         cch       = 1     # The following byte
         rgch      = 0x02  # Self reference
      else
         length    = 0x02 + sheetname.length
         cch       = sheetname.length
         rgch      = 0x03  # Reference to a sheet in the current workbook
      end

      header = [record, length].pack('vv')
      data   = [cch, rgch].pack('CC')

      prepend(header, data, sheetname)
   end


   ###############################################################################
   #
   # _store_panes(y, x, colLeft, no_split, pnnAct)
   #    y           = args[0] || 0   # Vertical split position
   #    x           = $_[1] || 0;   # Horizontal split position
   #    my $rwTop       = $_[2];        # Top row visible
   #    my $colLeft     = $_[3];        # Leftmost column visible
   #    my $no_split    = $_[4];        # No used here.
   #    my $pnnAct      = $_[5];        # Active pane
   #
   #
   # Writes the Excel BIFF PANE record.
   # The panes can either be frozen or thawed (unfrozen).
   # Frozen panes are specified in terms of a integer number of rows and columns.
   # Thawed panes are specified in terms of Excel's units for rows and columns.
   #
   def store_panes(y, x, colLeft, no_split, pnnAct)
      record      = 0x0041       # Record identifier
      length      = 0x000A       # Number of bytes to follow

      y = 0 if y.nil?
      x = 0 if x.nil?

      # Code specific to frozen or thawed panes.
      if @frozen != 0
         # Set default values for $rwTop and $colLeft
         rwTop   = y unless defined? rwTop
         colLeft = x unless defined? colLeft
      else
         # Set default values for $rwTop and $colLeft
         rwTop   = 0  unless defined? rwTop
         colLeft = 0  unless defined? colLeft

         # Convert Excel's row and column units to the internal units.
         # The default row height is 12.75
         # The default column width is 8.43
         # The following slope and intersection values were interpolated.
         #
         y = 20*y      + 255
         x = 113.879*x + 390
      end


      # Determine which pane should be active. There is also the undocumented
      # option to override this should it be necessary: may be removed later.
      #
      unless defined? pnnAct
         pnnAct = 0 if (x != 0 && y != 0) # Bottom right
         pnnAct = 1 if (x != 0 && y == 0) # Top right
         pnnAct = 2 if (x == 0 && y != 0) # Bottom left
         pnnAct = 3 if (x == 0 && y == 0) # Top left
      end

      @active_pane = pnnAct # Used in _store_selection

      header = [record, length].pack('vv')
      data   = [x, y, rwTop, colLeft, pnnAct].pack('vvvvv')

      append(header, data)
   end


   ###############################################################################
   #
   # _store_setup()
   #
   # Store the page setup SETUP BIFF record.
   #
   def store_setup
      record       = 0x00A1                  # Record identifier
      length       = 0x0022                  # Number of bytes to follow

      iPaperSize   = @paper_size    # Paper size
      iScale       = @print_scale   # Print scaling factor
      iPageStart   = @page_start    # Starting page number
      iFitWidth    = @fit_width     # Fit to number of pages wide
      iFitHeight   = @fit_height    # Fit to number of pages high
      grbit        = 0x00                    # Option flags
      iRes         = 0x0258                  # Print resolution
      iVRes        = 0x0258                  # Vertical print resolution
      numHdr       = @margin_header # Header Margin
      numFtr       = @margin_footer # Footer Margin
      iCopies      = 0x01                    # Number of copies

      fLeftToRight = @page_order    # Print over then down
      fLandscape   = @orientation   # Page orientation
      fNoPls       = 0x0                     # Setup not read from printer
      fNoColor     = @black_white   # Print black and white
      fDraft       = @draft_quality # Print draft quality
      fNotes       = @print_comments# Print notes
      fNoOrient    = 0x0            # Orientation not set
      fUsePage     = @custom_start  # Use custom starting page

      grbit           = fLeftToRight
      grbit          |= fLandscape    << 1
      grbit          |= fNoPls        << 2
      grbit          |= fNoColor      << 3
      grbit          |= fDraft        << 4
      grbit          |= fNotes        << 5
      grbit          |= fNoOrient     << 6
      grbit          |= fUsePage      << 7


      numHdr = [numHdr].pack('d')
      numFtr = [numFtr].pack('d')

      if @byte_order != 0
         numHdr = numHdr.reverse
         numFtr = numFtr.reverse
      end

      header = [record, length].pack('vv')
      data1  = [iPaperSize, iScale, iPageStart,
         iFitWidth, iFitHeight, grbit, iRes, iVRes].pack("vvvvvvvv")

      data2  = numHdr + numFtr
      data3  = [iCopies].pack('v')

      prepend(header, data1, data2, data3)

   end

   ###############################################################################
   #
   # _store_header()
   #
   # Store the header caption BIFF record.
   #
   def store_header
      record      = 0x0014                       # Record identifier
      # length                                     # Bytes to follow

      str         = @header             # header string
      cch         = str.length                 # Length of header string
      encoding    = @header_encoding    # Character encoding


      # Character length is num of chars not num of bytes
      cch           /= 2 if encoding != 0

      # Change the UTF-16 name from BE to LE
      str            = [str].unpack('v*').pack('n*') if encoding != 0

      length         = 3 + str.length

      header      = [record, length].pack('vv')
      data        =  [cch, encoding].pack('vC')

      prepend(header, data, str)
   end


   ###############################################################################
   #
   # _store_footer()
   #
   # Store the footer caption BIFF record.
   #
   def store_footer
      record      = 0x0015                       # Record identifier
      # length;                                     # Bytes to follow

      str         = @footer             # footer string
      cch         = str.length                 # Length of ooter string
      encoding    = @footer_encoding    # Character encoding


      # Character length is num of chars not num of bytes
      cch           /= 2 if encoding != 0

      # Change the UTF-16 name from BE to LE
      str            = [str].unpack('v*').pack('n*')

      length         = 3 + str.length

      header      = [record, length].pack('vv')
      data        =  [cch, encoding].pack('vC')

      prepend(header, data, str)
   end


   ###############################################################################
   #
   # _store_hcenter()
   #
   # Store the horizontal centering HCENTER BIFF record.
   #
   def store_hcenter
      record   = 0x0083              # Record identifier
      length   = 0x0002              # Bytes to follow

      fHCenter = @hcenter   # Horizontal centering

      header      = [record, length].pack('vv')
      data      = [fHCenter].pack('v')

      prepend(header, data)
   end


   ###############################################################################
   #
   # _store_vcenter()
   #
   # Store the vertical centering VCENTER BIFF record.
   #
   def store_vcenter
      record   = 0x0084              # Record identifier
      length   = 0x0002              # Bytes to follow

      mfVCenter = @vcenter   # Horizontal centering

      header      = [record, length].pack('vv')
      data      = [mfVCenter].pack('v')

      prepend(header, data)
   end


   ###############################################################################
   #
   # _store_margin_left()
   #
   # Store the LEFTMARGIN BIFF record.
   #
   def store_margin_left
      record  = 0x0026                   # Record identifier
      length  = 0x0008                   # Bytes to follow

      margin  = @margin_left    # Margin in inches

      header  = [record, length].pack('vv')
      data    = [margin].pack('d')

      data = data.reverse if @byte_order != 0

      prepend(header, data)
   end


   ###############################################################################
   #
   # _store_margin_right()
   #
   # Store the RIGHTMARGIN BIFF record.
   #
   def store_margin_right
      record  = 0x0027                   # Record identifier
      length  = 0x0008                   # Bytes to follow

      margin  = @margin_right   # Margin in inches

      header  = [record, length].pack('vv')
      data    = [margin].pack('d')

      data = data.reverse if @byte_order != 0

      prepend(header, data)
   end


   ###############################################################################
   #
   # _store_margin_top()
   #
   # Store the TOPMARGIN BIFF record.
   #
   def store_margin_top
      record  = 0x0028                   # Record identifier
      length  = 0x0008                   # Bytes to follow

      margin  = @margin_top     # Margin in inches

      header  = [record, length].pack('vv')
      data    = [margin].pack('d')

      data = data.reverse if @byte_order != 0

      prepend(header, data)
   end


   ###############################################################################
   #
   # _store_margin_bottom()
   #
   # Store the BOTTOMMARGIN BIFF record.
   #
   def store_margin_bottom
      record  = 0x0029                   # Record identifier
      length  = 0x0008                   # Bytes to follow

      margin  = @margin_bottom  # Margin in inches

      header  = [record, length].pack('vv')
      data    = [margin].pack('d')

      data = data.reverse if @byte_order != 0

      prepend(header, data)
   end

   ###############################################################################
   #
   # merge_cells($first_row, $first_col, $last_row, $last_col)
   #
   # This is an Excel97/2000 method. It is required to perform more complicated
   # merging than the normal align merge in Format.pm
   #
   def merge_cells(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end

      record  = 0x00E5                    # Record identifier
      length  = 0x000A                    # Bytes to follow

      cref     = 1                        # Number of refs
      rwFirst  = args[0]                  # First row in reference
      colFirst = args[1]                  # First col in reference
      rwLast   = args[2] || rwFirst       # Last  row in reference
      colLast  = args[3] || colFirst      # Last  col in reference

      # Excel doesn't allow a single cell to be merged
      return if rwFirst == rwLast and colFirst == colLast

      # Swap last row/col with first row/col as necessary
      rwFirst,  rwLast  = rwLast,  rwFirst  if rwFirst  > rwLast
      colFirst, colLast = colLast, colFirst if colFirst > colLast

      header   = [record, length].pack("vv")
      data     = [cref, rwFirst, rwLast, colFirst, colLast].pack("vvvvv")

      append(header, data)
   end

   ###############################################################################
   #
   # merge_range($row1, $col1, $row2, $col2, $string, $format, $encoding)
   #
   # This is a wrapper to ensure correct use of the merge_cells method, i.e., write
   # the first cell of the range, write the formatted blank cells in the range and
   # then call the merge_cells record. Failing to do the steps in this order will
   # cause Excel 97 to crash.
   #
   def merge_range(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end
      raise "Incorrect number of arguments" if args.size != 6 and args.size != 7
      raise "Format argument is not a format object" unless args[5].kind_of?(Format)

      rwFirst  = args[0]
      colFirst = args[1]
      rwLast   = args[2]
      colLast  = args[3]
      string   = args[4]
      format   = args[5]
      encoding = args[6] ? 1 : 0

      # Temp code to prevent merged formats in non-merged cells.
      error = "Error: refer to merge_range() in the documentation. " +
        "Can't use previously non-merged format in merged cells"

      raise error if format.used_merge == -1
      format.used_merge = 0   # Until the end of this function.

      # Set the merge_range property of the format object. For BIFF8+.
      format.set_merge_range

      # Excel doesn't allow a single cell to be merged
      raise "Can't merge single cell" if rwFirst  == rwLast and
        colFirst == colLast

      # Swap last row/col with first row/col as necessary
      rwFirst,  rwLast  = rwLast,  rwFirst  if rwFirst  > rwLast
      colFirst, colLast = colLast, colFirst if colFirst > colLast

      # Write the first cell
      if encoding != 0
         write_utf16be_string(rwFirst, colFirst, string, format)
      else
         write(rwFirst, colFirst, string, format)
      end

      # Pad out the rest of the area with formatted blank cells.
      (rwFirst .. rwLast).each do |row|
         (colFirst .. colLast).each do |col|
            next if row == rwFirst and col == colFirst
            write_blank(row, col, format)
         end
      end

      merge_cells(rwFirst, colFirst, rwLast, colLast)

      # Temp code to prevent merged formats in non-merged cells.
      format.used_merge = 1
   end

   ###############################################################################
   #
   # _store_print_headers()
   #
   # Write the PRINTHEADERS BIFF record.
   #
   def store_print_headers
      record      = 0x002a                   # Record identifier
      length      = 0x0002                   # Bytes to follow

      fPrintRwCol = @print_headers           # Boolean flag

      header      = [record, length].pack("vv")
      data        = [fPrintRwCol].pack("v")

      prepend(header, data)
   end

   ###############################################################################
   #
   # _store_print_gridlines()
   #
   # Write the PRINTGRIDLINES BIFF record. Must be used in conjunction with the
   # GRIDSET record.
   #
   def store_print_gridlines
      record      = 0x002b                    # Record identifier
      length      = 0x0002                    # Bytes to follow

      fPrintGrid  = @print_gridlines          # Boolean flag

      header      = [record, length].pack("vv")
      data        = [fPrintGrid].pack("v")

      prepend(header, data)
   end

   ###############################################################################
   #
   # _store_gridset()
   #
   # Write the GRIDSET BIFF record. Must be used in conjunction with the
   # PRINTGRIDLINES record.
   #
   def store_gridset
      record      = 0x0082                        # Record identifier
      length      = 0x0002                        # Bytes to follow

      fGridSet    = !@print_gridlines          # Boolean flag

      header      = [record, length].pack("vv")
      data        = [fGridSet].pack("v")

      prepend(header, data)
   end
   
   ###############################################################################
   #
   # _store_guts()
   #
   # Write the GUTS BIFF record. This is used to configure the gutter margins
   # where Excel outline symbols are displayed. The visibility of the gutters is
   # controlled by a flag in WSBOOL. See also _store_wsbool().
   #
   # We are all in the gutter but some of us are looking at the stars.
   #
   def store_guts
      record      = 0x0080   # Record identifier
      length      = 0x0008   # Bytes to follow

      dxRwGut     = 0x0000   # Size of row gutter
      dxColGut    = 0x0000   # Size of col gutter

      row_level   = @outline_row_level
      col_level   = 0


      # Calculate the maximum column outline level. The equivalent calculation
      # for the row outline level is carried out in set_row().
      #
      @colinfo.each do |colinfo|
         # Skip cols without outline level info.
         next if colinfo.size < 6
         col_level = colinfo[5] if colinfo[5] > col_level
      end

      # Set the limits for the outline levels (0 <= x <= 7).
      col_level = 0 if col_level < 0
      col_level = 7 if col_level > 7

      # The displayed level is one greater than the max outline levels
      row_level = row_level + 1 if row_level > 0
      col_level = col_level + 1 if col_level > 0

      header = [record, length].pack("vv")
      data   = [dxRwGut, dxColGut, row_level, col_level].pack("vvvv")

      prepend(header, data)
   end

   ###############################################################################
   #
   # _store_wsbool()
   #
   # Write the WSBOOL BIFF record, mainly for fit-to-page. Used in conjunction
   # with the SETUP record.
   #
   def store_wsbool
      record      = 0x0081   # Record identifier
      length      = 0x0002   # Bytes to follow
   
      grbit       = 0x0000   # Option flags
   
      # Set the option flags
      grbit |= 0x0001                        # Auto page breaks visible
      grbit |= 0x0020 if @outline_style != 0 # Auto outline styles
      grbit |= 0x0040 if @outline_below != 0 # Outline summary below
      grbit |= 0x0080 if @outline_right != 0 # Outline summary right
      grbit |= 0x0100 if @fit_page      != 0 # Page setup fit to page
      grbit |= 0x0400 if @outline_on    != 0 # Outline symbols displayed
   
      header = [record, length].pack("vv")
      data   = [grbit].pack('v')
   
      prepend(header, data)
   end
   
   ###############################################################################
   #
   # _store_hbreak()
   #
   # Write the HORIZONTALPAGEBREAKS BIFF record.
   #
   def store_hbreak
      # Return if the user hasn't specified pagebreaks
      return if @hbreaks.size == 0
   
      # Sort and filter array of page breaks
      breaks  = sort_pagebreaks(@hbreaks)
   
      record  = 0x001b               # Record identifier
      cbrk    = breaks.size          # Number of page breaks
      length  = 2 + 6 * cbrk         # Bytes to follow
   
      header = [record, length].pack("vv")
      data   = [cbrk].pack("v")

      # Append each page break
      breaks.each do |brk|
         data = data + [brk, 0x0000, 0x00ff].pack("vvv")
      end
   
      prepend(header, data)
   end
   
   ###############################################################################
   #
   # _store_vbreak()
   #
   # Write the VERTICALPAGEBREAKS BIFF record.
   #
   def store_vbreak
      # Return if the user hasn't specified pagebreaks
      return if @vbreaks.size == 0
   
      # Sort and filter array of page breaks
      breaks  = sort_pagebreaks(@vbreaks)
   
      record  = 0x001a               # Record identifier
      cbrk    = breaks.size          # Number of page breaks
      length  = 2 + 6*cbrk           # Bytes to follow
   
      header = [record, length].pack("vv")
      data   = [cbrk].pack("v")
   
      # Append each page break
      breaks.each do |brk|
         data = data + [brk, 0x0000, 0x00ff].pack("vvv")
      end
   
      prepend(header, data)
   end
   
   
   ###############################################################################
   #
   # _store_protect()
   #
   # Set the Biff PROTECT record to indicate that the worksheet is protected.
   #
   def store_protect
      # Exit unless sheet protection has been specified
      return if @protect == 0
   
      record      = 0x0012               # Record identifier
      length      = 0x0002               # Bytes to follow
   
      fLock       = @protect             # Worksheet is protected
   
      header = [record, length].pack("vv")
      data   = [fLock].pack("v")
   
      prepend(header, data)
   end
   
   ###############################################################################
   #
   # _store_obj_protect()
   #
   # Set the Biff OBJPROTECT record to indicate that objects are protected.
   #
   def store_obj_protect
      # Exit unless sheet protection has been specified
      return if @protect == 0
   
      record      = 0x0063               # Record identifier
      length      = 0x0002               # Bytes to follow
   
      fLock       = @protect             # Worksheet is protected
   
      header = [record, length].pack("vv")
      data   = [fLock].pack("v")
   
      prepend(header, data)
   end
   
   ###############################################################################
   #
   # _store_password()
   #
   # Write the worksheet PASSWORD record.
   #
   def store_password
      # Exit unless sheet protection and password have been specified
      return if (@protect == 0 or @password.nil?)
   
      record      = 0x0013               # Record identifier
      length      = 0x0002               # Bytes to follow
   
      wPassword   = @password            # Encoded password
   
      header = [record, length].pack("vv")
      data   = [wPassword].pack("v")
   
      prepend(header, data)
   end
   
   #
   # Note about compatibility mode.
   #
   # Excel doesn't require every possible Biff record to be present in a file.
   # In particular if the indexing records INDEX, ROW and DBCELL aren't present
   # it just ignores the fact and reads the cells anyway. This is also true of
   # the EXTSST record. Gnumeric and OOo also take this approach. This allows
   # WriteExcel to ignore these records in order to minimise the amount of data
   # stored in memory. However, other third party applications that read Excel
   # files often expect these records to be present. In "compatibility mode"
   # WriteExcel writes these records and tries to be as close to an Excel
   # generated file as possible.
   #
   # This requires additional data to be stored in memory until the file is
   # about to be written. This incurs a memory and speed penalty and may not be
   # suitable for very large files.
   #
   
   
   
   ###############################################################################
   #
   # _store_table()
   #
   # Write cell data stored in the worksheet row/col table.
   #
   # This is only used when compatibity_mode() is in operation.
   #
   # This method writes ROW data, then cell data (NUMBER, LABELSST, etc) and then
   # DBCELL records in blocks of 32 rows. This is explained in detail (for a
   # change) in the Excel SDK and in the OOo Excel file format doc.
   #
   def store_table
      return if @compatibility == 0
   
      # Offset from the DBCELL record back to the first ROW of the 32 row block.
      row_offset = 0
   
      # Track rows that have cell data or modified by set_row().
      written_rows = []
   
   
      # Write the ROW records with updated max/min col fields.
      #
      (0 .. @dim_rowmax-1).each do |row|
         # Skip unless there is cell data in row or the row has been modified.
         next unless @table[row] or @row_data[row]
   
         # Store the rows with data.
         written_rows.push(row)
   
         # Increase the row offset by the length of a ROW record;
         row_offset += 20
   
         # The max/min cols in the ROW records are the same as in DIMENSIONS.
         col_min = @dim_colmin
         col_max = @dim_colmax
   
         # Write a user specifed ROW record (modified by set_row()).
         if @row_data[row]
            # Rewrite the min and max cols for user defined row record.
            packed_row = @row_data[row]
            packed_row[6..9] = [col_min, col_max].pack('vv')
            append(packed_row)
         else
            # Write a default Row record if there isn't a  user defined ROW.
            write_row_default(row, col_min, col_max)
         end
   
         # If 32 rows have been written or we are at the last row in the
         # worksheet then write the cell data and the DBCELL record.
         #
         if written_rows.size == 32 or row == @dim_rowmax -1
            # Offsets to the first cell of each row.
            cell_offsets = []
            cell_offsets.push(row_offset - 20)
   
            # Write the cell data in each row and sum their lengths for the
            # cell offsets.
            #
            written_rows.each do |row|
               cell_offset = 0
   
                   
               @table[row].each do |col|
                  next unless col
                  append(col)
                  ength = col.length
                  row_offset  = row_offset  + length
                  cell_offset = cell_offset + length
               end
               cell_offsets.push(cell_offset)
            end
   
            # The last offset isn't required.
            cell_offsets.pop
   
            # Stores the DBCELL offset for use in the INDEX record.
            @db_indices.push(@datasize)
   
            # Write the DBCELL record.
            store_dbcell(row_offset, cell_offsets)
   
            # Clear the variable for the next block of rows.
            written_rows   = []
            cell_offsets   = []
            row_offset     = 0
         end
      end
   end
   
   ###############################################################################
   #
   # _store_dbcell()
   #
   # Store the DBCELL record using the offset calculated in _store_table().
   #
   # This is only used when compatibity_mode() is in operation.
   #
   def store_dbcell(row_offset, cell_offsets)
      record          = 0x00D7                     # Record identifier
      length          = 4 + 2 * cell_offsets.size  # Bytes to follow
   
      header          = [record, length].pack('vv')
      data            = [row_offset].pack('V')
      cell_offsets.each do |co|
         data = data + [co].pack('v')
      end
   
      append(header, data)
   end
   
   
   ###############################################################################
   #
   # _store_index()
   #
   # Store the INDEX record using the DBCELL offsets calculated in _store_table().
   #
   # This is only used when compatibity_mode() is in operation.
   #
   def store_index
      return if @compatibility == 0
   
      indices     = @db_indices
      reserved    = 0x00000000
      row_min     = @dim_rowmin
      row_max     = @dim_rowmax
   
      record      = 0x020B                 # Record identifier
      length      = 16 + 4 * indices.size  # Bytes to follow
   
      header      = [record, length].pack('vv')
      data        = [reserved, row_min, row_max, reserved].pack('VVVV')
   
      indices.each do |index|
         data = data + [index + @offset + 20 + length + 4].pack('V')
      end
   
      prepend(header, data)
   end

   ###############################################################################
   #
   # embed_chart($row, $col, $filename, $x, $y, $scale_x, $scale_y)
   #
   # Embed an extracted chart in a worksheet.
   #
   def embed_chart(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end
   
      row         = args[0]
      col         = args[1]
      chart       = args[2]
      x_offset    = args[3] || 0
      y_offset    = args[4] || 0
      scale_x     = args[5] || 1
      scale_y     = args[6] || 1
   
      raise "Insufficient arguments in embed_chart()" unless args.size >= 3
      #       raise "Couldn't locate $chart: $!"              unless -e $chart;
   
      @charts[row][col] =  [row, col, chart,
         x_offset, y_offset, scale_x, scale_y, ]
   
   end
   
   ###############################################################################
   #
   # insert_image($row, $col, $filename, $x, $y, $scale_x, $scale_y)
   #
   # Insert an image into the worksheet.
   #
   def insert_image(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end
   
      row         = args[0]
      col         = args[1]
      image       = args[2]
      x_offset    = args[3] || 0
      y_offset    = args[4] || 0
      scale_x     = args[5] || 1
      scale_y     = args[6] || 1
   
      raise "Insufficient arguments in insert_image()" unless args.size >= 3
      raise "Couldn't locate #{image}: $!"             unless test(?e, image)
   
      @images[row][col] = [ row, col, image,
         x_offset, y_offset, scale_x, scale_y,]
   
   end
   
   # Older method name for backwards compatibility.
   #   *insert_bitmap = *insert_image;
   
   
   ###############################################################################
   #
   #  _position_object()
   #
   # Calculate the vertices that define the position of a graphical object within
   # the worksheet.
   #
   #         +------------+------------+
   #         |     A      |      B     |
   #   +-----+------------+------------+
   #   |     |(x1,y1)     |            |
   #   |  1  |(A1)._______|______      |
   #   |     |    |              |     |
   #   |     |    |              |     |
   #   +-----+----|    BITMAP    |-----+
   #   |     |    |              |     |
   #   |  2  |    |______________.     |
   #   |     |            |        (B2)|
   #   |     |            |     (x2,y2)|
   #   +---- +------------+------------+
   #
   # Example of a bitmap that covers some of the area from cell A1 to cell B2.
   #
   # Based on the width and height of the bitmap we need to calculate 8 vars:
   #     $col_start, $row_start, $col_end, $row_end, $x1, $y1, $x2, $y2.
   # The width and height of the cells are also variable and have to be taken into
   # account.
   # The values of $col_start and $row_start are passed in from the calling
   # function. The values of $col_end and $row_end are calculated by subtracting
   # the width and height of the bitmap from the width and height of the
   # underlying cells.
   # The vertices are expressed as a percentage of the underlying cell width as
   # follows (rhs values are in pixels):
   #
   #       x1 = X / W *1024
   #       y1 = Y / H *256
   #       x2 = (X-1) / W *1024
   #       y2 = (Y-1) / H *256
   #
   #       Where:  X is distance from the left side of the underlying cell
   #               Y is distance from the top of the underlying cell
   #               W is the width of the cell
   #               H is the height of the cell
   #
   # Note: the SDK incorrectly states that the height should be expressed as a
   # percentage of 1024.
   #
   def position_object(col_start, row_start, x1, y1, width, height)
      # col_start;  # Col containing upper left corner of object
      # x1;         # Distance to left side of object
   
      # row_start;  # Row containing top left corner of object
      # y1;         # Distance to top of object
   
      # col_end;    # Col containing lower right corner of object
      # x2;         # Distance to right side of object
   
      # row_end;    # Row containing bottom right corner of object
      # y2;         # Distance to bottom of object
   
      # width;      # Width of image frame
      # height;     # Height of image frame
   
      # Adjust start column for offsets that are greater than the col width
      while x1 >= size_col(col_start)
         x1 = x1 - size_col(col_start)
         col_start = col_start + 1
      end
   
      # Adjust start row for offsets that are greater than the row height
      while y1 >= size_row(row_start)
         y1 = y1 - size_row(row_start)
         row_start = row_start + 1
      end
   
      # Initialise end cell to the same as the start cell
      col_end    = col_start
      row_end    = row_start
   
      width      = width  + x1 -1
      height     = height + y1 -1
   
      # Subtract the underlying cell widths to find the end cell of the image
      while width >= size_col(col_end)
         width = width - size_col(col_end)
         col_end = col_end + 1
      end
   
      # Subtract the underlying cell heights to find the end cell of the image
      while height >= size_row(row_end)
         height  = height - size_row(row_end)
         row_end = row_end + 1
      end
   
      # Bitmap isn't allowed to start or finish in a hidden cell, i.e. a cell
      # with zero eight or width.
      #
      return if size_col(col_start) == 0
      return if size_col(col_end)   == 0
      return if size_row(row_start) == 0
      return if size_row(row_end)   == 0
   
      # Convert the pixel values to the percentage value expected by Excel
      x1 = x1     / size_col(col_start)   * 1024
      y1 = y1     / size_row(row_start)   *  256
      x2 = width  / size_col(col_end)     * 1024
      y2 = height / size_row(row_end)     *  256
   
      # Simulate ceil() without calling POSIX::ceil().
      x1 = (x1 +0.5).to_i
      y1 = (y1 +0.5).to_i
      x2 = (x2 +0.5).to_i
      y2 = (y2 +0.5).to_i
   
      return [col_start, x1,
         row_start, y1,
         col_end,   x2,
         row_end,   y2
      ]
   end

   ###############################################################################
   #
   # _size_col($col)
   #
   # Convert the width of a cell from user's units to pixels. Excel rounds the
   # column width to the nearest pixel. If the width hasn't been set by the user
   # we use the default value. If the column is hidden we use a value of zero.
   #
   def size_col(col)
      # Look up the cell value to see if it has been changed
      unless @col_sizes[col].nil?
         width = @col_sizes[col]
   
         # The relationship is different for user units less than 1.
         if width < 1
            return (width *12).to_i
         else
            return (width *7 +5 ).to_i
         end
      else
         return 64
      end
   end
   
   ###############################################################################
   #
   # _size_row($row)
   #
   # Convert the height of a cell from user's units to pixels. By interpolation
   # the relationship is: y = 4/3x. If the height hasn't been set by the user we
   # use the default value. If the row is hidden we use a value of zero. (Not
   # possible to hide row yet).
   #
   def size_row(row)
      # Look up the cell value to see if it has been changed
      unless @row_sizes[row].nil?
         if @row_sizes[row] == 0
            return 0
         else
            return (4/3 * @row_sizes[row]).to_i
         end
      else
         return 17
      end
   end

   ###############################################################################
   #
   # _store_zoom($zoom)
   #
   #
   # Store the window zoom factor. This should be a reduced fraction but for
   # simplicity we will store all fractions with a numerator of 100.
   #
   def store_zoom
      # If scale is 100 we don't need to write a record
      return if @zoom == 100
   
      record      = 0x00A0               # Record identifier
      length      = 0x0004               # Bytes to follow
   
      header      = [record, header].pack("vv")
      data        = [@zoom, 100].pack("vv")
   
      append(header, data)
   end
   
   ###############################################################################
   #
   # write_utf16be_string($row, $col, $string, $format)
   #
   # Write a Unicode string to the specified row and column (zero indexed).
   # $format is optional.
   # Returns  0 : normal termination
   #         -1 : insufficient number of arguments
   #         -2 : row or column out of range
   #         -3 : long string truncated to 255 chars
   #
   def write_utf16be_string(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end
   
      return -1 if (args.size < 3)                     # Check the number of args
   
      record      = 0x00FD                        # Record identifier
      length      = 0x000A                        # Bytes to follow
   
      row         = args[0]                         # Zero indexed row
      col         = args[1]                         # Zero indexed column
      strlen      = args[2].length
      str         = args[2]
      xf          = xf_record_index(row, col, args[3]) # The cell format
      encoding    = 0x1
      str_error   = 0
   
      # Check that row and col are valid and store max and min values
      return -2 if check_dimensions(row, col) != 0
   
      # Limit the utf16 string to the max number of chars (not bytes).
      if strlen > 32767* 2
         str       = str[0..32767*2]
         str_error = -3
      end
   
      num_bytes = str.length
      num_chars = (num_bytes / 2).to_i
   
      # Check for a valid 2-byte char string.
      raise "Uneven number of bytes in Unicode string" if num_bytes % 2
   
      # Change from UTF16 big-endian to little endian
      str = str.unpack('n*').pack('v')
   
      # Add the encoding and length header to the string.
      str_header  = [num_chars, encoding].pack("vC")
      str         = str_header + str
   
      if @str_table[str].nil?
         @str_table[str] = @str_unique
         @str_unique = @str_unique + 1
      end
   
      @str_total = @str_total + 1
   
      header = [record, length].pack("vv")
      data   = [row, col, xf, @str_table[str]].pack("vvvV")
   
      # Store the data or write immediately depending on the compatibility mode.
      if @compatibility != 0
         tmp = []
         tmp[col] = header + data
         @table[row] = tmp
      else
         append(header, data)
      end
   
      return str_error
   end
   
   ###############################################################################
   #
   # write_utf16le_string($row, $col, $string, $format)
   #
   # Write a UTF-16LE string to the specified row and column (zero indexed).
   # $format is optional.
   # Returns  0 : normal termination
   #         -1 : insufficient number of arguments
   #         -2 : row or column out of range
   #         -3 : long string truncated to 255 chars
   #
   def write_utf16le_string(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
         args = substitute_cellref(*args)
      end
   
      return -1 if (args.size < 3)                     # Check the number of args
   
      record      = 0x00FD                          # Record identifier
      length      = 0x000A                          # Bytes to follow
   
      row         = args[0]                         # Zero indexed row
      col         = args[1]                         # Zero indexed column
      str         = args[2]
      format      = args[3]                         # The cell format
   
      # Change from UTF16 big-endian to little endian
      str = str.unpack('n*').pack("v*")
   
      return write_utf16be_string(row, col, str, format)
   end
   
   
   # Older method name for backwards compatibility.
   #   *write_unicode    = *write_utf16be_string;
   #   *write_unicode_le = *write_utf16le_string;

   ###############################################################################
   #
   # _store_autofilters()
   #
   # Function to iterate through the columns that form part of an autofilter
   # range and write Biff AUTOFILTER records if a filter expression has been set.
   #
   def store_autofilters
      # Skip all columns if no filter have been set.
      return '' if @filter_on == 0

      col1 = @filter_area[2]
      col2 = @filter_area[3]

      i = col1
      while i <= col2
         # Reverse order since records are being pre-pended.
         col = col2 -i

         # Skip if column doesn't have an active filter.
         next unless @filter_cols[col]

         # Retrieve the filter tokens and write the autofilter records.
         store_autofilter(col, @filter_cols[col])
      end
   end

   ###############################################################################
   #
   # _store_autofilter()
   #
   # Function to write worksheet AUTOFILTER records. These contain 2 Biff Doper
   # structures to represent the 2 possible filter conditions.
   #
   def store_autofilter(index, operator, token_1, join, operator_2, token_2)
      record          = 0x009E
      length          = 0x0000

      top10_active    = 0
      top10_direction = 0
      top10_percent   = 0
      top10_value     = 101

      grbit       = join
      optimised_1 = 0
      optimised_2 = 0
      doper_1     = ''
      doper_2     = ''
      string_1    = ''
      string_2    = ''

      # Excel used an optimisation in the case of a simple equality.
      optimised_1 = 1 if                         operator_1 == 2
      optimised_2 = 1 if defined operator_2 and operator_2 == 2

      # Convert non-simple equalities back to type 2. See  _parse_filter_tokens().
      operator_1 = 2 if                        operator_1 == 22
      operator_2 = 2 if defined operator_2 and operator_2 == 22

      # Handle a "Top" style expression.
      if operator_1 >= 30
         # Remove the second expression if present.
         operator_2 = nil
         token_2    = nil

         # Set the active flag.
         top10_active    = 1

         if (operator_1 == 30 or operator_1 == 31)
            top10_direction = 1
         end

         if (operator_1 == 31 or operator_1 == 33)
            top10_percent = 1
         end

         if (top10_direction == 1)
            operator_1 = 6
         else
            operator_1 = 3
         end

         top10_value     = token_1
         token_1         = 0
      end

      grbit     |= optimised_1      << 2
      grbit     |= optimised_2      << 3
      grbit     |= top10_active     << 4
      grbit     |= top10_direction  << 5
      grbit     |= top10_percent    << 6
      grbit     |= top10_value      << 7

      doper_1, string_1 = pack_doper(operator_1, token_1)
      doper_2, string_2   = pack_doper(operator_2, token_2)

      data = [index].pack('v')
      data = data + [grbit].pack('v')
      data = data + doper_1 + doper_2 + string_1 + string_2

      length  = data.length
      header  = [record, length].pack('vv')

      prepend(header, data)
   end

   ###############################################################################
   #
   # _pack_doper()
   #
   # Create a Biff Doper structure that represents a filter expression. Depending
   # on the type of the token we pack an Empty, String or Number doper.
   #
   def pack_doper(operator, token)
      doper       = ''
      string      = ''

      # Return default doper for non-defined filters.
      if operator.nil?
         return @pack_unused_doper, string
      end

      if token =~ /^blanks|nonblanks$/i
         doper  = pack_blanks_doper(operator, token)
      elsif operator == 2 or
           !(token  =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/)
         # Excel treats all tokens as strings if the operator is equality, =.
         string = token

         encoding = 0
         length   = string.length

         # Handle utf8 strings in perl 5.8.
         #  if ($] >= 5.008) {
         #      require Encode;
         #
         #      if (Encode::is_utf8($string)) {
         #          $string = Encode::encode("UTF-16BE", $string);
         #          $encoding = 1;
         #      }
         #  }

         string = [encoding].pack('C') + string
         doper  = pack_string_doper(operator, length)
      else
         string = ''
         doper  = pack_number_doper(operator, token)
      end

      return doper, string
   end

   ###############################################################################
   #
   # _pack_unused_doper()
   #
   # Pack an empty Doper structure.
   #
   def pack_unused_doper
      return [0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0].pack('C10')
   end

   ###############################################################################
   #
   # _pack_blanks_doper()
   #
   # Pack an Blanks/NonBlanks Doper structure.
   #
   def pack_blanks_doper(operator, token)
      if token == 'blanks'
         type     = 0x0C
         operator = 2
      else
         type     = 0x0E
         operator = 5
      end

      doper = [type,       # Data type
         operator,
         0x0000, 0x0000     # Reserved
      ].pack('CCVV')
      return doper
   end

   ###############################################################################
   #
   # _pack_string_doper()
   #
   # Pack an string Doper structure.
   #
   def pack_string_doper(operator, length)
      doper = [0x06,     # Data type
         operator,
         0x0000,         #Reserved
         length,         # String char length
         0x0, 0x0, 0x0   # Reserved
      ].pack('CCVCCCC')
      return doper
   end

   ###############################################################################
   #
   # _pack_number_doper()
   #
   # Pack an IEEE double number Doper structure.
   #
   def pack_number_doper(operator, number)
      number = [number].pack('d')
      number.reverse! if @byte_order != 0

      doper  = [0x04, operator].pack('CC') + number
      return doper
   end

   #
   # Methods related to comments and MSO objects.
   #
   
   
   ###############################################################################
   #
   # _prepare_images()
   #
   # Turn the HoH that stores the images into an array for easier handling.
   #
   def prepare_images
       count  = 0
       images = []
   
       # We sort the images by row and column but that isn't strictly required.
       #
       rows = @images.keys.sort
   
       rows.each do |row|
           cols = @images[row].keys.sort
           cols.each do |col|
               images.push(@images[row][col])
               count = count + 1
           end
       end
   
       @images       = {}
       @images_array = @images
   
       return count
   end
   
   ###############################################################################
   #
   # _prepare_comments()
   #
   # Turn the HoH that stores the comments into an array for easier handling.
   #
   def prepare_comments
       count   = 0
       comments = []
       
       # We sort the comments by row and column but that isn't strictly required.
       #
       rows = @comments.keys.sort
       rows.each do |row|
           cols = @comments[row].keys.sort
           cols.each do |col|
               comments.push(@comments[row][col])
               count = count + 1
           end
       end
   
       @comments       = {}
       @comments_array = @comments
   
       return count
   end
   
   ###############################################################################
   #
   # _prepare_charts()
   #
   # Turn the HoH that stores the charts into an array for easier handling.
   #
   def prepare_charts
       count  = 0
       charts = []
   
       # We sort the charts by row and column but that isn't strictly required.
       #
       rows = @charts.keys.sort
       rows.each do |row|
           cols = @charts[row].keys.sort
           cols.each do |col|
               charts.push(@charts[row][col])
               count = count + 1
           end
       end
   
       @charts       = {}
       @charts_array = @charts
   end

   ###############################################################################
   #
   # _store_images()
   #
   # Store the collections of records that make up images.
   #
   def store_images
       record          = 0x00EC           # Record identifier
       length          = 0x0000           # Bytes to follow
   
       ids             = @object_ids
       spid            = ids.shift
   
       images          = @images_array
       num_images      = images.size
   
       num_filters     = @filter_count
       num_comments    = @comments_array.size
       num_charts      = @charts_array.size
   
       # Skip this if there aren't any images.
       return if num_images == 0
   
       (0 .. num_images-1).each do |i|
           row         =   images[i][0]
           col         =   images[i][1]
           name        =   images[i][2]
           x_offset    =   images[i][3]
           y_offset    =   images[i][4]
           scale_x     =   images[i][5]
           scale_y     =   images[i][6]
           image_id    =   images[i][7]
           type        =   images[i][8]
           width       =   images[i][9]
           height      =   images[i][10]
   
           width  = widhth  * scale_x unless scale_x == 0
           height = height  * scale_y unless scale_y == 0
   
           # Calculate the positions of image object.
           vertices = position_object(col,row,x_offset,y_offset,width,height)
   
           if (i == 0)
               # Write the parent MSODRAWIING record.
               dg_length   = 156 + 84*(num_images -1)
               spgr_length = 132 + 84*(num_images -1)
   
               dg_length   = dg_length   + 120 * num_charts
               spgr_length = spgr_length + 120 * num_charts
   
               dg_length   = dg_length   +  96 * num_filters
               spgr_length = spgr_length +  96 * num_filters
   
               dg_length   = dg_length   + 128 * num_comments
               spgr_length = spgr_length + 128 * num_comments
   
               data = store_mso_dg_container(dg_length) +
                  store_mso_dg(ids)                     +
                  store_mso_spgr_container(spgr_length) +
                  store_mso_sp_container(40)            +
                  store_mso_spgr()                      +
                  store_mso_sp(0x0, spid, 0x0005)
               spid = spid + 1
               data = data                              + 
                  store_mso_sp_container(76)            +
                  store_mso_sp(75, spid, 0x0A00)
               spid = spid + 1
               data = data                              +
                  store_mso_opt_image(image_id)         +
                  store_mso_client_anchor(2, vertices)  +
                  store_mso_client_data()
           else
               # Write the child MSODRAWIING record.
               data = store_mso_sp_container(76)        +
                  store_mso_sp(75, spid, 0x0A00)
               spid = spid + 1
               data = data                              +
                  store_mso_opt_image(image_id)         +
                  store_mso_client_anchor(2, vertices)  +
                  store_mso_client_data()
           end
           length      = data.length
           header      = [record, length].pack("vv")
           append(header, data)
   
           store_obj_image(i+1)
       end
   
       @object_ids[0] = spid
   end

   ###############################################################################
   #
   # _store_chart_binary
   #
   # Add a binary chart object extracted from an Excel file.
   #
   def store_chart_binary(filename)
       filehandle = File.open(filename, "rb")
#                        die "Couldn't open $filename in add_chart_ext(): $!.\n";
   
       while tmp = filehandle.read(4096)
           append(tmp)
       end
   end

   ###############################################################################
   #
   # _store_filters()
   #
   # Store the collections of records that make up filters.
   #
   def store_filters
       record          = 0x00EC           # Record identifier
       length          = 0x0000           # Bytes to follow
   
       ids             = @object_ids
       spid            = ids.shift
   
       filter_area     = @filter_area
       num_filters     = @filter_count
   
       num_comments    = @comments_array.size
   
       # Number of objects written so far.
       num_objects     = @images_array.size + @charts_array.size
   
       # Skip this if there aren't any filters.
       return if num_filters == 0
   
       row1, row2, col1, col2 = @filter_area
   
       (0 .. num_filters-1).each do |i|
           vertices = [ col1 +i,    0, row1   , 0,
                        col1 +i +1, 0, row1 +1, 0]
   
           if i == 0 and !num_objects.nil?
               # Write the parent MSODRAWIING record.
               dg_length   = 168 + 96*(num_filters -1)
               spgr_length = 144 + 96*(num_filters -1)
   
               dg_length   = dg_length   + 128 *num_comments
               spgr_length = spgr_length + 128 *num_comments
   
               data = store_mso_dg_container(dg_length)           +
                  store_mso_dg(ids)                               +
                  store_mso_spgr_container(spgr_length)           +
                  store_mso_sp_container(40)                      +
                  store_mso_spgr()                                +
                  store_mso_sp(0x0, spid, 0x0005)
               spid = spid + 1
               data = data + store_mso_sp_container(88)           +
                  store_mso_sp(201, spid, 0x0A00)                 +
                  store_mso_opt_filter()                          +
                  store_mso_client_anchor(1, vertices)            +
                  store_mso_client_data()
               spid = spid + 1
   
           else
               # Write the child MSODRAWIING record.
               data = store_mso_sp_container(88)                  +
                  store_mso_sp(201, spid, 0x0A00)                 +
                  store_mso_opt_filter()                          +
                  store_mso_client_anchor(1, vertices)            +
                  store_mso_client_data()
               spid = spid + 1
           end
           length      = data.length
           header      = [record, length].pack("vv")
           append(header, data)
   
           store_obj_filter(num_objects+i+1, col1 +i)
       end
   
       # Simulate the EXTERNSHEET link between the filter and data using a formula
       # such as '=Sheet1!A1'.
       # TODO. Won't work for external data refs. Also should use a more direct
       #       method.
       #
       formula = "=#{@name}!A1"
       store_formula(formula)
   
       @object_ids[0] = spid
   end
   
   ###############################################################################
   #
   # _store_comments()
   #
   # Store the collections of records that make up cell comments.
   #
   # NOTE: We write the comment objects last since that makes it a little easier
   # to write the NOTE records directly after the MSODRAWIING records.
   #
   def store_comments
       record          = 0x00EC           # Record identifier
       length          = 0x0000           # Bytes to follow
   
       ids             = @object_ids
       spid            = ids.shift
   
       comments        = @comments_array
       num_comments    = @comments.size
   
       # Number of objects written so far.
       num_objects     = @images_array.size + @filter_count + @charts_array.size
   
       # Skip this if there aren't any comments.
       return if num_comments == 0
   
       (0 .. num_comments-1).each do |i|
           row         = comments[i][0]
           col         = comments[i][1]
           str         = comments[i][2]
           encoding    = comments[i][3]
           visible     = comments[i][6]
           color       = comments[i][7]
           vertices    = comments[i][8]
           str_len     = str.length
           str_len     = str_len / 2 if encoding != 0 # Num of chars not bytes.
           formats     = [[0, 5], [str_len, 0]]
   
           if i == 0 and num_objects != 0
               # Write the parent MSODRAWIING record.
               dg_length   = 200 + 128*(num_comments -1)
               spgr_length = 176 + 128*(num_comments -1)
   
               data = store_mso_dg_container(dg_length)          +
                  store_mso_dg(ids)                              +
                  store_mso_spgr_container(spgr_length)          +
                  store_mso_sp_container(40)                     +
                  store_mso_spgr()                               +
                  store_mso_sp(0x0, spid, 0x0005)
               spid = spid + 1
               data = data + store_mso_sp_container(120)         +
                  store_mso_sp(202, spid, 0x0A00)                +
                  store_mso_opt_comment(0x80, visible, color)    +
                  store_mso_client_anchor(3, vertices)           +
                  store_mso_client_data()
               spid = spid + 1
   
           else
               # Write the child MSODRAWIING record.
               data = store_mso_sp_container(120)                +
                  store_mso_sp(202, spid, 0x0A00)                +
                  store_mso_opt_comment(0x80, visible, color)    +
                  store_mso_client_anchor(3, vertices)           +
                  store_mso_client_data()
               spid = spid + 1
           end
           length      = data.length
           header      = [record, length].pack("vv")
           append(eader, data)
   
           store_obj_comment(num_objects + i + 1)
           store_mso_drawing_text_box()
           store_txo(str_len)
           store_txo_continue_1(str, encoding)
           store_txo_continue_2(formats)
       end
   
       # Write the NOTE records after MSODRAWIING records.
       (0 .. num_comments-1).each do |i|
           row         = comments[i][0]
           col         = comments[i][1]
           author      = comments[i][4]
           author_enc  = comments[i][5]
           visible     = comments[i][6]
   
           store_note(row, col, num_objects + i + 1,
                              author, author_enc, visible)
       end
   end
   
   ###############################################################################
   #
   # _store_mso_dg_container()
   #
   # Write the Escher DgContainer record that is part of MSODRAWING.
   #
   def store_mso_dg_container(length)
       type        = 0xF002
       version     = 15
       instance    = 0
       data        = ''
       return add_mso_generic(type, version, instance, data, length)
   end
   
   ###############################################################################
   #
   # _store_mso_dg()
   #
   # Write the Escher Dg record that is part of MSODRAWING.
   #
   def store_mso_dg(instance, num_shapes, max_spid)
       type        = 0xF008
       version     = 0
       data        = ''
       length      = 8
       data        = [num_shapes, max_spid].pack("VV")
   
       return add_mso_generic(type, version, instance, data, length)
   end
   
   ###############################################################################
   #
   # _store_mso_spgr_container()
   #
   # Write the Escher SpgrContainer record that is part of MSODRAWING.
   #
   def store_mso_spgr_container(length)
       type        = 0xF003
       version     = 15
       instance    = 0
       data        = ''
   
       return add_mso_generic(type, version, instance, data, length)
   end
   
   ###############################################################################
   #
   # _store_mso_sp_container()
   #
   # Write the Escher SpContainer record that is part of MSODRAWING.
   #
   def store_mso_sp_container(length)
       type        = 0xF004
       version     = 15
       instance    = 0
       data        = ''
   
       return add_mso_generic(type, version, instance, data, length)
   end
   
   ###############################################################################
   #
   # _store_mso_spgr()
   #
   # Write the Escher Spgr record that is part of MSODRAWING.
   #
   def store_mso_spgr
       type        = 0xF009
       version     = 1
       instance    = 0
       data        = [0, 0, 0, 0].pack("VVVV")
       length      = 16
   
       return add_mso_generic(type, version, instance, data, length)
   end
   
   ###############################################################################
   #
   # _store_mso_sp()
   #
   # Write the Escher Sp record that is part of MSODRAWING.
   #
   def store_mso_sp(instance, spid, options)
       type        = 0xF00A
       version     = 2
       data        = ''
       length      = 8
       data        = [spid, options].pack('VV')

       return add_mso_generic(type, version, instance, data, length)
   end
   
   ###############################################################################
   #
   # _store_mso_opt_comment()
   #
   # Write the Escher Opt record that is part of MSODRAWING.
   #
   def store_mso_opt_comment(spid, visible, colour = 0x50)
       type        = 0xF00B
       version     = 3
       instance    = 9
       data        = ''
       length      = 54
   
       # Use the visible flag if set by the user or else use the worksheet value.
       # Note that the value used is the opposite of _store_note().
       #
       unless visible.nil?
          visible = visible           ? 0x0000 : 0x0002
       else
          visible = @comments_visible ? 0x0000 : 0x0002
       end
   
       data = [spid].pack('V')                            +
          ['0000BF00080008005801000000008101'].pack("H*") +
          [colour].pack("C")                              +
          ['000008830150000008BF011000110001'+'02000000003F0203000300BF03'].pack("H*")  +
          [visible].pack('c')                             +
          ['0A00'].pack('H*')
   
       return add_mso_generic(type, version, instance, data, length)
   end
   
   ###############################################################################
   #
   # _store_mso_opt_image()
   #
   # Write the Escher Opt record that is part of MSODRAWING.
   #
   def store_mso_opt_image(spid)
       type        = 0xF00B
       version     = 3
       instance    = 3
       data        = ''
       length      = nil
   
       data = [0x4104].pack('v') +
         [spid].pack('V')        +
         [0x01BF].pack('v')      +
         [0x00010000].pack('V')  +
         [0x03BF].pack( 'v')     +
         [0x00080000].pack( 'V')
   
       return add_mso_generic(type, version, instance, data, length)
   end
   
   ###############################################################################
   #
   # _store_mso_opt_chart()
   #
   # Write the Escher Opt record that is part of MSODRAWING.
   #
   def store_mso_opt_chart
       type        = 0xF00B
       version     = 3
       instance    = 9
       data        = ''
       length      = nil
   
       data = [0x007F].pack('v')       +        # Protection -> fLockAgainstGrouping
          [0x01040104].pack('V')       +
          [0x00BF].pack('v')           +        # Text -> fFitTextToShape
          [0x00080008].pack('V')       +
          [0x0181].pack('v')           +        # Fill Style -> fillColor
          [0x0800004E].pack('V')       +
          [0x0183].pack('v')           +        # Fill Style -> fillBackColor
          [0x0800004D].pack('V')       +
   
          [0x01BF].pack('v')           +         # Fill Style -> fNoFillHitTest
          [0x00110010].pack('V')       +
          [0x01C0].pack('v')           +        # Line Style -> lineColor
          [0x0800004D].pack('V')       +
          [0x01FF].pack('v')           +        # Line Style -> fNoLineDrawDash
          [0x00080008].pack('V')       +
          [0x023F].pack('v')            +        # Shadow Style -> fshadowObscured
          [0x00020000].pack('V')       +
          [0x03BF].pack('v')           +        # Group Shape -> fPrint
          [0x00080000].pack('V')
   
       return add_mso_generic(type, version, instance, data, length)
   end
   
   ###############################################################################
   #
   # _store_mso_opt_filter()
   #
   # Write the Escher Opt record that is part of MSODRAWING.
   #
   def store_mso_opt_filter
       type        = 0xF00B
       version     = 3
       instance    = 5
       data        = ''
       length      = nil
   
       data = [0x007F].pack('v')     +    # Protection -> fLockAgainstGrouping
          [0x01040104].pack('V')     +
          [0x00BF].pack('v')    +        # Text -> fFitTextToShape
          [0x00080008].pack('V')+
          [0x01BF].pack('v')    +        # Fill Style -> fNoFillHitTest
          [0x00010000].pack('V')+
          [0x01FF].pack('v')    +        # Line Style -> fNoLineDrawDash
          [0x00080000].pack('V')+
          [0x03BF].pack('v')    +        # Group Shape -> fPrint
          [0x000A0000].pack('V')
   
       return add_mso_generic(type, version, instance, data, length)
   end
   
   ###############################################################################
   #
   # _store_mso_client_anchor()
   #    my flag         = shift;
   #    my $col_start   = $_[0];    # Col containing upper left corner of object
   #    my $x1          = $_[1];    # Distance to left side of object
   #
   #    my $row_start   = $_[2];    # Row containing top left corner of object
   #    my $y1          = $_[3];    # Distance to top of object
   #
   #    my $col_end     = $_[4];    # Col containing lower right corner of object
   #    my $x2          = $_[5];    # Distance to right side of object
   #
   #    my $row_end     = $_[6];    # Row containing bottom right corner of object
   #    my $y2          = $_[7];    # Distance to bottom of object
   #
   # Write the Escher ClientAnchor record that is part of MSODRAWING.
   #
   def store_mso_client_anchor(flag, col_start, x1, row_start, y1, col_end, x2, row_end, y2)
       type        = 0xF010
       version     = 0
       instance    = 0
       data        = ''
       length      = 18
   
       data = [flag, col_start, x1, row_start, y1, col_end, x2, row_end, y2].pack( "v9")
   
       return add_mso_generic(type, version, instance, data, length)
   end
   
   ###############################################################################
   #
   # _store_mso_client_data()
   #
   # Write the Escher ClientData record that is part of MSODRAWING.
   #
   def store_mso_client_data
       type        = 0xF011
       version     = 0
       instance    = 0
       data        = ''
       length      = 0
   
       return add_mso_generic(type, version, instance, data, length)
   end
   
   ###############################################################################
   #
   # _store_obj_comment()
   #    my $obj_id      = $_[0];    # Object ID number.
   #
   # Write the OBJ record that is part of cell comments.
   #
   def store_obj_comment(obj_id)
       record      = 0x005D   # Record identifier
       length      = 0x0034   # Bytes to follow
   
       obj_type    = 0x0019   # Object type (comment).
       data        = ''       # Record data.
   
       sub_record  = 0x0000   # Sub-record identifier.
       sub_length  = 0x0000   # Length of sub-record.
       sub_data    = ''       # Data of sub-record.
       options     = 0x4011
       reserved    = 0x0000
   
       # Add ftCmo (common object data) subobject
       sub_record     = 0x0015   # ftCmo
       sub_length     = 0x0012
       sub_data       = [obj_type, obj_id, options, reserved, reserved, reserved].pack( "vvvVVV")
       data           = [sub_record, sub_length].pack("vv") + sub_data
   
       # Add ftNts (note structure) subobject
       sub_record  = 0x000D   # ftNts
       sub_length  = 0x0016
       sub_data    = [reserved,reserved,reserved,reserved,reserved,reserved].pack( "VVVVVv")
       data        = [sub_record, sub_length].pack("vv") + sub_data
   
       # Add ftEnd (end of object) subobject
       sub_record  = 0x0000   # ftNts
       sub_length  = 0x0000
       data        = data + [sub_record, sub_length].pack("vv")
   
       # Pack the record.
       header      = [record, length].pack("vv")
   
       append(header, data)
   
   end
   
   ###############################################################################
   #
   # _store_obj_image()
   #    my $obj_id      = $_[0];    # Object ID number.
   #
   # Write the OBJ record that is part of image records.
   #
   def store_obj_image(obj_id)
       record      = 0x005D   # Record identifier
       length      = 0x0026   # Bytes to follow
   
       obj_type    = 0x0008   # Object type (Picture).
       data        = ''       # Record data.
   
       sub_record  = 0x0000   # Sub-record identifier.
       sub_length  = 0x0000   # Length of sub-record.
       sub_data    = ''       # Data of sub-record.
       options     = 0x6011
       reserved    = 0x0000
   
       # Add ftCmo (common object data) subobject
       sub_record  = 0x0015   # ftCmo
       sub_length  = 0x0012
       sub_data    = [obj_type, obj_id, options, reserved, reserved, reserved].pack('vvvVVV')
       data        = [sub_record, sub_length].pack('vv') + sub_data
   
       # Add ftCf (Clipboard format) subobject
       sub_record  = 0x0007   # ftCf
       sub_length  = 0x0002
       sub_data    = [0xFFFF].pack( 'v')
       data        = data + [sub_record, sub_length].pack('vv') + sub_data
   
       # Add ftPioGrbit (Picture option flags) subobject
       sub_record  = 0x0008   # ftPioGrbit
       sub_length  = 0x0002
       sub_data    = [0x0001].pack('v')
       data        = data + [sub_record, sub_length].pack('vv') + sub_data
   
       # Add ftEnd (end of object) subobject
       sub_record  = 0x0000   # ftNts
       sub_length  = 0x0000
       data        = data + [sub_record, sub_length].pack('vv') + sub_data

       # Pack the record.
       header  = [record, length].pack('vv')
   
       append(header, data)
   
   end
   
   
   ###############################################################################
   #
   # _store_obj_chart()
   #    my $obj_id      = $_[0];    # Object ID number.
   #
   # Write the OBJ record that is part of chart records.
   #
   def store_obj_chart(obj_id)
       record      = 0x005D   # Record identifier
       length      = 0x001A   # Bytes to follow
   
       obj_type    = 0x0005   # Object type (chart).
       data        = ''       # Record data.
   
       sub_record  = 0x0000   # Sub-record identifier.
       sub_length  = 0x0000   # Length of sub-record.
       sub_data    = ''       # Data of sub-record.
       options     = 0x6011
       reserved    = 0x0000
   
       # Add ftCmo (common object data) subobject
       sub_record  = 0x0015   # ftCmo
       sub_length  = 0x0012
       sub_data    = [obj_type, obj_id, options, reserved, reserved, reserved].pack('vvvVVV')
       data        = [sub_record, sub_length].pack('vv') + sub_data
   
       # Add ftEnd (end of object) subobject
       sub_record  = 0x0000   # ftNts
       sub_length  = 0x0000
       data        = data + [sub_record, sub_length].pack('vv')
   
       # Pack the record.
       header  = [record, length].pack('vv')
   
       append(header, data)
   
   end
   
   ###############################################################################
   #
   # _store_obj_filter()
   #    my $obj_id      = $_[0];    # Object ID number.
   #    my $col         = $_[1];
   #
   # Write the OBJ record that is part of filter records.
   #
   def store_obj_filter(obj_id, col)
       record      = 0x005D   # Record identifier
       length      = 0x0046   # Bytes to follow
   
       obj_type    = 0x0014   # Object type (combo box).
       data        = ''       # Record data.
   
       sub_record  = 0x0000   # Sub-record identifier.
       sub_length  = 0x0000   # Length of sub-record.
       sub_data    = ''       # Data of sub-record.
       options     = 0x2101
       reserved    = 0x0000
   
       # Add ftCmo (common object data) subobject
       sub_record  = 0x0015   # ftCmo
       sub_length  = 0x0012
       sub_data    = [obj_type, obj_id, options, reserved, reserved, reserved].pack('vvvVVV')
       data        = [sub_record, sub_length].pack('vv') + sub_data
   
       # Add ftSbs Scroll bar subobject
       sub_record  = 0x000C   # ftSbs
       sub_length  = 0x0014
       sub_data    = ['0000000000000000640001000A00000010000100'].pack('H*')
       data        = data + [sub_record, sub_length].pack('vv') + sub_data
   
       # Add ftLbsData (List box data) subobject
       sub_record  = 0x0013   # ftLbsData
       sub_length  = 0x1FEE   # Special case (undocumented).
   
       # If the filter is active we set one of the undocumented flags.
   
       if @filter_cols[col]
           sub_data       = ['000000000100010300000A0008005700'].pack('H*')
       else
           sub_data       = ['00000000010001030000020008005700'].pack('H*')
       end
   
       data        = data + [sub_record, sub_length].pack('vv') + sub_data
   
       # Add ftEnd (end of object) subobject
       sub_record  = 0x0000   # ftNts
       sub_length  = 0x0000
       data        = data + [sub_record, sub_length].pack('vv')
   
       # Pack the record.
       header  = [record, length].pack('vv')
   
       append(header, data)
   end
   
   ###############################################################################
   #
   # _store_mso_drawing_text_box()
   #
   # Write the MSODRAWING ClientTextbox record that is part of comments.
   #
   def store_mso_drawing_text_box
       record      = 0x00EC           # Record identifier
       length      = 0x0008           # Bytes to follow
   
       data        = store_mso_client_text_box()
       header  = [record, length].pack('vv')
   
       append(header, data)
   end
   
   ###############################################################################
   #
   # _store_mso_client_text_box()
   #
   # Write the Escher ClientTextbox record that is part of MSODRAWING.
   #
   def store_mso_client_text_box
       type        = 0xF00D
       version     = 0
       instance    = 0
       data        = ''
       length      = 0
   
       return add_mso_generic(type, version, instance, data, length)
   end
   
   ###############################################################################
   #
   # _store_txo()
   #    my $string_len  = $_[0];                # Length of the note text.
   #    my $format_len  = $_[1] || 16;          # Length of the format runs.
   #    my $rotation    = $_[2] || 0;           # Options
   #
   # Write the worksheet TXO record that is part of cell comments.
   #
   def store_txo(string_len, format_len = 16, rotation = 0)
       record      = 0x01B6               # Record identifier
       length      = 0x0012               # Bytes to follow
   
       grbit       = 0x0212               # Options
       reserved    = 0x0000               # Options
   
       # Pack the record.
       header  = [record, length].pack('vv')
       data    = [grbit, rotation, reserved, reserved,
                  string_len, format_len, reserved].pack("vvVvvvV")
   
       append(header, data)
   end
   
   ###############################################################################
   #
   # _store_txo_continue_1()
   #    my $string      = $_[0];                # Comment string.
   #    my $encoding    = $_[1] || 0;           # Encoding of the string.
   #
   # Write the first CONTINUE record to follow the TXO record. It contains the
   # text data.
   #
   def store_txo_continue_1(string, encoding = 0)
       record      = 0x003C               # Record identifier
   
       # Split long comment strings into smaller continue blocks if necessary.
       # We can't let BIFFwriter::_add_continue() handled this since an extra
       # encoding byte has to be added similar to the SST block.
       #
       # We make the limit size smaller than the _add_continue() size and even
       # so that UTF16 chars occur in the same block.
       #
       limit = 8218
       while string.length > limit
           string[0 .. limit] = ""
           tmp_str = string
           data    = [encoding].pack("C") + tmp_str
           length  = data.length
           header  = [record, length].pack('vv')
   
           append(header, data)
       end
   
       # Pack the record.
       data    = [encoding].pack("C") + string
       length  = data.length
       header  = [record, length].pack('vv')
   
       append(header, data)
   end
   
   ###############################################################################
   #
   # _store_txo_continue_2()
   #    my $formats     = $_[0];                # Formatting information
   #
   # Write the second CONTINUE record to follow the TXO record. It contains the
   # formatting information for the string.
   #
   def store_txo_continue_2(formats)
       record      = 0x003C               # Record identifier
       length      = 0x0000               # Bytes to follow
   
       # Pack the record.
       data = ''
   
       formats.each do |a_ref|
           data = data + [a_ref[0], a_ref[1], 0x0].pack('vvV')
       end
   
       length  = data.length
       header  = [record, length].pack("vv")
   
       append(header, data)
   end
   
   ###############################################################################
   #
   # _store_note()
   #    my $row         = $_[0];
   #    my $col         = $_[1];
   #    my $obj_id      = $_[2];
   #    my $author      = $_[3] || $self->{_comments_author};
   #    my $author_enc  = $_[4] || $self->{_comments_author_enc};
   #    my $visible     = $_[5];
   #
   # Write the worksheet NOTE record that is part of cell comments.
   #
   def store_note(row, col, obj_id, author = nil, author_enc = nil, visible = nil)
       record      = 0x001C               # Record identifier
       length      = 0x000C               # Bytes to follow

       author     = @comments_author     if author.nil?
       author_enc = @comments_author_enc if author_enc.nil?
   
       # Use the visible flag if set by the user or else use the worksheet value.
       # The flag is also set in _store_mso_opt_comment() but with the opposite
       # value.
       unless visible.nil?
           visible = visible != 0      ? 0x0002 : 0x0000
       else
           visible = @comments_visible ? 0x0002 : 0x0000
       end
   
       # Get the number of chars in the author string (not bytes).
       num_chars  = author.length
       num_chars  = num_chars / 2 if author_enc != 0 && !author_enc.nil?
   
       # Null terminate the author string.
       author = author + "\0"
   
   
       # Pack the record.
       data    = [row, col, visible, obj_id, num_chars, author_enc].pack("vvvvvC")
   
       length  = data.length + author.length
       header  = [record, length].pack("vv")
   
       append(header, data)
   end

   ###############################################################################
   #
   # _comment_params()
   #
   # This method handles the additional optional parameters to write_comment() as
   # well as calculating the comment object position and vertices.
   #
   def comment_params(*args)
       row    = args.shift
       col    = args.shift
       string = args.shift

       params  = {
                       :author          => '',
                       :author_encoding => 0,
                       :encoding        => 0,
                       :color           => nil,
                       :start_cell      => nil,
                       :start_col       => nil,
                       :start_row       => nil,
                       :visible         => nil,
                       :width           => 129,
                       :height          => 75,
                       :x_offset        => nil,
                       :x_scale         => 1,
                       :y_offset        => nil,
                       :y_scale         => 1
                }
   
       # Overwrite the defaults with any user supplied values. Incorrect or
       # misspelled parameters are silently ignored.

       ###       params   = (%params, @_);  converted like this.  right?
       ary  = Array.new(args)
       alist = []
       while ary.size > 0
         alist << [ary[0], ary[1]]
         ary.shift
         ary.shift
       end
       pary = params.to_a
       alist.each { |a| pary << a }
       params = Hash[*pary.flatten]

       # Ensure that a width and height have been set.
       params[:width]  = 129 if not params[:width]
       params[:height] = 75  if not params[:height]
   
       # Check that utf16 strings have an even number of bytes.
       if params[:encoding] != 0
           raise "Uneven number of bytes in comment string" if string.length % 2
   
           # Change from UTF-16BE to UTF-16LE
           string = [string].unpack('n*').pack('v*')
       end
   
       if params[:author_encoding] != 0
           raise "Uneven number of bytes in author string"  if params[:author] % 2
   
           # Change from UTF-16BE to UTF-16LE
           params[:author] = [params[:author]].unpack('n*').pack('v*')
       end
   
       # Limit the string to the max number of chars (not bytes).
       max_len = 32767
       max_len = max_len * 2 if params[:encoding] != 0
   
       if string.length > max_len
           string = string[0 .. max_len]
       end
   
       # Set the comment background colour.
       color = params[:color]
       color = Format._get_color(color)
       color = 0x50 if color == 0x7FFF  # Default color.
       params[:color] = color
   
       # Convert a cell reference to a row and column.
       unless params[:start_cell].nil?
           row, col = substitute_cellref(params[:start_cell])
           params[:start_row] = row
           params[:start_col] = col
       end
   
       # Set the default start cell and offsets for the comment. These are
       # generally fixed in relation to the parent cell. However there are
       # some edge cases for cells at the, er, edges.
       #
       if params[:start_row].nil?
           case row
              when 0     then params[:start_row] = 0
              when 65533 then params[:start_row] = 65529
              when 65534 then params[:start_row] = 65530
              when 65535 then params[:start_row] = 65531
              else            params[:start_row] = row -1
           end
       end
   
       if params[:y_offset].nil?
           case row
              when 0     then params[:y_offset]  = 2
              when 65533 then params[:y_offset]  = 4
              when 65534 then params[:y_offset]  = 4
              when 65535 then params[:y_offset]  = 2
              else            params[:y_offset]  = 7
           end
       end
   
       if params[:start_col].nil?
           case col
              when 253   then params[:start_col] = 250
              when 254   then params[:start_col] = 251
              when 255   then params[:start_col] = 252
              else            params[:start_col] = col + 1
           end
       end
   
       if params[:x_offset].nil?
           case col
              when 253   then params[:x_offset] = 49
              when 254   then params[:x_offset] = 49
              when 255   then params[:x_offset] = 49
              else            params[:x_offset] = 15
           end
       end
   
       # Scale the size of the comment box if required. We scale the width and
       # height using the relationship d2 =(d1 -1)*s +1, where d is dimension
       # and s is scale. This gives values that match Excel's behaviour.
       #
       if params[:x_scale] != 0
           params[:width]  = ((params[:width]  -1) * params[:x_scale]) +1
       end
   
       if params[:y_scale] != 0
           params[:height] = ((params[:height] -1) * params[:y_scale]) +1
       end
   
       # Calculate the positions of comment object.
       vertices = position_object( params[:start_col],
                                   params[:start_row],
                                   params[:x_offset],
                                   params[:y_offset],
                                   params[:width],
                                   params[:height]
                                 )
   
       return [row, col, string,
               params[:encoding],
               params[:author],
               params[:author_encoding],
               params[:visible],
               params[:color],
               @vertices
             ]
   end

   #
   # DATA VALIDATION
   #
   
   ###############################################################################
   #
   # data_validation($row, $col, {...})
   #
   # This method handles the interface to Excel data validation.
   # Somewhat ironically the this requires a lot of validation code since the
   # interface is flexible and covers a several types of data validation.
   #
   # We allow data validation to be called on one cell or a range of cells. The
   # hashref contains the validation parameters and must be the last param:
   #    data_validation($row, $col, {...})
   #    data_validation($first_row, $first_col, $last_row, $last_col, {...})
   #
   # Returns  0 : normal termination
   #         -1 : insufficient number of arguments
   #         -2 : row or column out of range
   #         -3 : incorrect parameter.
   #
   def data_validation(*args)
       # Check for a cell reference in A1 notation and substitute row and column
       if args[0] =~ /^\D/
           args = substitute_cellref(*args)
       end
   
       # Check for a valid number of args.
       return -1 if args.size != 5 && args.size != 3
   
       # The final hashref contains the validation parameters.
       param = args.pop
   
       # Make the last row/col the same as the first if not defined.
       row1, col1, row2, col2 = args
       if row2.nil?
           row2 = row1
           col2 = col1
       end
   
       # Check that row and col are valid without storing the values.
       return -2 if check_dimensions(row1, col1, 1, 1) != 0
       return -2 if check_dimensions(row2, col2, 1, 1) != 0
   
       # Check that the last parameter is a hash list.
       unless param.kind_of?(Hash)
#           carp "Last parameter '$param' in data_validation() must be a hash ref";
           return -3
       end
   
       # List of valid input parameters.
       valid_parameter = {
                                 :validate          => 1,
                                 :criteria          => 1,
                                 :value             => 1,
                                 :source            => 1,
                                 :minimum           => 1,
                                 :maximum           => 1,
                                 :ignore_blank      => 1,
                                 :dropdown          => 1,
                                 :show_input        => 1,
                                 :input_title       => 1,
                                 :input_message     => 1,
                                 :show_error        => 1,
                                 :error_title       => 1,
                                 :error_message     => 1,
                                 :error_type        => 1,
                                 :other_cells       => 1
                          }
   
       # Check for valid input parameters.
       param.each_key do |param_key|
           if valid_parameter[param_key].nil?
#               carp "Unknown parameter '$param_key' in data_validation()";
               return -3
           end
       end
   
       # Map alternative parameter names 'source' or 'minimum' to 'value'.
       param[:value] = param[:source]  unless param[:source].nil?
       param[:value] = param[:minimum] unless param[:minimum].nil?
   
       # 'validate' is a required paramter.
       if param[:validate].nil?
#           carp "Parameter 'validate' is required in data_validation()";
           return -3
       end
   
       # List of  valid validation types.
       valid_type = {
                                 'any'             => 0,
                                 'any value'       => 0,
                                 'whole number'    => 1,
                                 'whole'           => 1,
                                 'integer'         => 1,
                                 'decimal'         => 2,
                                 'list'            => 3,
                                 'date'            => 4,
                                 'time'            => 5,
                                 'text length'     => 6,
                                 'length'          => 6,
                                 'custom'          => 7
                     }
   
       # Check for valid validation types.
       if valid_type[param[:validate].downcase].nil?
#           carp "Unknown validation type '$param->{validate}' for parameter " .
#                "'validate' in data_validation()";
           return -3
       else
           param[:validate] = valid_type[param[:validate].downcase]
       end
   
       # No action is requied for validation type 'any'.
       # TODO: we should perhaps store 'any' for message only validations.
       return 0 if param[:validate] == 0
   
       # The list and custom validations don't have a criteria so we use a default
       # of 'between'.
       if param[:validate] == 3 || param[:validate] == 7
           param[:criteria]  = 'between'
           param[:maximum]   = nil
       end
   
       # 'criteria' is a required parameter.
       if param[:criteria].nil?
#           carp "Parameter 'criteria' is required in data_validation()";
           return -3
       end
   
       # List of valid criteria types.
       criteria_type = {
                                 'between'                     => 0,
                                 'not between'                 => 1,
                                 'equal to'                    => 2,
                                 '='                           => 2,
                                 '=='                          => 2,
                                 'not equal to'                => 3,
                                 '!='                          => 3,
                                 '<>'                          => 3,
                                 'greater than'                => 4,
                                 '>'                           => 4,
                                 'less than'                   => 5,
                                 '<'                           => 5,
                                 'greater than or equal to'    => 6,
                                 '>='                          => 6,
                                 'less than or equal to'       => 7,
                                 '<='                          => 7
                       }
   
       # Check for valid criteria types.
       if criteria_type[param[:criteria].downcase].nil?
#           carp "Unknown criteria type '$param->{criteria}' for parameter " .
#                "'criteria' in data_validation()";
           return -3
       else
           param[:criteria] = criteria_type[param[:criteria].downcase]
       end
   
       # 'Between' and 'Not between' criterias require 2 values.
       if param[:criteria] == 0 || param[:criteria] == 1
           if param[:maximum].nil?
#               carp "Parameter 'maximum' is required in data_validation() " .
#                    "when using 'between' or 'not between' criteria";
               return -3
           end
       else
           param[:maximum] = nil
       end
   
       # List of valid error dialog types.
       error_type = {
                                 'stop'        => 0,
                                 'warning'     => 1,
                                 'information' => 2
                    }
   
       # Check for valid error dialog types.
       if param[:error_type].nil?
           param[:error_type] = 0
       elsif error_type[param[:error_type].downcase].nil?
#           carp "Unknown criteria type '$param->{error_type}' for parameter " .
#                "'error_type' in data_validation()";
           return -3
       else
           param[:error_type] = error_type[param[error_type].downcase]
       end
   
       # Convert date/times value sif required.
       if param[:validate] == 4 || param[:validate] == 5
           if param[:value] =~ /T/
               date_time = convert_date_time(param[:value])
               if date_time.nil?
#                   carp "Invalid date/time value '$param->{value}' " .
#                        "in data_validation()";
                   return -3
               else
                   param[:value] = date_time
               end
           end
           if !param[:maximum].nil? && param[:maximum] =~ /T/
               date_time = convert_date_time(param[:maximum])
   
               if date_time.nil?
#                   carp "Invalid date/time value '$param->{maximum}' " .
#                        "in data_validation()";
                   return -3
               else
                   param[:maximum] = date_time
               end
           end
       end
   
       # Set some defaults if they haven't been defined by the user.
       param[:ignore_blank]  = 1 if param[:ignore_blank].nil?
       param[:dropdown]      = 1 if param[:dropdown].nil?
       param[:show_input]    = 1 if param[:show_input].nil?
       param[:show_error]    = 1 if param[:show_error].nil?
   
       # These are the cells to which the validation is applied.
       param[:cells] = [[row1, col1, row2, col2]]
   
       # A (for now) undocumented parameter to pass additional cell ranges.
       if !param[:other_cells].nil?
   
           param[:cells].push(param[:other_cells])
       end
   
       # Store the validation information until we close the worksheet.
       @validations.push(param)
   end

   ###############################################################################
   #
   # _store_validation_count()
   #
   # Store the count of the DV records to follow.
   #
   # Note, this could be wrapped into _store_dv() but we may require separate
   # handling of the object id at a later stage.
   #
   def store_validation_count
       dv_count = @validations
       obj_id   = -1
   
       return unless dv_count
   
       store_dval(obj_id , dv_count)
   end
   
   ###############################################################################
   #
   # _store_validations()
   #
   # Store the data_validation records.
   #
   def store_validations
       return if @validations.size == 0
   
       @validations.each do |param|
           store_dv(           param[:cells],
                               param[:validate],
                               param[:criteria],
                               param[:value],
                               param[:maximum],
                               param[:input_title],
                               param[:input_message],
                               param[:error_title],
                               param[:error_message],
                               param[:error_type],
                               param[:ignore_blank],
                               param[:dropdown],
                               param[:show_input],
                               param[:show_error]
                    )
       end
   end
   
   ###############################################################################
   #
   # _store_dval()
   #    my $obj_id      = $_[0];        # Object ID number.
   #    my $dv_count    = $_[1];        # Count of DV structs to follow.
   #
   # Store the DV record which contains the number of and information common to
   # all DV structures.
   #
   def store_dval(obj_id, dv_count)
       record      = 0x01B2       # Record identifier
       length      = 0x0012       # Bytes to follow
   
       flags       = 0x0004       # Option flags.
       x_coord     = 0x00000000   # X coord of input box.
       y_coord     = 0x00000000   # Y coord of input box.
   
       # Pack the record.
       header = [record, length].pack('vv')
       data   = [flags, x_coord, y_coord, obj_id, dv_count].pack('vVVVV')
   
       append(header, data)
   end
   
   ###############################################################################
   #
   # _store_dv()
   #    my $cells           = $_[0];        # Aref of cells to which DV applies.
   #    my $validation_type = $_[1];        # Type of data validation.
   #    my $criteria_type   = $_[2];        # Validation criteria.
   #    my $formula_1       = $_[3];        # Value/Source/Minimum formula.
   #    my $formula_2       = $_[4];        # Maximum formula.
   #    my $input_title     = $_[5];        # Title of input message.
   #    my $input_message   = $_[6];        # Text of input message.
   #    my $error_title     = $_[7];        # Title of error message.
   #    my $error_message   = $_[8];        # Text of input message.
   #    my $error_type      = $_[9];        # Error dialog type.
   #    my $ignore_blank    = $_[10];       # Ignore blank cells.
   #    my $dropdown        = $_[11];       # Display dropdown with list.
   #    my $input_box       = $_[12];       # Display input box.
   #    my $error_box       = $_[13];       # Display error box.
   #
   # Store the DV record that specifies the data validation criteria and options
   # for a range of cells..
   #
   def store_dv(cells, validation_type, criteria_type,
                formula_1, formula_2, input_title, input_message,
                error_title, error_message, error_type,
                ignore_blank, dropdown, input_box, error_box)
       record          = 0x01BE       # Record identifier
       length          = 0x0000       # Bytes to follow
   
       flags           = 0x00000000   # DV option flags.
   
       ime_mode        = 0            # IME input mode for far east fonts.
       str_lookup      = 0            # See below.
   
       # Set the string lookup flag for 'list' validations with a string array.
       if validation_type == 3 && formula_1.kind_of?(Array)
           str_lookup = 1
       end
   
       # The dropdown flag is stored as a negated value.
       no_dropdown = !dropdown
   
       # Set the required flags.
       flags |= validation_type
       flags |= error_type       << 4
       flags |= str_lookup       << 7
       flags |= ignore_blank     << 8
       flags |= no_dropdown      << 9
       flags |= ime_mode         << 10
       flags |= input_box        << 18
       flags |= error_box        << 19
       flags |= criteria_type    << 20
   
       # Pack the validation formulas.
       formula_1 = pack_dv_formula(formula_1)
       formula_2 = pack_dv_formula(formula_2)
   
       # Pack the input and error dialog strings.
       input_title   = pack_dv_string(input_title,   32 )
       error_title   = pack_dv_string(error_title,   32 )
       input_message = pack_dv_string(input_message, 255)
       error_message = pack_dv_string(error_message, 255)
   
       # Pack the DV cell data.
       dv_count = cells.size
       dv_data  = [dv_count].pack('v')
       cells.each do |range|
           dv_data = dv_data + [range[0]+range[1]+range[2]+range[3]].pack('vvvv')
       end
   
       # Pack the record.
       data   = [flags].pack('V')            +
                input_title                  +
                error_title                  +
                input_message                +
                error_message                +
                formula_1                    +
                formula_2                    +
                dv_data
   
       header = [record, length].pack('vv')

       append(header, data)
   end

   ###############################################################################
   #
   # _pack_dv_string()
   #
   # Pack the strings used in the input and error dialog captions and messages.
   # Captions are limited to 32 characters. Messages are limited to 255 chars.
   #
   def pack_dv_string(string = nil, max_length = 0)
       str_length  = 0
       encoding    = 0
   
       # The default empty string is "\0".
       if string.nil? || string == ''
           string = "\0"
       end
   
       # Excel limits DV captions to 32 chars and messages to 255.
       if string.length > max_length
           string = string[0 .. max_length]
       end
   
       str_length = string.length
   
       return [str_length, encoding].pack('vC') + string
   end
   
   ###############################################################################
   #
   # _pack_dv_formula()
   #
   # Pack the formula used in the DV record. This is the same as an cell formula
   # with some additional header information. Note, DV formulas in Excel use
   # relative addressing (R1C1 and ptgXxxN) however we use the Formula.pm's
   # default absoulute addressing (A1 and ptgXxx).
   #
   def pack_dv_formula(formula = nil)
       encoding    = 0
       length      = 0
       unused      = 0x0000
       tokens      = []
   
       # Return a default structure for unused formulas.
       if formula.nil? || formula == ''
           return [0, unused].pack('vv')
       end
   
       # Pack a list array ref as a null separated string.
       if formula.kind_of?(Array)
           formula   = formula.join("\0")
           formula   = '"' + formula + '"'
       end
   
       # Strip the = sign at the beginning of the formula string
       formula.sub!(/^=/, '')
   
       # Parse the formula using the parser in Formula.pm
       parser  = @parser
   
       # In order to raise formula errors from the point of view of the calling
       # program we use an eval block and re-raise the error from here.
       #
       tokens = parser.parse_formula(formula)   # ????
   
#       if ($@) {
#           $@ =~ s/\n$//;  # Strip the \n used in the Formula.pm die()
#           croak $@;       # Re-raise the error
#       }
#       else {
#           # TODO test for non valid ptgs such as Sheet2!A1
#       }
   
       # Force 2d ranges to be a reference class.
       tokens.each do |t|
         t.sub!(/_range2d/, "_range2dR")
       end
   
       # Parse the tokens into a formula string.
       formula = parser.parse_tokens(tokens)
   
       return [formula.length, unused].pack('vv')
   end

end

=begin
= Differences between Worksheet.pm and worksheet.rb
---write_url
   I made this a public method to be called directly by the user if they want
   to write a url string.  A variable number of arguments made it a pain to
   integrate into the 'write' method.
=end
