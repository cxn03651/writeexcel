###############################################################################
#
# Formula - A class for generating Excel formulas.
#
#
# Used in conjunction with Spreadsheet::WriteExcel
#
# Copyright 2000-2008, John McNamara, jmcnamara@cpan.org
#
# original written in Perl by John McNamara
# converted to Ruby by Hideo Nakamura, cxn03651@msj.biglobe.ne.jp
#
require 'nkf'
require 'strscan'
require 'excelformulaparser'

class Formula < ExcelFormulaParser

  NonAscii = /[^!"#\$%&'\(\)\*\+,\-\.\/\:\;<=>\?@0-9A-Za-z_\[\\\]^` ~\0\n]/

  attr_accessor :byte_order, :workbook, :ext_sheets, :ext_refs, :ext_ref_count

  def initialize(byte_order)
    @byte_order     = byte_order
    @workbook       = ""
    @ext_sheets     = {}
    @ext_refs       = {}
    @ext_ref_count  = 0
    initialize_hashes
  end

  ###############################################################################
  #
  # parse_formula()
  #
  # Takes a textual description of a formula and returns a RPN encoded byte
  # string.
  #
  def parse_formula(formula, byte_stream = false)
    # Build the parse tree for the formula
    tokens = reverse(parse(formula))

    # Add a volatile token if the formula contains a volatile function.
    # This must be the first token in the list
    #
    tokens.unshift('_vol') if check_volatile(tokens) != 0

    # The return value depends on which Worksheet.pm method is the caller
    unless byte_stream
      # Parse formula to see if it throws any errors and then
      # return raw tokens to Worksheet::store_formula()
      #
      tokens
    else
      # Return byte stream to Worksheet::write_formula()
      parse_tokens(tokens)
    end
  end

  ###############################################################################
  #
  # parse_tokens()
  #
  # Convert each token or token pair to its Excel 'ptg' equivalent.
  #
  def parse_tokens(tokens)
    parse_str   = ''
    last_type   = ''
    modifier    = ''
    num_args    = 0
    _class      = 0
    _classary   = [1]
    args        = tokens.dup
    # A note about the class modifiers used below. In general the class,
    # "reference" or "value", of a function is applied to all of its operands.
    # However, in certain circumstances the operands can have mixed classes,
    # e.g. =VLOOKUP with external references. These will eventually be dealt
    # with by the parser. However, as a workaround the class type of a token
    # can be changed via the repeat_formula interface. Thus, a _ref2d token can
    # be changed by the user to _ref2dA or _ref2dR to change its token class.
    #
    while (!args.empty?)
      token = args.shift

      if (token == '_arg')
        num_args = args.shift
      elsif (token == '_class')
        token = args.shift
        _class = @functions[token][2]
        # If _class is undef then it means that the function isn't valid.
        exit "Unknown function #{token}() in formula\n" if _class.nil?
        _classary.push(_class)
      elsif (token == '_vol')
        parse_str = parse_str + convert_volatile()
      elsif (token == 'ptgBool')
        token = args.shift
        parse_str = parse_str + convert_bool(token)
      elsif (token == '_num')
        token = args.shift
        parse_str = parse_str + convert_number(token)
      elsif (token == '_str')
        token = args.shift
        parse_str = parse_str + convert_string(token)
      elsif (token =~ /^_ref2d/)
        modifier  = token.sub(/_ref2d/, '')
        _class      = _classary[-1]
        _class      = 0 if modifier == 'R'
        _class      = 1 if modifier == 'V'
        token      = args.shift
        parse_str = parse_str + convert_ref2d(token, _class)
      elsif (token =~ /^_ref3d/)
        modifier  = token.sub(/_ref3d/,'')
        _class      = _classary[-1]
        _class      = 0 if modifier == 'R'
        _class      = 1 if modifier == 'V'
        token      = args.shift
        parse_str = parse_str + convert_ref3d(token, _class)
      elsif (token =~ /^_range2d/)
        modifier  = token.sub(/_range2d/,'')
        _class      = _classary[-1]
        _class      = 0 if modifier == 'R'
        _class      = 1 if modifier == 'V'
        token      = args.shift
        parse_str = parse_str + convert_range2d(token, _class)
      elsif (token =~ /^_range3d/)
        modifier  = token.sub(/_range3d/,'')
        _class      = _classary[-1]
        _class      = 0 if modifier == 'R'
        _class      = 1 if modifier == 'V'
        token      = args.shift
        parse_str = parse_str + convert_range3d(token, _class)
      elsif (token == '_func')
        token = args.shift
        parse_str = parse_str + convert_function(token, num_args.to_i)
        _classary.pop
        num_args = 0 # Reset after use
      elsif @ptg[token]
        parse_str = parse_str + [@ptg[token]].pack("C")
      else
        # Unrecognised token
        return nil
      end
    end


    return parse_str
  end

  def scan(formula)
    s = StringScanner.new(formula)
    q = []
    until s.eos?
      # order is important.
      if    s.scan(/(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?/)
        q.push [:NUMBER, s.matched]
      elsif s.scan(/"([^"]|"")*"/)
        q.push [:STRING, s.matched]
      elsif s.scan(/\$?[A-I]?[A-Z]\$?(\d+)?:\$?[A-I]?[A-Z]\$?(\d+)?/)
        q.push [:RANGE2D , s.matched]
      elsif s.scan(/[^!(,]+!\$?[A-I]?[A-Z]\$?(\d+)?:\$?[A-I]?[A-Z]\$?(\d+)?/)
        q.push [:RANGE3D , s.matched]
      elsif s.scan(/\$?[A-I]?[A-Z]\$?\d+/)
        q.push [:REF2D,  s.matched]
      elsif s.scan(/[^!(,]+!\$?[A-I]?[A-Z]\$?\d+/)
        q.push [:REF3D , s.matched]
      elsif s.scan(/'[^']+'!\$?[A-I]?[A-Z]\$?\d+/)
        q.push [:REF3D , s.matched]
      elsif s.scan(/<=/)
        q.push [:LE , s.matched]
      elsif s.scan(/>=/)
        q.push [:GE , s.matched]
      elsif s.scan(/<>/)
        q.push [:NE , s.matched]
      elsif s.scan(/</)
        q.push [:LT , s.matched]
      elsif s.scan(/>/)
        q.push [:GT , s.matched]
      elsif s.scan(/TRUE/)
        q.push [:TRUE, s.matched]
      elsif s.scan(/FALSE/)
        q.push [:FALSE, s.matched]
      elsif s.scan(/[A-Z0-9_.]+/)
        q.push [:FUNC,   s.matched]
      elsif s.scan(/\s+/)
        ;
      elsif s.scan(/./)
        q.push [s.matched, s.matched]
      end
    end
    q.push [:EOL, nil]
  end

  def parse(formula)
    @q = scan(formula)
    @q.push [false, nil]
    do_parse
  end

  def next_token
    @q.shift
  end

  def reverse(expression)
    expression.flatten
  end

  ###############################################################################
  #
  # get_ext_sheets()
  #
  # This semi-public method is used to update the hash of sheet names. It is
  # updated by the add_worksheet() method of the Workbook class.
  #
  # TODO
  #
  def get_ext_sheets
    # TODO
    refs = @ext_refs
    return refs

    #my @refs = sort {$refs{$a} <=> $refs{$b}} keys %refs;

    #foreach my $ref (@refs) {
    #    $ref = [split /:/, $ref];
    #}

    #return @refs;
  end


  ###############################################################################

  private

  ###############################################################################


  ###############################################################################
  #
  #  _check_volatile()
  #
  # Check if the formula contains a volatile function, i.e. a function that must
  # be recalculated each time a cell is updated. These formulas require a ptgAttr
  # with the volatile flag set as the first token in the parsed expression.
  #
  # Examples of volatile functions: RAND(), NOW(), TODAY()
  #
  def check_volatile(tokens)
    volatile = 0

    (0..tokens.size-1).each do |i|
      # If the next token is a function check if it is volatile.
      if tokens[i] == '_func' and @functions[tokens[i+1]][3] != 0
        volatile = 1
        break
      end
    end

    return volatile
  end

  ###############################################################################
  #
  # _convert_bool()
  #
  # Convert a boolean token to ptgBool
  #
  def convert_bool(bool)
    return [@ptg['ptgBool'], bool].pack("CC")
  end


  ###############################################################################
  #
  # _convert_number()
  #
  # Convert a number token to ptgInt or ptgNum
  #
  def convert_number(num)
    # Integer in the range 0..2**16-1
    if ((num =~ /^\d+$/) && (num.to_i <= 65535))
      return [@ptg['ptgInt'], num.to_i].pack("Cv")
    else  # A float
      num = [num].pack("d")
      num.reverse! if @byte_order != 0 && @byte_order != ''
      return [@ptg['ptgNum']].pack("C") + num
    end
  end

  ###############################################################################
  #
  # _convert_string()
  #
  # Convert a string to a ptg Str.
  #
  def convert_string(str)
    encoding = 0

    str.sub!(/^"/,'')   # Remove leading  "
    str.sub!(/"$/,'')   # Remove trailing "
    str.gsub!(/""/,'"') # Substitute Excel's escaped double quote "" for "

    length = str.length

    # Handle utf8 strings
    if str =~ NonAscii
      str = NKF.nkf('-w16L0 -m0 -W', str)
      encoding = 1
    end
    
    exit "String in formula has more than 255 chars\n" if length > 255

    return [@ptg['ptgStr'], length, encoding].pack("CCC") + str
  end

  ###############################################################################
  #
  # _convert_ref2d()
  #
  # Convert an Excel reference such as A1, $B2, C$3 or $D$4 to a ptgRefV.
  #
  def convert_ref2d(cell, _class)
    # Convert the cell reference
    row, col = cell_to_packed_rowcol(cell)

    # The ptg value depends on the class of the ptg.
    if    (_class == 0)
      ptgref = [@ptg['ptgRef']].pack("C")
    elsif (_class == 1)
      ptgref = [@ptg['ptgRefV']].pack("C")
    elsif (_class == 2)
      ptgref = [@ptg['ptgRefA']].pack("C")
    else
      exit "Unknown function class in formula\n"
    end

    return ptgref + row + col
  end

  ###############################################################################
  #
  # _convert_ref3d
  #
  # Convert an Excel 3d reference such as "Sheet1!A1" or "Sheet1:Sheet2!A1" to a
  # ptgRef3dV.
  #
  def convert_ref3d(token, _class)
    # Split the ref at the ! symbol
    ext_ref, cell = token.split('!')

    # Convert the external reference part
    ext_ref = pack_ext_ref(ext_ref)

    # Convert the cell reference part
    row, col = cell_to_packed_rowcol(cell)

    # The ptg value depends on the class of the ptg.
    if    (_class == 0)
      ptgref = [@ptg['ptgRef3d']].pack("C")
    elsif (_class == 1)
      ptgref = [@ptg['ptgRef3dV']].pack("C")
    elsif (_class == 2)
      ptgref = [@ptg['ptgRef3dA']].pack("C")
    else
      exit "Unknown function class in formula\n"
    end

    return ptgref + ext_ref + row + col
  end

  ###############################################################################
  #
  # _convert_range2d()
  #
  # Convert an Excel range such as A1:D4 or A:D to a ptgRefV.
  #
  def convert_range2d(range, _class)
    # Split the range into 2 cell refs
    cell1, cell2 = range.split(':')

    # A range such as A:D is equivalent to A1:D65536, so add rows as required
    cell1 = cell1 + '1'     unless cell1 =~ /\d/
    cell2 = cell2 + '65536' unless cell2 =~ /\d/

    # Convert the cell references
    row1, col1 = cell_to_packed_rowcol(cell1)
    row2, col2 = cell_to_packed_rowcol(cell2)

    # The ptg value depends on the class of the ptg.
    if    (_class == 0)
      ptgarea = [@ptg['ptgArea']].pack("C")
    elsif (_class == 1)
      ptgarea = [@ptg['ptgAreaV']].pack("C")
    elsif (_class == 2)
      ptgarea = [@ptg['ptgAreaA']].pack("C")
    else
      exit "Unknown function class in formula\n"
    end

    return ptgarea + row1 + row2 + col1 + col2
  end

  ###############################################################################
  #
  # _convert_range3d
  #
  # Convert an Excel 3d range such as "Sheet1!A1:D4" or "Sheet1:Sheet2!A1:D4" to
  # a ptgArea3dV.
  #
  def convert_range3d(token, _class)
    # Split the ref at the ! symbol
    ext_ref, range = token.split('!')

    # Convert the external reference part
    ext_ref = pack_ext_ref(ext_ref)

    # Split the range into 2 cell refs
    cell1, cell2 = range.split(':')

    # A range such as A:D is equivalent to A1:D65536, so add rows as required
    cell1 = cell1 + '1'     unless cell1 =~ /\d/
    cell2 = cell2 + '65536' unless cell2 =~ /\d/

    # Convert the cell references
    row1, col1 = cell_to_packed_rowcol(cell1)
    row2, col2 = cell_to_packed_rowcol(cell2)

    # The ptg value depends on the class of the ptg.
    if    (_class == 0)
      ptgarea = [@ptg['ptgArea3d']].pack("C")
    elsif (_class == 1)
      ptgarea = [@ptg['ptgArea3dV']].pack("C")
    elsif (_class == 2)
      ptgarea = [@ptg['ptgArea3dA']].pack("C")
    else
      exit "Unknown function class in formula\n"
    end

    return ptgarea + ext_ref + row1 + row2 + col1+ col2
  end

  ###############################################################################
  #
  # _pack_ext_ref()
  #
  # Convert the sheet name part of an external reference, for example "Sheet1" or
  # "Sheet1:Sheet2", to a packed structure.
  #
  def pack_ext_ref(ext_ref)
    ext_ref.sub!(/^'/,'')   # Remove leading  ' if any.
    ext_ref.sub!(/'$/,'')   # Remove trailing ' if any.

    # Check if there is a sheet range eg., Sheet1:Sheet2.
    if (ext_ref =~ /:/)
      sheet1, sheet2 = ext_ref.split(':')

      sheet1 = get_sheet_index(sheet1)
      sheet2 = get_sheet_index(sheet2)

      # Reverse max and min sheet numbers if necessary
      if (sheet1 > sheet2)
        sheet1, sheet2 = [sheet2, sheet1]
      end
    else
      # Single sheet name only.
      sheet1, sheet2 = [ext_ref, ext_ref]

      sheet1 = get_sheet_index(sheet1)
      sheet2 = sheet1
    end

    key = "#{sheet1}:#{sheet2}"

    unless @ext_refs[key]
      index = @ext_refs[key]
    else
      index = @ext_ref_count
      @ext_refs[key] = index
      @ext_ref_count += 1
    end

    return [index].pack("v")
  end

  ###############################################################################
  #
  # _get_sheet_index()
  #
  # Look up the index that corresponds to an external sheet name. The hash of
  # sheet names is updated by the add_worksheet() method of the Workbook class.
  #
  def get_sheet_index(sheet_name)
    # Handle utf8 sheetnames
    if sheet_name =~ NonAscii
      sheet_name = NKF.nkf('-w16B0 -m0 -W', sheet_name)
    end
    
    if @ext_sheets[sheet_name].nil?
      exit "Unknown sheet name #{sheet_name} in formula\n"
    else
      return @ext_sheets[sheet_name]
    end
  end

  ###############################################################################
  #
  # get_ext_ref_count()
  #
  # TODO This semi-public method is used to update the hash of sheet names. It is
  # updated by the add_worksheet() method of the Workbook class.
  #
  def get_ext_ref_count
    return @ext_ref_count
  end

  ###############################################################################
  #
  # _convert_function()
  #
  # Convert a function to a ptgFunc or ptgFuncVarV depending on the number of
  # args that it takes.
  #
  def convert_function(token, num_args)
    exit "Unknown function #{token}() in formula\n" if @functions[token][0].nil?

    args = @functions[token][1]

    # Fixed number of args eg. TIME($i,$j,$k).
    if (args >= 0)
      # Check that the number of args is valid.
      if (args != num_args)
        raise "Incorrect number of arguments for #{token}() in formula\n";
      else
        return [@ptg['ptgFuncV'], @functions[token][0]].pack("Cv")
      end
    end

    # Variable number of args eg. SUM(i,j,k, ..).
    if (args == -1)
      return [@ptg['ptgFuncVarV'], num_args, @functions[token][0]].pack("CCv")
    end
  end

  ###############################################################################
  #
  # _cell_to_rowcol($cell_ref)
  #
  # Convert an Excel cell reference such as A1 or $B2 or C$3 or $D$4 to a zero
  # indexed row and column number. Also returns two boolean values to indicate
  # whether the row or column are relative references.
  # TODO use function in Utility.pm
  #
  def cell_to_rowcol(cell)
    cell =~ /(\$?)([A-I]?[A-Z])(\$?)(\d+)/

    col_rel = $1 == "" ? 1 : 0
    col     = $2
    row_rel = $3 == "" ? 1 : 0
    row     = $4.to_i

    # Convert base26 column string to a number.
    # All your Base are belong to us.
    chars  = col.split(//)
    expn   = 0
    col    = 0

    while (!chars.empty?)
      char = chars.pop   # LS char first
      col  = col + (char[0] - "A"[0] + 1) * (26**expn)
      expn += 1
    end
    # Convert 1-index to zero-index
    row -= 1
    col -= 1

    return [row, col, row_rel, col_rel]
  end

  ###############################################################################
  #
  # _cell_to_packed_rowcol($row, $col, $row_rel, $col_rel)
  #
  # pack() row and column into the required 3 byte format.
  #
  def cell_to_packed_rowcol(cell)
    row, col, row_rel, col_rel = cell_to_rowcol(cell)

    exit "Column #{cell} greater than IV in formula\n" if col >= 256
    exit "Row #{cell} greater than 65536 in formula\n" if row >= 65536

    # Set the high bits to indicate if row or col are relative.
    col    |= col_rel << 14
    col    |= row_rel << 15

    row     = [row].pack('v')
    col     = [col].pack('v')

    return [row, col]
  end

  ###############################################################################
  #
  # _initialize_hashes()
  #
  def initialize_hashes

    # The Excel ptg indices
    @ptg = {
      'ptgExp'            => 0x01,
      'ptgTbl'            => 0x02,
      'ptgAdd'            => 0x03,
      'ptgSub'            => 0x04,
      'ptgMul'            => 0x05,
      'ptgDiv'            => 0x06,
      'ptgPower'          => 0x07,
      'ptgConcat'         => 0x08,
      'ptgLT'             => 0x09,
      'ptgLE'             => 0x0A,
      'ptgEQ'             => 0x0B,
      'ptgGE'             => 0x0C,
      'ptgGT'             => 0x0D,
      'ptgNE'             => 0x0E,
      'ptgIsect'          => 0x0F,
      'ptgUnion'          => 0x10,
      'ptgRange'          => 0x11,
      'ptgUplus'          => 0x12,
      'ptgUminus'         => 0x13,
      'ptgPercent'        => 0x14,
      'ptgParen'          => 0x15,
      'ptgMissArg'        => 0x16,
      'ptgStr'            => 0x17,
      'ptgAttr'           => 0x19,
      'ptgSheet'          => 0x1A,
      'ptgEndSheet'       => 0x1B,
      'ptgErr'            => 0x1C,
      'ptgBool'           => 0x1D,
      'ptgInt'            => 0x1E,
      'ptgNum'            => 0x1F,
      'ptgArray'          => 0x20,
      'ptgFunc'           => 0x21,
      'ptgFuncVar'        => 0x22,
      'ptgName'           => 0x23,
      'ptgRef'            => 0x24,
      'ptgArea'           => 0x25,
      'ptgMemArea'        => 0x26,
      'ptgMemErr'         => 0x27,
      'ptgMemNoMem'       => 0x28,
      'ptgMemFunc'        => 0x29,
      'ptgRefErr'         => 0x2A,
      'ptgAreaErr'        => 0x2B,
      'ptgRefN'           => 0x2C,
      'ptgAreaN'          => 0x2D,
      'ptgMemAreaN'       => 0x2E,
      'ptgMemNoMemN'      => 0x2F,
      'ptgNameX'          => 0x39,
      'ptgRef3d'          => 0x3A,
      'ptgArea3d'         => 0x3B,
      'ptgRefErr3d'       => 0x3C,
      'ptgAreaErr3d'      => 0x3D,
      'ptgArrayV'         => 0x40,
      'ptgFuncV'          => 0x41,
      'ptgFuncVarV'       => 0x42,
      'ptgNameV'          => 0x43,
      'ptgRefV'           => 0x44,
      'ptgAreaV'          => 0x45,
      'ptgMemAreaV'       => 0x46,
      'ptgMemErrV'        => 0x47,
      'ptgMemNoMemV'      => 0x48,
      'ptgMemFuncV'       => 0x49,
      'ptgRefErrV'        => 0x4A,
      'ptgAreaErrV'       => 0x4B,
      'ptgRefNV'          => 0x4C,
      'ptgAreaNV'         => 0x4D,
      'ptgMemAreaNV'      => 0x4E,
      'ptgMemNoMemN'      => 0x4F,
      'ptgFuncCEV'        => 0x58,
      'ptgNameXV'         => 0x59,
      'ptgRef3dV'         => 0x5A,
      'ptgArea3dV'        => 0x5B,
      'ptgRefErr3dV'      => 0x5C,
      'ptgAreaErr3d'      => 0x5D,
      'ptgArrayA'         => 0x60,
      'ptgFuncA'          => 0x61,
      'ptgFuncVarA'       => 0x62,
      'ptgNameA'          => 0x63,
      'ptgRefA'           => 0x64,
      'ptgAreaA'          => 0x65,
      'ptgMemAreaA'       => 0x66,
      'ptgMemErrA'        => 0x67,
      'ptgMemNoMemA'      => 0x68,
      'ptgMemFuncA'       => 0x69,
      'ptgRefErrA'        => 0x6A,
      'ptgAreaErrA'       => 0x6B,
      'ptgRefNA'          => 0x6C,
      'ptgAreaNA'         => 0x6D,
      'ptgMemAreaNA'      => 0x6E,
      'ptgMemNoMemN'      => 0x6F,
      'ptgFuncCEA'        => 0x78,
      'ptgNameXA'         => 0x79,
      'ptgRef3dA'         => 0x7A,
      'ptgArea3dA'        => 0x7B,
      'ptgRefErr3dA'      => 0x7C,
      'ptgAreaErr3d'      => 0x7D
    };

    # Thanks to Michael Meeks and Gnumeric for the initial arg values.
    #
    # The following hash was generated by "function_locale.pl" in the distro.
    # Refer to function_locale.pl for non-English function names.
    #
    # The array elements are as follow:
    # ptg:   The Excel function ptg code.
    # args:  The number of arguments that the function takes:
    #           >=0 is a fixed number of arguments.
    #           -1  is a variable  number of arguments.
    # class: The reference, value or array class of the function args.
    # vol:   The function is volatile.
    #
    @functions  = {
      #                                     ptg  args  class  vol
      'COUNT'                         => [   0,   -1,    0,    0 ],
      'IF'                            => [   1,   -1,    1,    0 ],
      'ISNA'                          => [   2,    1,    1,    0 ],
      'ISERROR'                       => [   3,    1,    1,    0 ],
      'SUM'                           => [   4,   -1,    0,    0 ],
      'AVERAGE'                       => [   5,   -1,    0,    0 ],
      'MIN'                           => [   6,   -1,    0,    0 ],
      'MAX'                           => [   7,   -1,    0,    0 ],
      'ROW'                           => [   8,   -1,    0,    0 ],
      'COLUMN'                        => [   9,   -1,    0,    0 ],
      'NA'                            => [  10,    0,    0,    0 ],
      'NPV'                           => [  11,   -1,    1,    0 ],
      'STDEV'                         => [  12,   -1,    0,    0 ],
      'DOLLAR'                        => [  13,   -1,    1,    0 ],
      'FIXED'                         => [  14,   -1,    1,    0 ],
      'SIN'                           => [  15,    1,    1,    0 ],
      'COS'                           => [  16,    1,    1,    0 ],
      'TAN'                           => [  17,    1,    1,    0 ],
      'ATAN'                          => [  18,    1,    1,    0 ],
      'PI'                            => [  19,    0,    1,    0 ],
      'SQRT'                          => [  20,    1,    1,    0 ],
      'EXP'                           => [  21,    1,    1,    0 ],
      'LN'                            => [  22,    1,    1,    0 ],
      'LOG10'                         => [  23,    1,    1,    0 ],
      'ABS'                           => [  24,    1,    1,    0 ],
      'INT'                           => [  25,    1,    1,    0 ],
      'SIGN'                          => [  26,    1,    1,    0 ],
      'ROUND'                         => [  27,    2,    1,    0 ],
      'LOOKUP'                        => [  28,   -1,    0,    0 ],
      'INDEX'                         => [  29,   -1,    0,    1 ],
      'REPT'                          => [  30,    2,    1,    0 ],
      'MID'                           => [  31,    3,    1,    0 ],
      'LEN'                           => [  32,    1,    1,    0 ],
      'VALUE'                         => [  33,    1,    1,    0 ],
      'TRUE'                          => [  34,    0,    1,    0 ],
      'FALSE'                         => [  35,    0,    1,    0 ],
      'AND'                           => [  36,   -1,    1,    0 ],
      'OR'                            => [  37,   -1,    1,    0 ],
      'NOT'                           => [  38,    1,    1,    0 ],
      'MOD'                           => [  39,    2,    1,    0 ],
      'DCOUNT'                        => [  40,    3,    0,    0 ],
      'DSUM'                          => [  41,    3,    0,    0 ],
      'DAVERAGE'                      => [  42,    3,    0,    0 ],
      'DMIN'                          => [  43,    3,    0,    0 ],
      'DMAX'                          => [  44,    3,    0,    0 ],
      'DSTDEV'                        => [  45,    3,    0,    0 ],
      'VAR'                           => [  46,   -1,    0,    0 ],
      'DVAR'                          => [  47,    3,    0,    0 ],
      'TEXT'                          => [  48,    2,    1,    0 ],
      'LINEST'                        => [  49,   -1,    0,    0 ],
      'TREND'                         => [  50,   -1,    0,    0 ],
      'LOGEST'                        => [  51,   -1,    0,    0 ],
      'GROWTH'                        => [  52,   -1,    0,    0 ],
      'PV'                            => [  56,   -1,    1,    0 ],
      'FV'                            => [  57,   -1,    1,    0 ],
      'NPER'                          => [  58,   -1,    1,    0 ],
      'PMT'                           => [  59,   -1,    1,    0 ],
      'RATE'                          => [  60,   -1,    1,    0 ],
      'MIRR'                          => [  61,    3,    0,    0 ],
      'IRR'                           => [  62,   -1,    0,    0 ],
      'RAND'                          => [  63,    0,    1,    1 ],
      'MATCH'                         => [  64,   -1,    0,    0 ],
      'DATE'                          => [  65,    3,    1,    0 ],
      'TIME'                          => [  66,    3,    1,    0 ],
      'DAY'                           => [  67,    1,    1,    0 ],
      'MONTH'                         => [  68,    1,    1,    0 ],
      'YEAR'                          => [  69,    1,    1,    0 ],
      'WEEKDAY'                       => [  70,   -1,    1,    0 ],
      'HOUR'                          => [  71,    1,    1,    0 ],
      'MINUTE'                        => [  72,    1,    1,    0 ],
      'SECOND'                        => [  73,    1,    1,    0 ],
      'NOW'                           => [  74,    0,    1,    1 ],
      'AREAS'                         => [  75,    1,    0,    1 ],
      'ROWS'                          => [  76,    1,    0,    1 ],
      'COLUMNS'                       => [  77,    1,    0,    1 ],
      'OFFSET'                        => [  78,   -1,    0,    1 ],
      'SEARCH'                        => [  82,   -1,    1,    0 ],
      'TRANSPOSE'                     => [  83,    1,    1,    0 ],
      'TYPE'                          => [  86,    1,    1,    0 ],
      'ATAN2'                         => [  97,    2,    1,    0 ],
      'ASIN'                          => [  98,    1,    1,    0 ],
      'ACOS'                          => [  99,    1,    1,    0 ],
      'CHOOSE'                        => [ 100,   -1,    1,    0 ],
      'HLOOKUP'                       => [ 101,   -1,    0,    0 ],
      'VLOOKUP'                       => [ 102,   -1,    0,    0 ],
      'ISREF'                         => [ 105,    1,    0,    0 ],
      'LOG'                           => [ 109,   -1,    1,    0 ],
      'CHAR'                          => [ 111,    1,    1,    0 ],
      'LOWER'                         => [ 112,    1,    1,    0 ],
      'UPPER'                         => [ 113,    1,    1,    0 ],
      'PROPER'                        => [ 114,    1,    1,    0 ],
      'LEFT'                          => [ 115,   -1,    1,    0 ],
      'RIGHT'                         => [ 116,   -1,    1,    0 ],
      'EXACT'                         => [ 117,    2,    1,    0 ],
      'TRIM'                          => [ 118,    1,    1,    0 ],
      'REPLACE'                       => [ 119,    4,    1,    0 ],
      'SUBSTITUTE'                    => [ 120,   -1,    1,    0 ],
      'CODE'                          => [ 121,    1,    1,    0 ],
      'FIND'                          => [ 124,   -1,    1,    0 ],
      'CELL'                          => [ 125,   -1,    0,    1 ],
      'ISERR'                         => [ 126,    1,    1,    0 ],
      'ISTEXT'                        => [ 127,    1,    1,    0 ],
      'ISNUMBER'                      => [ 128,    1,    1,    0 ],
      'ISBLANK'                       => [ 129,    1,    1,    0 ],
      'T'                             => [ 130,    1,    0,    0 ],
      'N'                             => [ 131,    1,    0,    0 ],
      'DATEVALUE'                     => [ 140,    1,    1,    0 ],
      'TIMEVALUE'                     => [ 141,    1,    1,    0 ],
      'SLN'                           => [ 142,    3,    1,    0 ],
      'SYD'                           => [ 143,    4,    1,    0 ],
      'DDB'                           => [ 144,   -1,    1,    0 ],
      'INDIRECT'                      => [ 148,   -1,    1,    1 ],
      'CALL'                          => [ 150,   -1,    1,    0 ],
      'CLEAN'                         => [ 162,    1,    1,    0 ],
      'MDETERM'                       => [ 163,    1,    2,    0 ],
      'MINVERSE'                      => [ 164,    1,    2,    0 ],
      'MMULT'                         => [ 165,    2,    2,    0 ],
      'IPMT'                          => [ 167,   -1,    1,    0 ],
      'PPMT'                          => [ 168,   -1,    1,    0 ],
      'COUNTA'                        => [ 169,   -1,    0,    0 ],
      'PRODUCT'                       => [ 183,   -1,    0,    0 ],
      'FACT'                          => [ 184,    1,    1,    0 ],
      'DPRODUCT'                      => [ 189,    3,    0,    0 ],
      'ISNONTEXT'                     => [ 190,    1,    1,    0 ],
      'STDEVP'                        => [ 193,   -1,    0,    0 ],
      'VARP'                          => [ 194,   -1,    0,    0 ],
      'DSTDEVP'                       => [ 195,    3,    0,    0 ],
      'DVARP'                         => [ 196,    3,    0,    0 ],
      'TRUNC'                         => [ 197,   -1,    1,    0 ],
      'ISLOGICAL'                     => [ 198,    1,    1,    0 ],
      'DCOUNTA'                       => [ 199,    3,    0,    0 ],
      'ROUNDUP'                       => [ 212,    2,    1,    0 ],
      'ROUNDDOWN'                     => [ 213,    2,    1,    0 ],
      'RANK'                          => [ 216,   -1,    0,    0 ],
      'ADDRESS'                       => [ 219,   -1,    1,    0 ],
      'DAYS360'                       => [ 220,   -1,    1,    0 ],
      'TODAY'                         => [ 221,    0,    1,    1 ],
      'VDB'                           => [ 222,   -1,    1,    0 ],
      'MEDIAN'                        => [ 227,   -1,    0,    0 ],
      'SUMPRODUCT'                    => [ 228,   -1,    2,    0 ],
      'SINH'                          => [ 229,    1,    1,    0 ],
      'COSH'                          => [ 230,    1,    1,    0 ],
      'TANH'                          => [ 231,    1,    1,    0 ],
      'ASINH'                         => [ 232,    1,    1,    0 ],
      'ACOSH'                         => [ 233,    1,    1,    0 ],
      'ATANH'                         => [ 234,    1,    1,    0 ],
      'DGET'                          => [ 235,    3,    0,    0 ],
      'INFO'                          => [ 244,    1,    1,    1 ],
      'DB'                            => [ 247,   -1,    1,    0 ],
      'FREQUENCY'                     => [ 252,    2,    0,    0 ],
      'ERROR.TYPE'                    => [ 261,    1,    1,    0 ],
      'REGISTER.ID'                   => [ 267,   -1,    1,    0 ],
      'AVEDEV'                        => [ 269,   -1,    0,    0 ],
      'BETADIST'                      => [ 270,   -1,    1,    0 ],
      'GAMMALN'                       => [ 271,    1,    1,    0 ],
      'BETAINV'                       => [ 272,   -1,    1,    0 ],
      'BINOMDIST'                     => [ 273,    4,    1,    0 ],
      'CHIDIST'                       => [ 274,    2,    1,    0 ],
      'CHIINV'                        => [ 275,    2,    1,    0 ],
      'COMBIN'                        => [ 276,    2,    1,    0 ],
      'CONFIDENCE'                    => [ 277,    3,    1,    0 ],
      'CRITBINOM'                     => [ 278,    3,    1,    0 ],
      'EVEN'                          => [ 279,    1,    1,    0 ],
      'EXPONDIST'                     => [ 280,    3,    1,    0 ],
      'FDIST'                         => [ 281,    3,    1,    0 ],
      'FINV'                          => [ 282,    3,    1,    0 ],
      'FISHER'                        => [ 283,    1,    1,    0 ],
      'FISHERINV'                     => [ 284,    1,    1,    0 ],
      'FLOOR'                         => [ 285,    2,    1,    0 ],
      'GAMMADIST'                     => [ 286,    4,    1,    0 ],
      'GAMMAINV'                      => [ 287,    3,    1,    0 ],
      'CEILING'                       => [ 288,    2,    1,    0 ],
      'HYPGEOMDIST'                   => [ 289,    4,    1,    0 ],
      'LOGNORMDIST'                   => [ 290,    3,    1,    0 ],
      'LOGINV'                        => [ 291,    3,    1,    0 ],
      'NEGBINOMDIST'                  => [ 292,    3,    1,    0 ],
      'NORMDIST'                      => [ 293,    4,    1,    0 ],
      'NORMSDIST'                     => [ 294,    1,    1,    0 ],
      'NORMINV'                       => [ 295,    3,    1,    0 ],
      'NORMSINV'                      => [ 296,    1,    1,    0 ],
      'STANDARDIZE'                   => [ 297,    3,    1,    0 ],
      'ODD'                           => [ 298,    1,    1,    0 ],
      'PERMUT'                        => [ 299,    2,    1,    0 ],
      'POISSON'                       => [ 300,    3,    1,    0 ],
      'TDIST'                         => [ 301,    3,    1,    0 ],
      'WEIBULL'                       => [ 302,    4,    1,    0 ],
      'SUMXMY2'                       => [ 303,    2,    2,    0 ],
      'SUMX2MY2'                      => [ 304,    2,    2,    0 ],
      'SUMX2PY2'                      => [ 305,    2,    2,    0 ],
      'CHITEST'                       => [ 306,    2,    2,    0 ],
      'CORREL'                        => [ 307,    2,    2,    0 ],
      'COVAR'                         => [ 308,    2,    2,    0 ],
      'FORECAST'                      => [ 309,    3,    2,    0 ],
      'FTEST'                         => [ 310,    2,    2,    0 ],
      'INTERCEPT'                     => [ 311,    2,    2,    0 ],
      'PEARSON'                       => [ 312,    2,    2,    0 ],
      'RSQ'                           => [ 313,    2,    2,    0 ],
      'STEYX'                         => [ 314,    2,    2,    0 ],
      'SLOPE'                         => [ 315,    2,    2,    0 ],
      'TTEST'                         => [ 316,    4,    2,    0 ],
      'PROB'                          => [ 317,   -1,    2,    0 ],
      'DEVSQ'                         => [ 318,   -1,    0,    0 ],
      'GEOMEAN'                       => [ 319,   -1,    0,    0 ],
      'HARMEAN'                       => [ 320,   -1,    0,    0 ],
      'SUMSQ'                         => [ 321,   -1,    0,    0 ],
      'KURT'                          => [ 322,   -1,    0,    0 ],
      'SKEW'                          => [ 323,   -1,    0,    0 ],
      'ZTEST'                         => [ 324,   -1,    0,    0 ],
      'LARGE'                         => [ 325,    2,    0,    0 ],
      'SMALL'                         => [ 326,    2,    0,    0 ],
      'QUARTILE'                      => [ 327,    2,    0,    0 ],
      'PERCENTILE'                    => [ 328,    2,    0,    0 ],
      'PERCENTRANK'                   => [ 329,   -1,    0,    0 ],
      'MODE'                          => [ 330,   -1,    2,    0 ],
      'TRIMMEAN'                      => [ 331,    2,    0,    0 ],
      'TINV'                          => [ 332,    2,    1,    0 ],
      'CONCATENATE'                   => [ 336,   -1,    1,    0 ],
      'POWER'                         => [ 337,    2,    1,    0 ],
      'RADIANS'                       => [ 342,    1,    1,    0 ],
      'DEGREES'                       => [ 343,    1,    1,    0 ],
      'SUBTOTAL'                      => [ 344,   -1,    0,    0 ],
      'SUMIF'                         => [ 345,   -1,    0,    0 ],
      'COUNTIF'                       => [ 346,    2,    0,    0 ],
      'COUNTBLANK'                    => [ 347,    1,    0,    0 ],
      'ROMAN'                         => [ 354,   -1,    1,    0 ]
    }

  end


end

if $0 ==__FILE__


  parser = Formula.new
  puts
  puts 'type "Q" to quit.'
  puts
  while true
    puts
    print '? '
    str = gets.chop!
    break if /q/i =~ str
    begin
      e = parser.parse(str)
      p   parser.reverse(e)
    rescue ParseError
      puts $!
    end
  end

end
