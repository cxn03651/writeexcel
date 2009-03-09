require 'digest/md5'
require 'Tempfile'
require 'biffwriter'
require 'olewriter'
require 'formula'

class Workbook < BIFFWriter
  BOF = 11
  EOF = 4
  SheetName = "Sheet"

  attr_accessor :date_system, :str_unique, :biff_only
  attr_reader :formats, :xf_index, :worksheets, :extsst_buckets, :extsst_bucket_size
  attr_writer :mso_size

  ###############################################################################
  #
  # new()
  #
  # Constructor. Creates a new Workbook object from a BIFFwriter object.
  #
  def initialize(filename)
    super
    @filename              = filename
    @parser                = Formula.new(@byte_order)
    @tempdir               = nil
    @date_1904             = false
    @activesheet           = 0
    @firstsheet            = 0
    @selected              = 0
    @xf_index              = 0
    @fileclosed            = 0
    @biffsize              = 0
    @sheetname             = "Sheet"
    @url_format            = ''
    @codepage              = 0x04E4
    @worksheets            = []
    @sheetnames            = []
    @formats               = []
    @palette               = []
    @biff_only             = 0

    @using_tmpfile         = 1
    @filehandle            = ""
    @temp_file             = ""
    @internal_fh           = 0
    @fh_out                = ""

    @str_total             = 0
    @str_unique            = 0
    @str_table             = {}
    @str_array             = []
    @str_block_sizes       = []
    @extsst_offsets        = []
    @extsst_buckets        = 0
    @extsst_bucket_size    = 0

    @ext_ref_count         = 0
    @ext_refs              = {}

    @mso_clusters          = []
    @mso_size              = 0

    @hideobj               = 0
    @compatibility         = 0

    @add_doc_properties    = 0
    @localtime             = Time.now

    # Add the in-built style formats and the default cell format.
    add_format(:type => 1)                        #  0 Normal
    add_format(:type => 1)                        #  1 RowLevel 1
    add_format(:type => 1)                        #  2 RowLevel 2
    add_format(:type => 1)                        #  3 RowLevel 3
    add_format(:type => 1)                        #  4 RowLevel 4
    add_format(:type => 1)                        #  5 RowLevel 5
    add_format(:type => 1)                        #  6 RowLevel 6
    add_format(:type => 1)                        #  7 RowLevel 7
    add_format(:type => 1)                        #  8 ColLevel 1
    add_format(:type => 1)                        #  9 ColLevel 2
    add_format(:type => 1)                        # 10 ColLevel 3
    add_format(:type => 1)                        # 11 ColLevel 4
    add_format(:type => 1)                        # 12 ColLevel 5
    add_format(:type => 1)                        # 13 ColLevel 6
    add_format(:type => 1)                        # 14 ColLevel 7
    add_format()                                  # 15 Cell XF
    add_format(:type => 1, :num_format => 0x2B)   # 16 Comma
    add_format(:type => 1, :num_format => 0x29)   # 17 Comma[0]
    add_format(:type => 1, :num_format => 0x2C)   # 18 Currency
    add_format(:type => 1, :num_format => 0x2A)   # 19 Currency[0]
    add_format(:type => 1, :num_format => 0x09)   # 20 Percent

    # Add the default format for hyperlinks
    @url_format = add_format(:color => 'blue', :underline => 1)

    # Convert the filename to a filehandle to pass to the OLE writer when the
    # file is closed. If the filename is a reference it is assumed that it is
    # a valid filehandle.
    #
    if filename.kind_of?(String) && filename != ''
      @fh_out      = open(filename, "wb")
      @internal_fh = 1
    else
      print "Workbook#new - filename required."
      exit
    end

    # Set colour palette.
    set_palette_xl97

    _initialize
    get_checksum_method
  end


  ###############################################################################
  #
  # _initialize()
  #
  # Open a tmp file to store the majority of the Worksheet data. If this fails,
  # for example due to write permissions, store the data in memory. This can be
  # slow for large files.
  #
  # TODO: Move this and other methods shared with Worksheet up into BIFFWriter.
  #
  def _initialize
    basename = 'spreadsheetwriteexcelworkbook'

    if @tempdir.nil? || @tempdir == ''
      fh = Tempfile.new(basename)
    else
      fh = Tempfile.new(basename, @tempdir)
    end

    if fh
      @filehandle = fh
    else
      @using_tmpfile = 0
    end
  end

  ###############################################################################
  #
  # _get_checksum_method.
  #
  # Check for modules available to calculate image checksum. Excel uses MD4 but
  # MD5 will also work.
  #
  # ------- cxn03651 add -------
  # md5 can use in ruby. so, @checksum_method is always 3.

  def get_checksum_method
    @checksum_method = 3
  end

  ###############################################################################
  #
  # _append(), overloaded.
  #
  # Store Worksheet data in memory using the base class _append() or to a
  # temporary file, the default.
  #
  def append(*args)
    data = ''
    if @using_tmpfile != 0
      data = args.join('')

      # Add CONTINUE records if necessary
      data = add_continue(data) if data.length > @limit

      #         # Protect print() from -l on the command line.
      #         local $\ = undef;
      #
      @filehandle.write data
      @datasize += data.length
    else
      data = super(*args)
    end

    return data
  end

  ###############################################################################
  #
  # get_data().
  #
  # Retrieves data from memory in one chunk, or from disk in $buffer
  # sized chunks.
  #
  def get_data
    bufsize = 4096

    # Return data stored in memory
    unless @data.nil?
      tmp   = @data
      @data = nil
      @filehandle.seek(0, IO::SEEK_SET) if @using_tmpfile != 0
      return tmp
    end

    # Return data stored on disk
    if @using_tmpfile != 0
      tmp = @filehandle.read(bufsize)
      return tmp unless tmp.nil?
    end

    # No data to return
    return nil
  end

  ###############################################################################
  #
  # close()
  #
  # Calls finalization methods and explicitly close the OLEwriter file
  # handle.
  #
  def close
    return if @fileclosed != 0   # Prevent close() from being called twice.

    @fileclosed = 1
    return store_workbook
  end

  ###############################################################################
  #
  # sheets(slice,...)
  #
  # An accessor for the _worksheets[] array
  #
  # Returns: an optionally sliced list of the worksheet objects in a workbook.
  #
  def sheets(*args)
    if args.empty?
      @worksheets
    else
      ary = []
      args.each do |i|
        ary << @worksheets[i]
      end
      ary
    end
  end

  ###############################################################################
  #
  # add_worksheet($name, $encoding)
  #
  # Add a new worksheet to the Excel workbook.
  #
  # Returns: reference to a worksheet object
  #
  def add_worksheet(name = '', encoding = 0)
    name, encoding = check_sheetname(name, encoding)

    index = @worksheets.size

    worksheet = Worksheet.new(
      name,
      index,
      encoding,
      @activesheet,
      @firstsheet,
      @url_format,
      @parser,
      @tempdir,
      @str_total,
      @str_unique,
      @str_table,
      @date_1904,
      @compatibility
    )
    @worksheets[index] = worksheet     # Store ref for iterator
    @sheetnames[index] = name          # Store EXTERNSHEET names
    #       @parser->set_ext_sheets($name, $index) # Store names in Formula.pm
    return worksheet
  end

  ###############################################################################
  #
  # add_chart_ext($filename, $name)
  #
  # Add an externally created chart.
  #
  #
  def add_chart_ext(filename, name, encoding = nil)
    index    = @worksheets.size

    name, encoding = check_sheetname(name, encoding)

    init_data = [
      filename,
      name,
      index,
      encoding,
      @activesheet,
      @firstsheet
    ]

    worksheet = Chart.new(*init_data)
    @worksheets[index] = worksheet      # Store ref for iterator
    @sheetnames[index] = name           # Store EXTERNSHEET names
    @parser.set_ext_sheets(name, index) # Store names in Formula.pm
    return worksheet
  end

  ###############################################################################
  #
  # _check_sheetname($name, $encoding)
  #
  # Check for valid worksheet names. We check the length, if it contains any
  # invalid characters and if the name is unique in the workbook.
  #
  def check_sheetname(name, encoding = 0)
    name = '' if name.nil?
    limit           = encoding != 0 ? 62 : 31
    invalid_char    = %r![\[\]:*?/\\]!

    # Supply default "Sheet" name if none has been defined.
    index     = @worksheets.size
    sheetname = @sheetname

    if name == ""
      name     = sheetname + (index+1).to_s
      encoding = 0
    end

    # Check that sheetname is <= 31 (1 or 2 byte chars). Excel limit.
    raise "Sheetname #{name} must be <= 31 chars" if name.length > limit

    # Check that Unicode sheetname has an even number of bytes
    if encoding == 1 and name.length % 2
      raise 'Odd number of bytes in Unicode worksheet name:' + name
    end

    # Check that sheetname doesn't contain any invalid characters
    if encoding != 1 and name =~ invalid_char
      # Check ASCII names
      raise 'Invalid character []:*?/\\ in worksheet name: ' + name
    else
      #         # Extract any 8bit clean chars from the UTF16 name and validate them.
      #         for my $wchar ($name =~ /../sg) {
      #            my ($hi, $lo) = unpack "aa", $wchar;
      #            if ($hi eq "\0" and $lo =~ $invalid_char)
      #               raise 'Invalid character []:*?/\\ in worksheet name: ' + name
      #            end
      #        }
    end

    # Check that the worksheet name doesn't already exist since this is a fatal
    # error in Excel 97. The check must also exclude case insensitive matches
    # since the names 'Sheet1' and 'sheet1' are equivalent. The tests also have
    # to take the encoding into account.
    #
    @worksheets.each do |worksheet|
      name_a  = name
      encd_a  = encoding
      name_b  = worksheet.name
      encd_b  = worksheet.encoding
      error   = 0;

      if    encd_a == 0 and encd_b == 0
        error  = 1 if name_a.downcase == name_b.downcase
      elsif encd_a == 0 and encd_b == 1
        name_a = [name_a].unpack("C*").pack("n*")
        error  = 1 if name_a.downcase == name_b.downcase
      elsif encd_a == 1 and encd_b == 0
        name_b = [name_b].unpack("C*").pack("n*")
        error  = 1 if name_a.downcase == name_b.downcase
      elsif encd_a == 1 and encd_b == 1
        #            # We can do a true case insensitive test with Perl 5.8 and utf8.
        #            if ($] >= 5.008) {
        #               $name_a = Encode::decode("UTF-16BE", $name_a);
        #               $name_b = Encode::decode("UTF-16BE", $name_b);
        #               $error  = 1 if lc($name_a) eq lc($name_b);
        #            }
        #            else {
        #               # We can't easily do a case insensitive test of the UTF16 names.
        #               # As a special case we check if all of the high bytes are nulls and
        #               # then do an ASCII style case insensitive test.
        #
        #               # Strip out the high bytes (funkily).
        #               my $hi_a = grep {ord} $name_a =~ /(.)./sg;
        #               my $hi_b = grep {ord} $name_b =~ /(.)./sg;
        #
        #               if ($hi_a or $hi_b) {
        #                  $error  = 1 if    $name_a  eq    $name_b;
        #               }
        #               else {
        #                  $error  = 1 if lc($name_a) eq lc($name_b);
        #               }
        #            }
        #         }
        #         # If any of the cases failed we throw the error here.
      end
      if error != 0
        raise "Worksheet name '#{name}', with case ignored, " +
        "is already in use";
      end
    end
    return [name,  encoding]
  end

  ###############################################################################
  #
  # add_format(%properties)
  #
  # Add a new format to the Excel workbook. This adds an XF record and
  # a FONT record. Also, pass any properties to the Format::new().
  #
  def add_format(*args)
    format = Format.new(@xf_index, args)
    @xf_index += 1
    @formats.push format # Store format reference
    return format
  end

  ###############################################################################
  #
  # compatibility_mode()
  #
  # Set the compatibility mode.
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
  def compatibility_mode(mode = 1)
    unless sheets.empty?
      raise "compatibility_mode() must be called before add_worksheet()"
    end
    @compatibility = mode
  end

  ###############################################################################
  #
  # set_1904()
  #
  # Set the date system: false = 1900 (the default), true = 1904
  #
  def set_1904(mode = true)
    unless sheets.empty?
      raise "set_1904() must be called before add_worksheet()"
    end
    @date_1904 = mode
  end

  ###############################################################################
  #
  # set_custom_color()
  #
  # Change the RGB components of the elements in the colour palette.
  #
  def set_custom_color(index = nil, red = nil, green = nil, blue = nil)
    # Match a HTML #xxyyzz style parameter
    if !red.nil? && red =~ /^#(\w\w)(\w\w)(\w\w)/
      red   = $1.hex
      green = $2.hex
      blue  = $3.hex
    end

    # Check that the colour index is the right range
    if index < 8 || index > 64
      raise "Color index #{index} outside range: 8 <= index <= 64";
    end

    # Check that the colour components are in the right range
    if (red   < 0 || red   > 255) ||
      (green < 0 || green > 255) ||
      (blue  < 0 || blue  > 255)
      raise "Color component outside range: 0 <= color <= 255";
    end

    index -=8       # Adjust colour index (wingless dragonfly)

    # Set the RGB value
    @palette[index] = [red, green, blue, 0]

    return index +8
  end

  ###############################################################################
  #
  # set_palette_xl97()
  #
  # Sets the colour palette to the Excel 97+ default.
  #
  def set_palette_xl97
    @palette = [
      [0x00, 0x00, 0x00, 0x00],   # 8
      [0xff, 0xff, 0xff, 0x00],   # 9
      [0xff, 0x00, 0x00, 0x00],   # 10
      [0x00, 0xff, 0x00, 0x00],   # 11
      [0x00, 0x00, 0xff, 0x00],   # 12
      [0xff, 0xff, 0x00, 0x00],   # 13
      [0xff, 0x00, 0xff, 0x00],   # 14
      [0x00, 0xff, 0xff, 0x00],   # 15
      [0x80, 0x00, 0x00, 0x00],   # 16
      [0x00, 0x80, 0x00, 0x00],   # 17
      [0x00, 0x00, 0x80, 0x00],   # 18
      [0x80, 0x80, 0x00, 0x00],   # 19
      [0x80, 0x00, 0x80, 0x00],   # 20
      [0x00, 0x80, 0x80, 0x00],   # 21
      [0xc0, 0xc0, 0xc0, 0x00],   # 22
      [0x80, 0x80, 0x80, 0x00],   # 23
      [0x99, 0x99, 0xff, 0x00],   # 24
      [0x99, 0x33, 0x66, 0x00],   # 25
      [0xff, 0xff, 0xcc, 0x00],   # 26
      [0xcc, 0xff, 0xff, 0x00],   # 27
      [0x66, 0x00, 0x66, 0x00],   # 28
      [0xff, 0x80, 0x80, 0x00],   # 29
      [0x00, 0x66, 0xcc, 0x00],   # 30
      [0xcc, 0xcc, 0xff, 0x00],   # 31
      [0x00, 0x00, 0x80, 0x00],   # 32
      [0xff, 0x00, 0xff, 0x00],   # 33
      [0xff, 0xff, 0x00, 0x00],   # 34
      [0x00, 0xff, 0xff, 0x00],   # 35
      [0x80, 0x00, 0x80, 0x00],   # 36
      [0x80, 0x00, 0x00, 0x00],   # 37
      [0x00, 0x80, 0x80, 0x00],   # 38
      [0x00, 0x00, 0xff, 0x00],   # 39
      [0x00, 0xcc, 0xff, 0x00],   # 40
      [0xcc, 0xff, 0xff, 0x00],   # 41
      [0xcc, 0xff, 0xcc, 0x00],   # 42
      [0xff, 0xff, 0x99, 0x00],   # 43
      [0x99, 0xcc, 0xff, 0x00],   # 44
      [0xff, 0x99, 0xcc, 0x00],   # 45
      [0xcc, 0x99, 0xff, 0x00],   # 46
      [0xff, 0xcc, 0x99, 0x00],   # 47
      [0x33, 0x66, 0xff, 0x00],   # 48
      [0x33, 0xcc, 0xcc, 0x00],   # 49
      [0x99, 0xcc, 0x00, 0x00],   # 50
      [0xff, 0xcc, 0x00, 0x00],   # 51
      [0xff, 0x99, 0x00, 0x00],   # 52
      [0xff, 0x66, 0x00, 0x00],   # 53
      [0x66, 0x66, 0x99, 0x00],   # 54
      [0x96, 0x96, 0x96, 0x00],   # 55
      [0x00, 0x33, 0x66, 0x00],   # 56
      [0x33, 0x99, 0x66, 0x00],   # 57
      [0x00, 0x33, 0x00, 0x00],   # 58
      [0x33, 0x33, 0x00, 0x00],   # 59
      [0x99, 0x33, 0x00, 0x00],   # 60
      [0x99, 0x33, 0x66, 0x00],   # 61
      [0x33, 0x33, 0x99, 0x00],   # 62
      [0x33, 0x33, 0x33, 0x00]    # 63
    ]
    return 0
  end

  ###############################################################################
  #
  # set_tempdir()
  #
  # Change the default temp directory used by _initialize() in Worksheet.pm.
  #
  def set_tempdir(dir = '')
    raise "#{dir} is not a valid directory" if dir != '' && !FileTest.directory?(dir)
    raise "set_tempdir must be called before add_worksheet" unless sheets.empty?

    @tempdir = dir
  end

  ###############################################################################
  #
  # set_codepage()
  #
  # See also the _store_codepage method. This is used to store the code page, i.e.
  # the character set used in the workbook.
  #
  def set_codepage(type = 1)
    if type == 2
      @codepage = 0x8000
    else
      @codepage = 0x04E4
    end
  end

  ###############################################################################
  #
  # set_properties()
  #
  # Set the document properties such as Title, Author etc. These are written to
  # property sets in the OLE container.
  #
  def set_properties(*args)
    # Ignore if no args were passed.
    return -1 unless args.size == 0

    # Allow the parameters to be passed as a hash or hash ref.
    param = args[0].kind_of?(Hash) ? args[0] : Hash[*args]

    # List of valid input parameters.
    properties = {
      :codepage      => [0x0001, 'VT_I2'      ],
      :title         => [0x0002, 'VT_LPSTR'   ],
      :subject       => [0x0003, 'VT_LPSTR'   ],
      :author        => [0x0004, 'VT_LPSTR'   ],
      :keywords      => [0x0005, 'VT_LPSTR'   ],
      :comments      => [0x0006, 'VT_LPSTR'   ],
      :last_author   => [0x0008, 'VT_LPSTR'   ],
      :created       => [0x000C, 'VT_FILETIME'],
      :category      => [0x0002, 'VT_LPSTR'   ],
      :manager       => [0x000E, 'VT_LPSTR'   ],
      :company       => [0x000F, 'VT_LPSTR'   ],
      :utf8          => 1
    }

    # Check for valid input parameters.
    param.each_key do |k|
      unless properties.has_key?(k)
        raise "Unknown parameter '#{k}' in set_properties()";
      end
    end

    # Set the creation time unless specified by the user.
    unless param.has_key(:created)
      param[:created] = @localtime
    end

    #
    # Create the SummaryInformation property set.
    #

    # Get the codepage of the strings in the property set.
    strings = ["title", "subject", "author", "keywords",  "comments", "last_author"]
    param[:codepage] = get_property_set_codepage(param, strings)

    # Create an array of property set values.
    property_sets = []
    strings.unshift("codepage")
    strings.push("created")
    strings.each do |property|
      if param.has_key?(property) && !param[property].nil?
        property_sets.push(
        [ properties[property][0],
          properties[property][1],
        param[property]  ]
        )
      end
    end

    # Pack the property sets.
    @summary = create_summary_property_set(property_sets)

    #
    # Create the DocSummaryInformation property set.
    #

    # Get the codepage of the strings in the property set.
    strings = ["category", "manager", "company"]
    param[:codepage] = get_property_set_codepage(param, strings)

    # Create an array of property set values.
    property_sets = []

    ["codepage", "category", "manager", "company"].each do |property|
      if param.has_key?(property) && !param[property].nil?
        property_sets.push(
        [ properties[property][0],
          properties[property][1],
        param[property]  ]
        )
      end
    end

    # Pack the property sets.
    @doc_summary = create_doc_summary_property_set(property_sets)

    # Set a flag for when the files is written.
    @add_doc_properties = 1
  end

  ###############################################################################
  #
  # _get_property_set_codepage()
  #
  # Get the character codepage used by the strings in a property set. If one of
  # the strings used is utf8 then the codepage is marked as utf8. Otherwise
  # Latin 1 is used (although in our case this is limited to 7bit ASCII).
  #
  def get_property_set_codepage(params, strings)
    # Allow for manually marked utf8 strings.
    unless params[utf8].nil?
      return 0xFDE9
    else
      return 0x04E4; # Default codepage, Latin 1.
    end
  end

  ###############################################################################
  #
  # _store_workbook()
  #
  # Assemble worksheets into a workbook and send the BIFF data to an OLE
  # storage.
  #
  def store_workbook
    # Add a default worksheet if non have been added.
    add_worksheet if @worksheets.size == 0

    # Calculate size required for MSO records and update worksheets.
    calc_mso_sizes

    # Ensure that at least one worksheet has been selected.
    if @activesheet == 0
      @worksheets[0].selected = 1
      @worksheets[0].hidden   = 0
    end

    # Calculate the number of selected worksheet tabs and call the finalization
    # methods for each worksheet
    @worksheets.each do |sheet|
      @selected    += 1 if sheet.selected != 0
      sheet.active  = 1 if sheet.index == @activesheet
    end

    # Add Workbook globals
    store_bof(0x0005)
    store_codepage()
    store_window1()
    store_hideobj()
    store_1904()
    store_all_fonts()
    store_all_num_formats()
    store_all_xfs()
    store_all_styles()
    store_palette()

    # Calculate the offsets required by the BOUNDSHEET records
    calc_sheet_offsets

    # Add BOUNDSHEET records.
    @worksheets.each do |sheet|
      store_boundsheet(sheet.name,
      sheet.offset,
      sheet.type,
      sheet.hidden,
      sheet.encoding)
    end

    # NOTE: If any records are added between here and EOF the
    # _calc_sheet_offsets() should be updated to include the new length.
    store_country()
    if @ext_ref_count != 0
      store_supbook
      store_externsheet
      store_names
    end
    add_mso_drawing_group
    store_shared_strings
    store_extsst

    # End Workbook globals
    store_eof

    # Store the workbook in an OLE container
    return store_OLE_file
  end

  ###############################################################################
  #
  # _store_OLE_file()
  #
  # Store the workbook in an OLE container using the default handler or using
  # OLE::Storage_Lite if the workbook data is > ~ 7MB.
  #
  def store_OLE_file
    maxsize = 7_087_104

    if @add_doc_properties == 0 && @biffsize <= maxsize
      # Write the OLE file using OLEwriter if data <= 7MB
      ole  = OLEWriter.new(@fh_out)

      # Write the BIFF data without the OLE container for testing.
      ole.biff_only = @biff_only

      # Indicate that we created the filehandle and want to close it.
      ole.internal_fh = @internal_fh

      ole.set_size(@biffsize)
      ole.write_header

      while tmp = get_data
        ole.write(tmp)
      end

      @worksheets.each do |worksheet|
        while tmp = worksheet.get_data
          ole.write(tmp)
        end
      end

      return ole.close
=begin
    else
      # Write the OLE file using ruby-ole if data > 7MB
   
      # Create the Workbook stream.
      stream   = 'Workbook'.unpack('C*').pack('v*')
      workbook = OLE::Storage_Lite::PPS::File->newFile($stream);
   
      while (my $tmp = $self->get_data()) {
         $workbook->append($tmp);
      }
   
      foreach my $worksheet (@{$self->{_worksheets}}) {
        while (my $tmp = $worksheet->get_data()) {
              $workbook->append($tmp);
         }
      }
   
      push @streams, $workbook;
   
   
      # Create the properties streams, if any.
      if ($self->{_add_doc_properties}) {
        my $stream;
        my $summary;
   
        $stream  = pack 'v*', unpack 'C*', "\5SummaryInformation";
        $summary = $self->{summary};
        $summary = OLE::Storage_Lite::PPS::File->new($stream, $summary);
        push @streams, $summary;
   
        $stream  = pack 'v*', unpack 'C*', "\5DocumentSummaryInformation";
        $summary = $self->{doc_summary};
        $summary = OLE::Storage_Lite::PPS::File->new($stream, $summary);
        push @streams, $summary;
      }
   
      # Create the OLE root document and add the substreams.
      my @localtime = @{ $self->{_localtime} };
      splice(@localtime, 6);
   
      my $ole_root = OLE::Storage_Lite::PPS::Root->new(\@localtime,
                                                          \@localtime,
                                                          \@streams);
      $ole_root->save($self->{_filename});
   
   
      # Close the filehandle if it was created internally.
      return CORE::close($self->{_fh_out}) if $self->{_internal_fh};
=end
    end
  end

  ###############################################################################
  #
  # _calc_sheet_offsets()
  #
  # Calculate Worksheet BOF offsets records for use in the BOUNDSHEET records.
  #
  def calc_sheet_offsets
    _bof     = 12
    _eof     = 4
    offset  = @datasize

    # Add the length of the COUNTRY record
    offset += 8

    # Add the length of the SST and associated CONTINUEs
    offset += calculate_shared_string_sizes

    # Add the length of the EXTSST record.
    offset += calculate_extsst_size

    # Add the length of the SUPBOOK, EXTERNSHEET and NAME records
    offset += calculate_extern_sizes

    # Add the length of the MSODRAWINGGROUP records including an extra 4 bytes
    # for any CONTINUE headers. See _add_mso_drawing_group_continue().
    mso_size = @mso_size
    mso_size += 4 * Integer((mso_size -1) / Float(@limit))
    offset   += mso_size

    @worksheets.each do |sheet|
      offset += _bof + sheet.name.length
    end

    offset += _eof

    @worksheets.each do |sheet|
      sheet.offset = offset
      sheet.close(*@sheetnames)
      offset += sheet.datasize
    end

    @biffsize = offset
  end

  ###############################################################################
  #
  # _calc_mso_sizes()
  #
  # Calculate the MSODRAWINGGROUP sizes and the indexes of the Worksheet
  # MSODRAWING records.
  #
  # In the following SPID is shape id, according to Escher nomenclature.
  #
  def calc_mso_sizes
    mso_size        = 0    # Size of the MSODRAWINGGROUP record
    start_spid      = 1024 # Initial spid for each sheet
    max_spid        = 1024 # spidMax
    num_clusters    = 1    # cidcl
    shapes_saved    = 0    # cspSaved
    drawings_saved  = 0    # cdgSaved
    clusters        = []

    process_images

    # Add Bstore container size if there are images.
    mso_size += 8 unless @images_data.empty?

    # Iterate through the worksheets, calculate the MSODRAWINGGROUP parameters
    # and space required to store the record and the MSODRAWING parameters
    # required by each worksheet.
    #
    @worksheets.each do |sheet|
      next unless sheet.kind_of?(Worksheet)

      num_images     = sheet.num_images || 0
      image_mso_size = sheet.image_mso_size || 0
      num_comments   = sheet.prepare_comments
      num_charts     = sheet.prepare_charts
      num_filters    = sheet.filter_count

      next unless num_images + num_comments + num_charts + num_filters != 0

      # Include 1 parent MSODRAWING shape, per sheet, in the shape count.
      num_shapes    = 1 + num_images   +  num_comments +
                          num_charts   +  num_filters
      shapes_saved += num_shapes
      mso_size     += image_mso_size

      # Add a drawing object for each sheet with comments.
      drawings_saved += 1

      # For each sheet start the spids at the next 1024 interval.
      max_spid   = 1024 * (1 + Integer((max_spid -1)/1024.0))
      start_spid = max_spid

      # Max spid for each sheet and eventually for the workbook.
      max_spid  += num_shapes

      # Store the cluster ids
      i = num_shapes
      while i > 0
        num_clusters  += 1
        mso_size      += 8
        size           = i > 1024 ? 1024 : i

        clusters.push([drawings_saved, size])
        i -= 1024
      end

      # Pass calculated values back to the worksheet
      sheet.object_ids = [start_spid, drawings_saved,
      num_shapes, max_spid -1]
    end


    # Calculate the MSODRAWINGGROUP size if we have stored some shapes.
    mso_size              += 86 if mso_size != 0 # Smallest size is 86+8=94

    @mso_size      = mso_size
    @mso_clusters  = [
      max_spid, num_clusters, shapes_saved,
      drawings_saved, clusters
    ]
  end

  ###############################################################################
  #
  # _process_images()
  #
  # We need to process each image in each worksheet and extract information.
  # Some of this information is stored and used in the Workbook and some is
  # passed back into each Worksheet. The overall size for the image related
  # BIFF structures in the Workbook is calculated here.
  #
  # MSO size =  8 bytes for bstore_container +
  #            44 bytes for blip_store_entry +
  #            25 bytes for blip
  #          = 77 + image size.
  #
  def process_images
    images_seen     = {}
    image_data      = []
    previous_images = []
    image_id        = 1;
    images_size     = 0;

    @worksheets.each do |sheet|
      next unless sheet.kind_of?(Worksheet)
      next if sheet.prepare_images == 0

      num_images      = 0
      image_mso_size  = 0

      sheet.images_array.each do |image|
        filename = image[2]
        num_images += 1

        #
        # For each Worksheet image we get a structure like this
        # [
        #   $row,
        #   $col,
        #   $name,
        #   $x_offset,
        #   $y_offset,
        #   $scale_x,
        #   $scale_y,
        # ]
        #
        # And we add additional information:
        #
        #   $image_id,
        #   $type,
        #   $width,
        #   $height;

        if images_seen[filename].nil?
          # TODO should also match seen images based on checksum.

          # Open the image file and import the data.
          fh = open(filename, "rb")
          raise "Couldn't import #{filename}: #{$!}" unless fh

          # Slurp the file into a string and do some size calcs.
          #             my $data        = do {local $/; <$fh>};
          size        = data.length
          checksum1   = image_checksum(data, image_id)
          checksum2   = checksum1
          ref_count   = 1

          # Process the image and extract dimensions.
          # Test for PNGs...
          if  data.unpack('x A3') ==  'PNG'
            type, width, height = process_png(data)
            # Test for JFIF and Exif JPEGs...
          elsif ( data.unpack('n') == 0xFFD8 &&
            (data.unpack('x6 A4') == 'JFIF' ||
            data.unpack('x6 A4') == 'Exif')
            )
            type, width, height = process_jpg(data, filename)
            # Test for BMPs...
          elsif data.unpack('A2') == 'BM'
            type, width, height = process_bmp(data, filename)
            # The 14 byte header of the BMP is stripped off.
            data[0, 13] = ''

            # A checksum of the new image data is also required.
            checksum2  = image_checksum(data, image_id, image_id)

            # Adjust size -14 (header) + 16 (extra checksum).
            size += 2
          else
            raise "Unsupported image format for file: #{filename}\n"
          end

          # Push the new data back into the Worksheet array;
          image.push(image_id, type, width, height)

          # Also store new data for use in duplicate images.
          previous_images.push([image_id, type, width, height])

          # Store information required by the Workbook.
          image_data.push([ref_count, type, data, size,
          checksum1, checksum2])

          # Keep track of overall data size.
          images_size       += size +61; # Size for bstore container.
          image_mso_size    += size +69; # Size for dgg container.

          images_seen[filename] = image_id
          image_id += 1
          close(fh)
        else
          # We've processed this file already.
          index = images_seen[filename] -1

          # Increase image reference count.
          image_data[index][0] += 1

          # Add previously calculated data back onto the Worksheet array.
          # $image_id, $type, $width, $height
          a_ref = images_array[index]
          image_ref.push(previous_images[index])
        end
      end

      # Store information required by the Worksheet.
      @num_images     = num_images
      @image_mso_size = image_mso_size

    end


    # Store information required by the Workbook.
    @images_size = images_size
    @images_data = image_data     # Store the data for MSODRAWINGGROUP.

  end

  ###############################################################################
  #
  # _image_checksum()
  #
  # Generate a checksum for the image using whichever module is available..The
  # available modules are checked in _get_checksum_method(). Excel uses an MD4
  # checksum but any other will do. In the event of no checksum module being
  # available we simulate a checksum using the image index.
  #
  def image_checksum(data, index1, index2 = 0)
    if    @checksum_method == 1
      # Digest::MD4
      #           return Digest::MD4::md4_hex($data);
    elsif @checksum_method == 2
      # Digest::Perl::MD4
      #           return Digest::Perl::MD4::md4_hex($data);
    elsif @checksum_method == 3
      # Digest::MD5
      return Digest::MD5.hexdigest(data)
    else
      # Default
      return sprintf('%016X%016X', index2, index1)
    end
  end

  ###############################################################################
  #
  # _process_png()
  #
  # Extract width and height information from a PNG file.
  #
  def process_png(data)
    type    = 6 # Excel Blip type (MSOBLIPTYPE).
    width   = data[16, 4].unpack("N")
    height  = data[20, 4].unpack("N")

    return [type, width, height]
  end

  ###############################################################################
  #
  # _process_bmp()
  #
  # Extract width and height information from a BMP file.
  #
  # Most of these checks came from the old Worksheet::_process_bitmap() method.
  #
  def process_bmp(data, filename)
    type     = 7   # Excel Blip type (MSOBLIPTYPE).

    # Check that the file is big enough to be a bitmap.
    if data.length  <= 0x36
      raise "#{filename} doesn't contain enough data."
    end

    # Read the bitmap width and height. Verify the sizes.
    width, height = data.unpack("x18 V2")

    if width > 0xFFFF
      raise "#{filename}: largest image width #{width} supported is 65k."
    end

    if height > 0xFFFF
      raise "#{filename}: largest image height supported is 65k."
    end

    # Read the bitmap planes and bpp data. Verify them.
    planes, bitcount = data.unpack("x26 v2")

    if bitcount != 24
      raise "#{filename} isn't a 24bit true color bitmap."
    end

    if planes != 1
      raise "#{filename}: only 1 plane supported in bitmap image."
    end

    # Read the bitmap compression. Verify compression.
    compression = data.unpack("x30 V")

    if compression != 0
      raise "#{filename}: compression not supported in bitmap image."
    end

    return [type, width, height]
  end

  ###############################################################################
  #
  # _process_jpg()
  #
  # Extract width and height information from a JPEG file.
  #
  def process_jpg(data, filename)
    type     = 5  # Excel Blip type (MSOBLIPTYPE).

    offset = 2;
    data_length = data.length

    # Search through the image data to find the 0xFFC0 marker. The height and
    # width are contained in the data for that sub element.
    while offset < data_length
      marker  = data[offset,   2].unpack("n")
      marker = marker[0]
      length  = data[offset+2, 2].unpack("n")
      length = length[0]
      
      if marker == 0xFFC0
        height = data[offset+5, 2].unpack("n")
        height = height[0]
        width  = data[offset+7, 2].unpack("n")
        width  = width[0]
        break
      end

      offset = offset + length + 2
      break if marker == 0xFFDA
    end

    if height.nil?
      raise "#{filename}: no size data found in image.\n"
    end

    return [type, width, height]
  end

  ###############################################################################
  #
  # _store_all_fonts()
  #
  # Store the Excel FONT records.
  #
  def store_all_fonts
    format  = @formats[15]   # The default cell format.
    font    = format.get_font

    # Fonts are 0-indexed. According to the SDK there is no index 4,
    (0..3).each do
      append(font)
    end

    # Add the font for comments. This isn't connected to any XF format.
    tmp    = Format.new(nil, :font => 'Tahoma', :size => 8)
    font   = tmp.get_font
    append(font)

    # Iterate through the XF objects and write a FONT record if it isn't the
    # same as the default FONT and if it hasn't already been used.
    #
    fonts = {}
    index = 6                    # The first user defined FONT

    key = format.get_font_key    # The default font for cell formats.
    fonts[key] = 0               # Index of the default font

    # Fonts that are marked as '_font_only' are always stored. These are used
    # mainly for charts and may not have an associated XF record.

    @formats.each do |format|
      key = format.get_font_key

      if format.font_only == 0 and !fonts[key].nil?
        # FONT has already been used
        format.font_index = fonts[key]
      else
        # Add a new FONT record

        if format.font_only == 0
          fonts[key] = index
        end

        format.font_index = index
        index += 1
        font = format.get_font
        append(font)
      end
    end
  end

  ###############################################################################
  #
  # _store_all_num_formats()
  #
  # Store user defined numerical formats i.e. FORMAT records
  #
  def store_all_num_formats
    num_formats = {}
    index = 164       # User defined FORMAT records start from 0xA4

    # Iterate through the XF objects and write a FORMAT record if it isn't a
    # built-in format type and if the FORMAT string hasn't already been used.
    #
    @formats.each do |format|
      num_format = format.num_format
      encoding   = format.num_format_enc

      # Check if $num_format is an index to a built-in format.
      # Also check for a string of zeros, which is a valid format string
      # but would evaluate to zero.
      #
      unless num_format =~ /^0+\d/
        next if num_format =~ /^\d+$/   # built-in
      end

      if num_formats[num_format]
        # FORMAT has already been used
        format.num_format = num_formats[num_format]
      else
        # Add a new FORMAT
        num_formats[num_format] = index
        format.num_format       = index
        store_num_format(num_format, index, encoding)
        index += 1
      end
    end
  end

  ###############################################################################
  #
  # _store_all_xfs()
  #
  # Write all XF records.
  #
  def store_all_xfs
    @formats.each do |format|
      xf = format.get_xf
      append(xf)
    end
  end

  ###############################################################################
  #
  # _store_all_styles()
  #
  # Write all STYLE records.
  #
  def store_all_styles
    # Excel adds the built-in styles in alphabetical order.
    built_ins = [
      [0x03, 16], # Comma
      [0x06, 17], # Comma[0]
      [0x04, 18], # Currency
      [0x07, 19], # Currency[0]
      [0x00,  0], # Normal
      [0x05, 20]  # Percent

      # We don't deal with these styles yet.
      #[0x08, 21], # Hyperlink
      #[0x02,  8], # ColLevel_n
      #[0x01,  1], # RowLevel_n
    ]

    built_ins.each do |aref|
      type     = aref[0]
      xf_index = aref[1]

      store_style(type, xf_index)
    end
  end

  ###############################################################################
  #
  # _store_names()
  #
  # Write the NAME record to define the print area and the repeat rows and cols.
  #
  def store_names
    index       = 0

    # Create the print area NAME records
    @worksheets.each do |worksheet|

      key = "#{index}:#{index}"
      ref = @ext_refs[key]
      index += 1

      # Write a Name record if Autofilter has been defined
      if worksheet.filter_count != 0
        store_name_short(
        worksheet.index,
        0x0D, # NAME type = Filter Database
        ref,
        worksheet.filter_area[0],
        worksheet.filter_area[1],
        worksheet.filter_area[2],
        worksheet.filter_area[3],
        1     # Hidden
        )
      end

      # Write a Name record if the print area has been defined
      if worksheet.print_rowmin
        store_name_short(
        worksheet.index,
        0x06, # NAME type = Print_Area
        ref,
        worksheet.print_rowmin,
        worksheet.print_rowmax,
        worksheet.print_colmin,
        worksheet.print_colmax
        )
      end

    end

    index = 0

    # Create the print title NAME records
    @worksheets.each do |worksheet|

      rowmin = worksheet.title_rowmin
      rowmax = worksheet.title_rowmax
      colmin = worksheet.title_colmin
      colmax = worksheet.title_colmax
      key = "#{index}:#{index}"
      ref = @ext_refs[key]
      index += 1

      # Determine if row + col, row, col or nothing has been defined
      # and write the appropriate record
      #
      if rowmin && colmin
        # Row and column titles have been defined.
        # Row title has been defined.
        store_name_long(
        worksheet.index,
        0x07, # NAME type = Print_Titles
        ref,
        rowmin,
        rowmax,
        colmin,
        colmax
        )
      elsif rowmin
        # Row title has been defined.
        store_name_short(
        worksheet.index,
        0x07, # NAME type = Print_Titles
        ref,
        rowmin,
        rowmax,
        0x00,
        0xff
        )
      elsif colmin
        # Column title has been defined.
        store_name_short(
        worksheet.index,
        0x07, # NAME type = Print_Titles
        ref,
        0x0000,
        0xffff,
        colmin,
        colmax
        )
      else
        # Nothing left to do
      end
    end
  end

  ###############################################################################
  ###############################################################################
  #
  # BIFF RECORDS
  #


  ###############################################################################
  #
  # _store_window1()
  #
  # Write Excel BIFF WINDOW1 record.
  #
  def store_window1
    record    = 0x003D                 # Record identifier
    length    = 0x0012                 # Number of bytes to follow

    xWn       = 0x0000                 # Horizontal position of window
    yWn       = 0x0000                 # Vertical position of window
    dxWn      = 0x355C                 # Width of window
    dyWn      = 0x30ED                 # Height of window

    grbit     = 0x0038                 # Option flags
    ctabsel   = @selected              # Number of workbook tabs selected
    wTabRatio = 0x0258                 # Tab to scrollbar ratio

    itabFirst = @firstsheet            # 1st displayed worksheet
    itabCur   = @activesheet           # Active worksheet

    header    = [record, length].pack("vv")
    data      = [xWn, yWn, dxWn, dyWn,
      grbit, itabCur, itabFirst,
    ctabsel, wTabRatio].pack("vvvvvvvvv")

    append(header, data)
  end

  ###############################################################################
  #
  # _store_boundsheet()
  #    my $sheetname = $_[0];                # Worksheet name
  #    my $offset    = $_[1];                # Location of worksheet BOF
  #    my $type      = $_[2];                # Worksheet type
  #    my $hidden    = $_[3];                # Worksheet hidden flag
  #    my $encoding  = $_[4];                # Sheet name encoding
  #
  # Writes Excel BIFF BOUNDSHEET record.
  #
  def store_boundsheet(sheetname, offset, type, hidden, encoding)
    record    = 0x0085                    # Record identifier
    length    = 0x08 + sheetname.length   # Number of bytes to follow

    cch       = sheetname.length          # Length of sheet name

    grbit     = type | hidden

    # Character length is num of chars not num of bytes
    cch /= 2 if encoding

    # Change the UTF-16 name from BE to LE
    sheetname = [sheetname].unpack('v*').pack('n*') if encoding != 0

    header    = [record, length].pack("vv")
    data      = [offset, grbit, cch, encoding].pack("VvCC")

    append(header, data, sheetname)
  end

  ###############################################################################
  #
  # _store_style()
  #    type      = $_[0]  # Built-in style
  #    xf_index  = $_[1]  # Index to style XF
  #
  # Write Excel BIFF STYLE records.
  #
  def store_style(type, xf_index)
    record    = 0x0293    # Record identifier
    length    = 0x0004    # Bytes to follow

    level     = 0xff      # Outline style level

    xf_index    |= 0x8000 # Add flag to indicate built-in style.

    header    = [record, length].pack("vv")
    data      = [xf_index, type, level].pack("vCC")

    append(header, data)
  end

  ###############################################################################
  #
  # _store_num_format()
  #    my $format    = $_[0];          # Custom format string
  #    my $ifmt      = $_[1];          # Format index code
  #    my $encoding  = $_[2];          # Char encoding for format string
  #
  # Writes Excel FORMAT record for non "built-in" numerical formats.
  #
  def store_num_format(format, ifmt, encoding)
    format = format.to_s unless format.kind_of?(String)
    record    = 0x041E         # Record identifier
    # length                   # Number of bytes to follow

    # Char length of format string
    cch = format.length


    # Handle Unicode format strings.
    if encoding == 1
      raise "Uneven number of bytes in Unicode font name" if cch % 2 != 0
      cch /= 2 if encoding != 0
      format  = [format].unpack('n*').pack('v*')
    end

    # Special case to handle Euro symbol, 0x80, in non-Unicode strings.
    if encoding == 0 and format =~ /\x80/
      format   =  [format].unpack('C*').pack('v*')
      format.gsub!(/\x80\x00/, "\xAC\x20")
      encoding =  1
    end

    length    = 0x05 + format.length

    header    = [record, length].pack("vv")
    data      = [ifmt, cch, encoding].pack("vvC")

    append(header, data, format)
  end

  ###############################################################################
  #
  # _store_1904()
  #
  # Write Excel 1904 record to indicate the date system in use.
  #
  def store_1904
    record    = 0x0022         # Record identifier
    length    = 0x0002         # Bytes to follow

    f1904     = @date_1904 ? 1 : 0     # Flag for 1904 date system

    header    = [record, length].pack("vv")
    data      = [f1904].pack("v")

    append(header, data)
  end

  ###############################################################################
  #
  # _store_supbook()
  #
  # Write BIFF record SUPBOOK to indicate that the workbook contains external
  # references, in our case, formula, print area and print title refs.
  #
  def store_supbook
    record      = 0x01AE                   # Record identifier
    length      = 0x0004                   # Number of bytes to follow

    ctabs       = @worksheets.size         # Number of worksheets
    stVirtPath  = 0x0401                   # Encoded workbook filename

    header    = [record, length].pack("vv")
    data      = [ctabs, stVirtPath].pack("vv")

    append(header, data)
  end

  ###############################################################################
  #
  # _store_externsheet()
  #
  # Writes the Excel BIFF EXTERNSHEET record. These references are used by
  # formulas. TODO NAME record is required to define the print area and the
  # repeat rows and columns.
  #
  def store_externsheet
    record      = 0x0017                   # Record identifier

    # Get the external refs
    ext_refs = @ext_refs
    ext = sort ext_refs.keys

    # Change the external refs from stringified "1:1" to [1, 1]
    ext.each do |e|
      e = e.split(/:/)
    end

    cxti        = @ext.size                # Number of Excel XTI structures
    rgxti       = ''                       # Array of XTI structures

    # Write the XTI structs
    ext.each do |e|
      rgxti = rgxti + [0, e[0], e[1]].pack("vvv")
    end

    data        = [cxti].pack("v") + rgxti
    header    = [record, data.length].pack("vv")

    append(header, data)
  end

  ###############################################################################
  #
  # _store_name_short()
  #    index           = shift        # Sheet index
  #    type            = shift
  #    ext_ref         = shift        # TODO
  #    rowmin          = $_[0]        # Start row
  #    rowmax          = $_[1]        # End row
  #    colmin          = $_[2]        # Start column
  #    colmax          = $_[3]        # end column
  #    hidden          = $_[4]        # Name is hidden
  #
  #
  # Store the NAME record in the short format that is used for storing the print
  # area, repeat rows only and repeat columns only.
  #
  def store_name_short(index, type, ext_ref, rowmin, rowmax, colmin, colmax)
    record          = 0x0018       # Record identifier
    length          = 0x001b       # Number of bytes to follow

    index           = shift        # Sheet index
    type            = shift
    ext_ref         = shift        # TODO

    grbit           = 0x0020       # Option flags
    chKey           = 0x00         # Keyboard shortcut
    cch             = 0x01         # Length of text name
    cce             = 0x000b       # Length of text definition
    unknown01       = 0x0000       #
    ixals           = index +1     # Sheet index
    unknown02       = 0x00         #
    cchCustMenu     = 0x00         # Length of cust menu text
    cchDescription  = 0x00         # Length of description text
    cchHelptopic    = 0x00         # Length of help topic text
    cchStatustext   = 0x00         # Length of status bar text
    rgch            = type         # Built-in name type
    unknown03       = 0x3b         #

    grbit           = 0x0021 if hidden

    header          = [record, length].pack("vv")
    data            = [grbit].pack("v")
    data            = data + [chKey].pack("C")
    data            = data + [cch].pack("C")
    data            = data + [cce].pack("v")
    data            = data + [unknown01].pack("v")
    data            = data + [ixals].pack("v")
    data            = data + [unknown02].pack("C")
    data            = data + [cchCustMenu].pack("C")
    data            = data + [cchDescription].pack("C")
    data            = data + [cchHelptopic].pack("C")
    data            = data + [cchStatustext].pack("C")
    data            = data + [rgch].pack("C")
    data            = data + [unknown03].pack("C")
    data            = data + [ext_ref].pack("v")

    data            = data + [rowmin].pack("v")
    data            = data + [rowmax].pack("v")
    data            = data + [colmin].pack("v")
    data            = data + [colmax].pack("v")

    append(header, data)
  end

  ###############################################################################
  #
  # _store_name_long()
  #    my $index           = shift;        # Sheet index
  #    my $type            = shift;
  #    my $ext_ref         = shift;        # TODO
  #    my $rowmin          = $_[0];        # Start row
  #    my $rowmax          = $_[1];        # End row
  #    my $colmin          = $_[2];        # Start column
  #    my $colmax          = $_[3];        # end column
  #
  #
  # Store the NAME record in the long format that is used for storing the repeat
  # rows and columns when both are specified. This share a lot of code with
  # _store_name_short() but we use a separate method to keep the code clean.
  # Code abstraction for reuse can be carried too far, and I should know. ;-)
  #
  def store_name_long(index, type, ext_ref, rowmin, rowmax, colmin, colmax)
    record          = 0x0018       # Record identifier
    length          = 0x002a       # Number of bytes to follow

    index           = shift        # Sheet index
    type            = shift
    ext_ref         = shift        # TODO

    grbit           = 0x0020       # Option flags
    chKey           = 0x00         # Keyboard shortcut
    cch             = 0x01         # Length of text name
    cce             = 0x001a       # Length of text definition
    unknown01       = 0x0000       #
    ixals           = index +1     # Sheet index
    unknown02       = 0x00         #
    cchCustMenu     = 0x00         # Length of cust menu text
    cchDescription  = 0x00         # Length of description text
    cchHelptopic    = 0x00         # Length of help topic text
    cchStatustext   = 0x00         # Length of status bar text
    rgch            = type         # Built-in name type

    unknown03       = 0x29
    unknown04       = 0x0017
    unknown05       = 0x3b

    header          = [record, length].pack("vv")
    data            = [grbit].pack("v")
    data            = data + [chKey].pack("C")
    data            = data + [cch].pack("C")
    data            = data + [cce].pack("v")
    data            = data + [unknown01].pack("v")
    data            = data + [ixals].pack("v")
    data            = data + [unknown02].pack("C")
    data            = data + [cchCustMenu].pack("C")
    data            = data + [cchDescription].pack("C")
    data            = data + [cchHelptopic].pack("C")
    data            = data + [cchStatustext].pack("C")
    data            = data + [rgch].pack("C")

    # Column definition
    data            = data + [unknown03].pack("C")
    data            = data + [unknown04].pack("v")
    data            = data + [unknown05].pack("C")
    data            = data + [ext_ref].pack("v")
    data            = data + [0x0000].pack("v")
    data            = data + [0xffff].pack("v")
    data            = data + [colmin].pack("v")
    data            = data + [colmax].pack("v")

    # Row definition
    data            = data + [unknown05].pack("C")
    data            = data + [ext_ref].pack("v")
    data            = data + [rowmin].pack("v")
    data            = data + [rowmax].pack("v")
    data            = data + [0x00].pack("v")
    data            = data + [0xff].pack("v")
    # End of data
    data            = data + [0x10].pack("C")

    append(header, data)
  end

  ###############################################################################
  #
  # _store_palette()
  #
  # Stores the PALETTE biff record.
  #
  def store_palette
    record          = 0x0092                 # Record identifier
    length          = 2 + 4 * @palette.size  # Number of bytes to follow
    ccv             =         @palette.size  # Number of RGB values to follow
    data            = ''                     # The RGB data

    # Pack the RGB data
    @palette.each do |p|
      data = data + p.pack('CCCC')
    end
    
    header = [record, length, ccv].pack("vvv")

    append(header, data)
  end

  ###############################################################################
  #
  # _store_codepage()
  #
  # Stores the CODEPAGE biff record.
  #
  def store_codepage
    record          = 0x0042               # Record identifier
    length          = 0x0002               # Number of bytes to follow
    cv              = @codepage            # The code page

    header          = [record, length].pack("vv")
    data            = [cv].pack("v")

    append(header, data)
  end

  ###############################################################################
  #
  # _store_country()
  #
  # Stores the COUNTRY biff record.
  #
  # Will add setter method for the country codes when/if required.
  #
  def store_country
    record          = 0x008C               # Record identifier
    length          = 0x0004               # Number of bytes to follow
    country_default = 1
    country_win_ini = 1

    header          = [record, length].pack("vv")
    data            = [country_default, country_win_ini].pack("vv")

    append(header, data)
  end

  ###############################################################################
  #
  # _store_hideobj()
  #
  # Stores the HIDEOBJ biff record.
  #
  def store_hideobj
    record          = 0x008D               # Record identifier
    length          = 0x0002               # Number of bytes to follow
    hide            = @hideobj             # Option to hide objects

    header          = [record, length].pack("vv")
    data            = [hide].pack("v")

    append(header, data)
  end

  ###############################################################################
  ###############################################################################
  ###############################################################################



  ###############################################################################
  #
  # _calculate_extern_sizes()
  #
  # We need to calculate the space required by the SUPBOOK, EXTERNSHEET and NAME
  # records so that it can be added to the BOUNDSHEET offsets.
  #
  def calculate_extern_sizes
    ext_refs        = @parser.get_ext_sheets
    ext_ref_count   = ext_refs.keys.size
    length          = 0
    index           = 0

    @worksheets.each do |worksheet|

      rowmin      = worksheet.title_rowmin
      colmin      = worksheet.title_colmin
      filter      = worksheet.filter_count
      key         = "#{index}:#{index}"
      index += 1

      # Add area NAME records
      #
      if worksheet.print_rowmin
        if ext_ref[key].nil?
          ext_refs[key] = ext_ref_count
          ext_ref_count += 1
        end
        length += 31
      end

      # Add title  NAME records
      #
      if rowmin and colmin
        if ext_ref[key].nil?
          ext_refs[key] = ext_ref_count
          ext_ref_count += 1
        end

        length += 46
      elsif rowmin or colmin
        if ext_ref[key].nil?
          ext_refs[key] = ext_ref_count
          ext_ref_count += 1
        end
        length += 31
      else
        # TODO, may need this later.
      end

      # Add Autofilter  NAME records
      #
      if filter != 0
        if ext_ref[key].nil?
          ext_refs[key] = ext_ref_count
          ext_ref_count += 1
        end
        length += 31
      end
    end

    # Update the ref counts.
    @ext_ref_count = ext_ref_count
    @ext_refs      = ext_refs

    # If there are no external refs then we don't write, SUPBOOK, EXTERNSHEET
    # and NAME. Therefore the length is 0.

    return length = 0 if ext_ref_count == 0

    # The SUPBOOK record is 8 bytes
    length += 8

    # The EXTERNSHEET record is 6 bytes + 6 bytes for each external ref
    length += 6 * (1 + ext_ref_count)

    return length
  end

  ###############################################################################
  #
  # _calculate_shared_string_sizes()
  #
  # Handling of the SST continue blocks is complicated by the need to include an
  # additional continuation byte depending on whether the string is split between
  # blocks or whether it starts at the beginning of the block. (There are also
  # additional complications that will arise later when/if Rich Strings are
  # supported). As such we cannot use the simple CONTINUE mechanism provided by
  # the _add_continue() method in BIFFwriter.pm. Thus we have to make two passes
  # through the strings data. The first is to calculate the required block sizes
  # and the second, in _store_shared_strings(), is to write the actual strings.
  # The first pass through the data is also used to calculate the size of the SST
  # and CONTINUE records for use in setting the BOUNDSHEET record offsets. The
  # downside of this is that the same algorithm repeated in _store_shared_strings.
  #
  def calculate_shared_string_sizes
    strings = []
    #strings = self->{_str_unique} -1 # Pre-extend array

    @str_table.each_key do |key|
      strings[@str_table[key]] = key
    end
    # The SST data could be very large, free some memory (maybe).
    @str_table = nil
    @str_array = strings

    # Iterate through the strings to calculate the CONTINUE block sizes.
    #
    # The SST blocks requires a specialised CONTINUE block, so we have to
    # ensure that the maximum data block size is less than the limit used by
    # _add_continue() in BIFFwriter.pm. For simplicity we use the same size
    # for the SST and CONTINUE records:
    #   8228 : Maximum Excel97 block size
    #     -4 : Length of block header
    #     -8 : Length of additional SST header information
    #     -8 : Arbitrary number to keep within _add_continue() limit
    # = 8208
    #
    continue_limit = 8208
    block_length   = 0
    written        = 0
    block_sizes    = []
    continue       = 0

    strings.each do |string|

      string_length = string.length
      encoding      = [string].unpack("xx C")
      split_string  = 0

      # Block length is the total length of the strings that will be
      # written out in a single SST or CONTINUE block.
      #
      block_length += string_length

      # We can write the string if it doesn't cross a CONTINUE boundary
      if block_length < continue_limit
        written += string_length
        next
      end


      # Deal with the cases where the next string to be written will exceed
      # the CONTINUE boundary. If the string is very long it may need to be
      # written in more than one CONTINUE record.
      #
      while block_length >= continue_limit

        # We need to avoid the case where a string is continued in the first
        # n bytes that contain the string header information.
        #
        header_length   = 3 # Min string + header size -1
        space_remaining = continue_limit -written -continue


        # Unicode data should only be split on char (2 byte) boundaries.
        # Therefore, in some cases we need to reduce the amount of available
        # space by 1 byte to ensure the correct alignment.
        align = 0

        # Only applies to Unicode strings
        if encoding == 1
          # Min string + header size -1
          header_length = 4

          if space_remaining > header_length
            # String contains 3 byte header => split on odd boundary
            if split_string == 0 and space_remaining % 2 != 1
              space_remaining -= 1
              align = 1
              # Split section without header => split on even boundary
            elsif split_string != 0 and space_remaining % 2 == 1
              space_remaining -= 1
              align = 1
            end

            split_string = 1
          end
        end

        if space_remaining > header_length
          # Write as much as possible of the string in the current block
          written      += space_remaining

          # Reduce the current block length by the amount written
          block_length -= continue_limit -continue -align

          # Store the max size for this block
          block_sizes.push(continue_limit -align)

          # If the current string was split then the next CONTINUE block
          # should have the string continue flag (grbit) set unless the
          # split string fits exactly into the remaining space.
          #
          if block_length > 0
            continue = 1
          else
            continue = 0
          end
        else
          # Store the max size for this block
          block_sizes.push(written +continue)

          # Not enough space to start the string in the current block
          block_length -= continue_limit -space_remaining -continue
          continue = 0
        end

        # If the string (or substr) is small enough we can write it in the
        # new CONTINUE block. Else, go through the loop again to write it in
        # one or more CONTINUE blocks
        #
        if block_length < continue_limit
          written = block_length
        else
          written = 0
        end
      end
    end

    # Store the max size for the last block unless it is empty
    block_sizes.push(written +continue) if written +continue != 0

    @str_block_sizes = block_sizes

    # Calculate the total length of the SST and associated CONTINUEs (if any).
    # The SST record will have a length even if it contains no strings.
    # This length is required to set the offsets in the BOUNDSHEET records since
    # they must be written before the SST records
    #
    length  = 12
    length +=     block_sizes.shift unless block_sizes.empty? # SST
    while !block_sizes.empty? do
      length += 4 + block_sizes.shift                         # CONTINUEs
    end

    return length
  end

  ###############################################################################
  #
  # _store_shared_strings()
  #
  # Write all of the workbooks strings into an indexed array.
  #
  # See the comments in _calculate_shared_string_sizes() for more information.
  #
  # We also use this routine to record the offsets required by the EXTSST table.
  # In order to do this we first identify the first string in an EXTSST bucket
  # and then store its global and local offset within the SST table. The offset
  # occurs wherever the start of the bucket string is written out via append().
  #
  def store_shared_strings
    strings = @str_array

    record              = 0x00FC   # Record identifier
    length              = 0x0008   # Number of bytes to follow
    total               = 0x0000

    # Iterate through the strings to calculate the CONTINUE block sizes
    continue_limit = 8208
    block_length   = 0
    written        = 0
    continue       = 0

    # The SST and CONTINUE block sizes have been pre-calculated by
    # _calculate_shared_string_sizes()
    block_sizes    = @str_block_sizes

    # The SST record is required even if it contains no strings. Thus we will
    # always have a length
    #
    if block_sizes.size != 0
      length = 8 + block_sizes.shift
    else
      # No strings
      length = 8
    end

    # Initialise variables used to track EXTSST bucket offsets.
    extsst_str_num  = -1
    sst_block_start = @datasize

    # Write the SST block header information
    header      = [record, length].pack("vv")
    data        = [@str_total, @str_unique].pack("VV")
    append(header, data)

    # Iterate through the strings and write them out
    return if strings.empty?
    strings.each do |string|

      string_length = string.length
      encoding      = [string].unpack("xx C")
      split_string  = 0
      bucket_string = 0 # Used to track EXTSST bucket offsets.

      # Check if the string is at the start of a EXTSST bucket.
      extsst_str_num += 1
      if extsst_str_num % @extsst_bucket_size == 0
        bucket_string = 1
      end

      # Block length is the total length of the strings that will be
      # written out in a single SST or CONTINUE block.
      #
      block_length += string_length

      # We can write the string if it doesn't cross a CONTINUE boundary
      if block_length < continue_limit

        # Store location of EXTSST bucket string.
        if bucket_string != 0
          global_offset   = @datasize
          local_offset    = @datasize - sst_block_start

          @extsst_offsets.push([global_offset, local_offset])
          bucket_string = 0
        end

        append(string)
        written += string_length
        next
      end

      # Deal with the cases where the next string to be written will exceed
      # the CONTINUE boundary. If the string is very long it may need to be
      # written in more than one CONTINUE record.
      #
      while block_length >= continue_limit

        # We need to avoid the case where a string is continued in the first
        # n bytes that contain the string header information.
        #
        header_length   = 3 # Min string + header size -1
        space_remaining = continue_limit -written -continue


        # Unicode data should only be split on char (2 byte) boundaries.
        # Therefore, in some cases we need to reduce the amount of available
        # space by 1 byte to ensure the correct alignment.
        align = 0

        # Only applies to Unicode strings
        if encoding == 1
          # Min string + header size -1
          header_length = 4

          if space_remaining > header_length
            # String contains 3 byte header => split on odd boundary
            if split_string == 0 and space_remaining % 2 != 1
              space_remaining -= 1
              align = 1
              # Split section without header => split on even boundary
            elsif split_string != 0 and space_remaining % 2 == 1
              space_remaining -= 1
              align = 1
            end

            split_string = 1
          end
        end

        if space_remaining > header_length
          # Write as much as possible of the string in the current block
          tmp = string[0, space_remaining]

          # Store location of EXTSST bucket string.
          if bucket_string != 0
            global_offset   = @datasize
            local_offset    = @datasize - sst_block_start

            @extsst_offsets.push([global_offset, local_offset])
            bucket_string = 0
          end

          append(tmp)


          # The remainder will be written in the next block(s)
          string = string[space_remaining .. string.length-1]

          # Reduce the current block length by the amount written
          block_length -= continue_limit -continue -align

          # If the current string was split then the next CONTINUE block
          # should have the string continue flag (grbit) set unless the
          # split string fits exactly into the remaining space.
          #
          if block_length > 0
            continue = 1
          else
            continue = 0
          end
        else
          # Not enough space to start the string in the current block
          block_length -= continue_limit -space_remaining -continue
          continue = 0
        end

        # Write the CONTINUE block header
        if block_sizes.size != 0
          sst_block_start= @datasize # Reset EXTSST offset.

          record         = 0x003C
          length         = block_sizes.shift

          header         = [record, length].pack("vv")
          header         = header + [encoding].pack("C") if continue != 0

          append(header)
        end

        # If the string (or substr) is small enough we can write it in the
        # new CONTINUE block. Else, go through the loop again to write it in
        # one or more CONTINUE blocks
        #
        if block_length < continue_limit

          # Store location of EXTSST bucket string.
          if bucket_string != 0
            global_offset   = @datasize
            local_offset    = @datasize - sst_block_start

            @extsst_offsets.push([global_offset, local_offset])

            bucket_string = 0
          end
          append(string)

          written = block_length
        else
          written = 0
        end
      end
    end
  end

  ###############################################################################
  #
  # _calculate_extsst_size
  #
  # The number of buckets used in the EXTSST is between 0 and 128. The number of
  # strings per bucket (bucket size) has a minimum value of 8 and a theoretical
  # maximum of 2^16. For "number of strings" < 1024 there is a constant bucket
  # size of 8. The following algorithm generates the same size/bucket ratio
  # as Excel.
  #
  def calculate_extsst_size
    unique_strings  = @str_unique

    if unique_strings < 1024
      bucket_size = 8
    else
      bucket_size = 1 + Integer(unique_strings / 128.0)
    end

    buckets = Integer((unique_strings + bucket_size -1)  / Float(bucket_size))

    @extsst_buckets        = buckets
    @extsst_bucket_size    = bucket_size

    return 6 + 8 * buckets
  end

  ###############################################################################
  #
  # _store_extsst
  #
  # Write EXTSST table using the offsets calculated in _store_shared_strings().
  #
  def store_extsst
    offsets     = @extsst_offsets
    bucket_size = @extsst_bucket_size

    record      = 0x00FF                 # Record identifier
    length      = 2 + 8 * offsets.size   # Bytes to follow

    header      = [record, length].pack('vv')
    data        = [bucket_size].pack('v')

    offsets.each do |offset|
      data = data + [offset[0], offset[1], 0].pack('Vvv')
    end

    append(header, data)

  end

  #
  # Methods related to comments and MSO objects.
  #

  ###############################################################################
  #
  # _add_mso_drawing_group()
  #
  # Write the MSODRAWINGGROUP record that keeps track of the Escher drawing
  # objects in the file such as images, comments and filters.
  #
  def add_mso_drawing_group
    return unless @mso_size != 0

    record  = 0x00EB               # Record identifier
    length  = 0x0000               # Number of bytes to follow

    data    = store_mso_dgg_container
    data    = data + store_mso_dgg(*@mso_clusters)
    data    = data + store_mso_bstore_container
    @images_data.each do |image|
      data = data + store_mso_images(image)
    end
    data    = data + store_mso_opt
    data    = data + store_mso_split_menu_colors

    length  = data.length
    header  = [record, length].pack("vv")

    add_mso_drawing_group_continue(header + data)

    return header + data # For testing only.
  end

  ###############################################################################
  #
  # _add_mso_drawing_group_continue()
  #
  # See first the Spreadsheet::WriteExcel::BIFFwriter::_add_continue() method.
  #
  # Add specialised CONTINUE headers to large MSODRAWINGGROUP data block.
  # We use the Excel 97 max block size of 8228 - 4 bytes for the header = 8224.
  #
  # The structure depends on the size of the data block:
  #
  #     Case 1:  <=   8224 bytes      1 MSODRAWINGGROUP
  #     Case 2:  <= 2*8224 bytes      1 MSODRAWINGGROUP + 1 CONTINUE
  #     Case 3:  >  2*8224 bytes      2 MSODRAWINGGROUP + n CONTINUE
  #
  def add_mso_drawing_group_continue(data)
    limit       = 8228 -4
    mso_group   = 0x00EB # Record identifier
    continue    = 0x003C # Record identifier
    block_count = 1

    # Ignore the base class _add_continue() method.
    @ignore_continue = 1

    # Case 1 above. Just return the data as it is.
    if data.length <= limit
      append(data)
      return
    end

    # Change length field of the first MSODRAWINGGROUP block. Case 2 and 3.
    tmp = data.dup
    tmp[0, limit + 4] = ""
    tmp[2, 2] = [limit].pack('v')
    append(tmp)

    # Add MSODRAWINGGROUP and CONTINUE blocks for Case 3 above.
    while data.length > limit
      if block_count == 1
        # Add extra MSODRAWINGGROUP block header.
        header = [mso_group, limit].pack("vv")
        block_count += 1
      else
        # Add normal CONTINUE header.
        header = [continue, limit].pack("vv")
      end

      tmp = data.dup
      tmp[0, limit] = ''
      append(header, tmp)
    end

    # Last CONTINUE block for remaining data. Case 2 and 3 above.
    header = [continue, data.length].pack("vv")
    append(header, data)

    # Turn the base class _add_continue() method back on.
    @ignore_continue = 0
  end

  ###############################################################################
  #
  # _store_mso_dgg_container()
  #
  # Write the Escher DggContainer record that is part of MSODRAWINGGROUP.
  #
  def store_mso_dgg_container
    type        = 0xF000
    version     = 15
    instance    = 0
    data        = ''
    length      = @mso_size -12 # -4 (biff header) -8 (for this).

    return add_mso_generic(type, version, instance, data, length)
  end


  ###############################################################################
  #
  # _store_mso_dgg()
  #    my $max_spid        = $_[0];
  #    my $num_clusters    = $_[1];
  #    my $shapes_saved    = $_[2];
  #    my $drawings_saved  = $_[3];
  #    my $clusters        = $_[4];
  #
  # Write the Escher Dgg record that is part of MSODRAWINGGROUP.
  #
  def store_mso_dgg(max_spid, num_clusters, shapes_saved, drawings_saved, clusters)
    type            = 0xF006
    version         = 0
    instance        = 0
    data            = ''
    length          = nil  # Calculate automatically.

    data            = [max_spid, num_clusters,
    shapes_saved, drawings_saved].pack("VVVV")

    clusters.each do |aref|
      drawing_id      = aref[0]
      shape_ids_used  = aref[1]

      data            = data + [drawing_id, shape_ids_used].pack("VV")
    end

    return add_mso_generic(type, version, instance, data, length)
  end

  ###############################################################################
  #
  # _store_mso_bstore_container()
  #
  # Write the Escher BstoreContainer record that is part of MSODRAWINGGROUP.
  #
  def store_mso_bstore_container
    return '' if @images_size == 0

    type        = 0xF001
    version     = 15
    instance    = @images_data.size          # Number of images.
    data        = ''
    length      = @images_size +8 *instance

    return add_mso_generic(type, version, instance, data, length)
  end

  ###############################################################################
  #
  # _store_mso_images()
  #    ref_count   = $_[0]
  #    image_type  = $_[1]
  #    image       = $_[2]
  #    size        = $_[3]
  #    checksum1   = $_[4]
  #    checksum2   = $_[5]
  #
  # Write the Escher BstoreContainer record that is part of MSODRAWINGGROUP.
  #
  def store_mso_images(ref_count, image_type, image, size, checksum1, checksum2)
    blip_store_entry =  store_mso_blip_store_entry(
    ref_count,
    image_type,
    size,
    checksum1)

    blip             =  store_mso_blip(
    image_type,
    image,
    size,
    checksum1,
    checksum2)

    return blip_store_entry + blip
  end

  ###############################################################################
  #
  # _store_mso_blip_store_entry()
  #    ref_count   = $_[0]
  #    image_type  = $_[1]
  #    size        = $_[2]
  #    checksum1   = $_[3]
  #
  # Write the Escher BlipStoreEntry record that is part of MSODRAWINGGROUP.
  #
  def store_mso_blip_store_entry(ref_count, image_type, size, checksum1)
    type        = 0xF007
    version     = 2
    instance    = image_type
    length      = size +61
    data        = [image_type].pack('C')  +    # Win32
    [image_type].pack('C')  +    # Mac
    [checksum1].pack('H*')  +    # Uid checksum
    [0xFF].pack('v')        +    # Tag
    [size +25].pack('V')    +    # Next Blip size
    [ref_count].pack('V')   +    # Image ref count
    [0x00000000].pack('V')  +    # File offset
    [0x00].pack('C')        +    # Usage
    [0x00].pack('C')        +    # Name length
    [0x00].pack('C')        +    # Unused
    [0x00].pack('C')             # Unused

    return add_mso_generic(type, version, instance, data, length)
  end

  ###############################################################################
  #
  # _store_mso_blip()
  #    image_type  = $_[0]
  #    image_data  = $_[1]
  #    size        = $_[2]
  #    checksum1   = $_[3]
  #    checksum2   = $_[4]
  #
  # Write the Escher Blip record that is part of MSODRAWINGGROUP.
  #
  def store_mso_blip(image_type, image_data, size, checksum1, checksum2)
    instance = 0x046A if image_type == 5 # JPG
    instance = 0x06E0 if image_type == 6 # PNG
    instance = 0x07A9 if image_type == 7 # BMP

    # BMPs contain an extra checksum for the stripped data.
    if image_type == 7
      checksum1 = checksum2 + checksum1
    end

    type        = 0xF018 + image_type
    version     = 0x0000
    length      = size +17
    data        = [checksum1].pack('H*')  +     # Uid checksum
    [0xFF].pack('C')        +     # Tag
    image_data                   # Image

    return add_mso_generic(type, version, instance, data, length)
  end

  ###############################################################################
  #
  # _store_mso_opt()
  #
  # Write the Escher Opt record that is part of MSODRAWINGGROUP.
  #
  def store_mso_opt
    type        = 0xF00B
    version     = 3
    instance    = 3
    data        = ''
    length      = 18

    data        = ['BF0008000800810109000008C0014000'+'0008'].pack("H*")

    return add_mso_generic(type, version, instance, data, length)
  end

  ###############################################################################
  #
  # _store_mso_split_menu_colors()
  #
  # Write the Escher SplitMenuColors record that is part of MSODRAWINGGROUP.
  #
  def store_mso_split_menu_colors
    type        = 0xF11E
    version     = 0
    instance    = 4
    data        = ''
    length      = 16

    data        = ['0D0000080C00000817000008F7000010'].pack("H*")

    return add_mso_generic(type, version, instance, data, length)
  end

end

=begin
= Notes on the difference between Workbook.pm and workbook.rb
---deprecated methods
   I generally elminated any deprecated methods.  That means no 'write'
   methods.
---date_system
   This is the 1904 attribute.  However, since a number can't be a method,
   this doesn't work very well for attribute_accessor.  Besides, date_system
   is more descriptive.
=end
