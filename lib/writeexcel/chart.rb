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

class Chart < Worksheet

  ###############################################################################
  #
  # new()
  #
  # Constructor. Creates a new Chart object from a BIFFwriter object
  #
  def initialize(workbook, filename, name, index, encoding, activesheet, firstsheet)
    super(workbook, name, index, encoding)

    @filename          = filename
    @name              = name
    @index             = index
    @encoding          = encoding
    @activesheet       = activesheet
    @firstsheet        = firstsheet

    @type              = 0x0200
    @using_tmpfile     = 1
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
  # get_data().
  #
  # Retrieves data from memory in one chunk, or from disk in $buffer
  # sized chunks.
  #
  def get_data
    length = 4096

    @filehandle.read(length)
  end


  ###############################################################################
  #
  # select()
  #
  # Set this worksheet as a selected worksheet, i.e. the worksheet has its tab
  # highlighted.
  #
  def select
    @hidden         = 0 # Selected worksheet can't be hidden.
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
    @hidden      = 0 # Active worksheet can't be hidden.
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
    @hidden      = 1

    # A hidden worksheet shouldn't be active or selected.
    @selecte     = 0
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
    hidden      = 0 # Active worksheet can't be hidden.
    firstsheet  = index
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

  private

  ###############################################################################


  ###############################################################################
  #
  # _initialize()
  #
  def _initialize
    filehandle = open(@filename, "rb") or
    die "Couldn't open #{@filename} in add_chart_ext(): $!.\n"
    @filehandle = filehandle
    @datasize   = FileTest.size(@filename)
  end

end
