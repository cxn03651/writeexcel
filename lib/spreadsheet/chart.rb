###############################################################################
#
# Chart - A writer class for Excel Charts.
#
#
# Used in conjunction with Spreadsheet::WriteExcel
#
# Copyright 2000-2008, John McNamara, jmcnamara@cpan.org
#

require 'biffwriter'

class Chart


  attr_reader :name


  ###############################################################################
  #
  # new()
  #
  # Constructor. Creates a new Chart object from a BIFFwriter object
  #
  def initialize(filename, name, index, encoding, activesheet, firstsheet)
    @filename          = filename
    @name              = name
    @index             = index
    @encoding          = encoding
    @activesheet       = activesheet
    @firstsheet        = firstsheet

    @type              = 0x0200
    @ext_sheets        = []
    @using_tmpfile     = 1
    @filehandle        = ""
    @fileclosed        = false
    @offset            = 0
    @xls_rowmax        = 0
    @xls_colmax        = 0
    @xls_strmax        = 0
    @dim_rowmin        = 0
    @dim_rowmax        = 0
    @dim_colmin        = 0
    @dim_colmax        = 0
    @dim_changed       = 0
    @colinfo           = []
    @selection         = [0, 0]
    @panes             = []
    @active_pane       = 3
    @frozen            = 0
    @selected          = 0
    @hidden            = 0

    @paper_size        = 0x0
    @orientation       = 0x1
    @header            = ''
    @footer            = ''
    @hcenter           = 0
    @vcenter           = 0
    @margin_head       = 0.50
    @margin_foot       = 0.50
    @margin_left       = 0.75
    @margin_right      = 0.75
    @margin_top        = 1.00
    @margin_bottom     = 1.00

    @title_rowmin      = nil
    @title_rowmax      = nil
    @title_colmin      = nil
    @title_colmax      = nil
    @print_rowmin      = nil
    @print_rowmax      = nil
    @print_colmin      = nil
    @print_colmax      = nil

    @print_gridlines   = 1
    @screen_gridlines  = 1
    @print_headers     = 0

    @fit_page          = 0
    @fit_width         = 0
    @fit_height        = 0

    @hbreaks           = []
    @vbreaks           = []

    @protect           = 0
    @password          = nil

    @col_sizes         = {}
    @row_sizes         = {}

    @col_formats       = {}
    @row_formats       = {}

    @zoom              = 100
    @print_scale       = 100

    @leading_zeros     = 0

    @outline_row_level = 0
    @outline_style     = 0
    @outline_below     = 1
    @outline_right     = 1
    @outline_on        = 1

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
    buffer = 4096

    return tmp if read(@filehandle, tmp, buffer)

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
    @hidd        = 1

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
    @datasize   = File.Stat.size(@filename)
  end

  ###############################################################################
  #
  # _close()
  #
  # Add data to the beginning of the workbook (note the reverse order)
  # and to the end of the workbook.
  #
  def _close
  end

end
