module Writeexcel

class Worksheet < BIFFWriter
  class CellRange
    attr_accessor :row_min, :row_max, :col_min, :col_max

    def initialize(worksheet)
      @worksheet = worksheet
    end

    def increment_row_max
      @row_max += 1 if @row_max
    end

    def increment_col_max
      @col_max += 1 if @col_max
    end

    def row(val)
      @row_min = val if !@row_min || (val < row_min)
      @row_max = val if !@row_max || (val > row_max)
    end

    def col(val)
      @col_min = val if !@col_min || (val < col_min)
      @col_max = val if !@col_max || (val > col_max)
    end

    #
    # assemble the NAME record in the long format that is used for storing the repeat
    # rows and columns when both are specified. This share a lot of code with
    # name_record_short() but we use a separate method to keep the code clean.
    # Code abstraction for reuse can be carried too far, and I should know. ;-)
    #
    #    type
    #    ext_ref           # TODO
    #
    def name_record_long(type, ext_ref)       #:nodoc:
      record          = 0x0018       # Record identifier
      length          = 0x002a       # Number of bytes to follow

      grbit           = 0x0020       # Option flags
      chkey           = 0x00         # Keyboard shortcut
      cch             = 0x01         # Length of text name
      cce             = 0x001a       # Length of text definition
      unknown01       = 0x0000       #
      ixals           = @worksheet.index + 1    # Sheet index
      unknown02       = 0x00         #
      cch_cust_menu   = 0x00         # Length of cust menu text
      cch_description = 0x00         # Length of description text
      cch_helptopic   = 0x00         # Length of help topic text
      cch_statustext  = 0x00         # Length of status bar text
      rgch            = type         # Built-in name type

      unknown03       = 0x29
      unknown04       = 0x0017
      unknown05       = 0x3b

      header          = [record, length].pack("vv")
      data            = [grbit].pack("v")
      data           += [chkey].pack("C")
      data           += [cch].pack("C")
      data           += [cce].pack("v")
      data           += [unknown01].pack("v")
      data           += [ixals].pack("v")
      data           += [unknown02].pack("C")
      data           += [cch_cust_menu].pack("C")
      data           += [cch_description].pack("C")
      data           += [cch_helptopic].pack("C")
      data           += [cch_statustext].pack("C")
      data           += [rgch].pack("C")

      # Column definition
      data           += [unknown03].pack("C")
      data           += [unknown04].pack("v")
      data           += [unknown05].pack("C")
      data           += [ext_ref].pack("v")
      data           += [0x0000].pack("v")
      data           += [0xffff].pack("v")
      data           += [@col_min].pack("v")
      data           += [@col_max].pack("v")

      # Row definition
      data           += [unknown05].pack("C")
      data           += [ext_ref].pack("v")
      data           += [@row_min].pack("v")
      data           += [@row_max].pack("v")
      data           += [0x00].pack("v")
      data           += [0xff].pack("v")
      # End of data
      data           += [0x10].pack("C")

      [header, data]
    end

    #
    # assemble the NAME record in the short format that is used for storing the print
    # area, repeat rows only and repeat columns only.
    #
    #    type
    #    ext_ref          # TODO
    #    hidden           # Name is hidden
    #
    def name_record_short(type, ext_ref, hidden = nil)       #:nodoc:
      record          = 0x0018       # Record identifier
      length          = 0x001b       # Number of bytes to follow

      grbit           = 0x0020       # Option flags
      chkey           = 0x00         # Keyboard shortcut
      cch             = 0x01         # Length of text name
      cce             = 0x000b       # Length of text definition
      unknown01       = 0x0000       #
      ixals           = @worksheet.index + 1    # Sheet index
      unknown02       = 0x00         #
      cch_cust_menu   = 0x00         # Length of cust menu text
      cch_description = 0x00         # Length of description text
      cch_helptopic   = 0x00         # Length of help topic text
      cch_statustext  = 0x00         # Length of status bar text
      rgch            = type         # Built-in name type
      unknown03       = 0x3b         #

      grbit           = 0x0021 if hidden

      rowmin = row_min
      rowmax = row_max
      rowmin, rowmax = 0x0000, 0xffff unless row_min

      colmin = col_min
      colmax = col_max
      colmin, colmax = 0x00, 0xff unless col_min

      header          = [record, length].pack("vv")
      data            = [grbit].pack("v")
      data           += [chkey].pack("C")
      data           += [cch].pack("C")
      data           += [cce].pack("v")
      data           += [unknown01].pack("v")
      data           += [ixals].pack("v")
      data           += [unknown02].pack("C")
      data           += [cch_cust_menu].pack("C")
      data           += [cch_description].pack("C")
      data           += [cch_helptopic].pack("C")
      data           += [cch_statustext].pack("C")
      data           += [rgch].pack("C")
      data           += [unknown03].pack("C")
      data           += [ext_ref].pack("v")

      data           += [rowmin].pack("v")
      data           += [rowmax].pack("v")
      data           += [colmin].pack("v")
      data           += [colmax].pack("v")

      [header, data]
    end
  end

  class CellDimension < CellRange
    def row_min
      @row_min || 0
    end

    def col_min
      @col_min || 0
    end

    def row_max
      @row_max || 0
    end

    def col_max
      @col_max || 0
    end
  end

  class PrintRange < CellRange
    def name_record_short(ext_ref, hidden)
      super(0x06, ext_ref, hidden) # 0x06  NAME type = Print_Area
    end
  end

  class TitleRange < CellRange
    def name_record_long(ext_ref)
      super(0x07, ext_ref) # 0x07  NAME type = Print_Titles
    end

    def name_record_short(ext_ref, hidden)
      super(0x07, ext_ref, hidden) # 0x07  NAME type = Print_Titles
    end
  end

  class FilterRange < CellRange
    def name_record_short(ext_ref, hidden)
      super(0x0D, ext_ref, hidden) # 0x0D  NAME type = Filter Database
    end

    def count
      if @col_min && @col_max
        1 + @col_max - @col_min
      else
        0
      end
    end

    def inside?(col)
      col < @col_min || col > @col_max
    end
  end
end

end
