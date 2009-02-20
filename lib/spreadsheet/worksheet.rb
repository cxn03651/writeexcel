class MaxSizeError < StandardError; end

class Worksheet < BIFFWriter

   RowMax = 65536
   ColMax = 256
   StrMax = 255
   Buffer = 4096

   attr_reader :name, :xf_index
   attr_accessor :index, :colinfo, :selection, :offset

   def initialize(name, index=0, active_sheet=0, first_sheet=0, url_format=nil)
      super(name,index)

      @name         = name
      @index        = index
      @active_sheet = active_sheet
      @first_sheet  = first_sheet
      @url_format   = url_format

      @offset = 0

      @dim_rowmin = RowMax + 1
      @dim_rowmax = 0
      @dim_colmin = ColMax + 1
      @dim_colmax = 0

      @colinfo = []
      @selection = [0,0]
      @row_formats = {}
      @column_formats = {}
      @outline_row_level = 0
   end

   def close
      store_dimensions

      unless @colinfo.empty?
         while @colinfo.length > 0
            store_colinfo(*@colinfo.pop)
         end
         store_defcol
      end

      store_guts
      store_bof(0x0010)

      store_window2
      store_selection(*@selection)
      store_eof
   end

   def data
      if @data && @data.length > 0
         tmp = @data
         @data = ""
         return tmp
      end
   end

   def activate
      @active_sheet = @index
   end

   def set_first_sheet
      @first_sheet = @index
   end

   def column(*range)
      @colinfo = *range
   end

   def store_dimensions
      record   = 0x0000
      length   = 0x000A
      reserved = 0x0000

      header = [record, length].pack("vv")
      fields = [@dim_rowmin, @dim_rowmax, @dim_colmin, @dim_colmax, reserved]
      data   = fields.pack("vvvvv")

      prepend(header, data)
   end

   def store_guts
      record   = 0x0080 # record identifier
      length   = 0x0008 # bytes to follow
      
      dxRwGut  = 0x0000 # size of row gutter
      dxColGut = 0x0000 # size of col gutter
      
      row_level = @outline_row_level
      col_level = 0
 
      @colinfo.each do |colinfo|
         next if colinfo.length < 6
         col_level = colinfo[5] if colinfo[5] > col_level
      end
      
      col_level = 0 if col_level < 0
      col_level = 7 if col_level > 7
      
      row_level += 1 if row_level > 0
      col_level += 1 if col_level > 0
      
      header = [record, length].pack("vv")
      fields = [dxRwGut, dxColGut, row_level, col_level]
      data   = fields.pack("vvvv")
 
      prepend(header, data)
   end

   def store_window2
      record  = 0x023E
      length  = 0x000A

      grbit   = 0x00B6
      rwTop   = 0x0000
      colLeft = 0x0000
      rgbHdr  = 0x00000000

      if @active_sheet == @index
        grbit = 0x06B6
      end

      header = [record, length].pack("vv")
      data   = [grbit, rwTop, colLeft, rgbHdr].pack("vvvV")

      append(header, data)
   end

   def store_defcol
      record   = 0x0055
      length   = 0x0002

      colwidth = 0x0008

      header = [record, length].pack("vv")
      data   = [colwidth].pack("v")

      prepend(header, data)
   end

   def store_colinfo(first=0, last=0, coldx=8.43, ixfe=0x0F, grbit=0)
      record   = 0x007D
      length   = 0x000B

      coldx += 0.72
      coldx *= 256
      reserved = 0x00

      if ixfe.kind_of?(Format)
         ixfe = ixfe.xf_index
      end

      header = [record, length].pack("vv")
      data   = [first, last, coldx, ixfe, grbit, reserved].pack("vvvvvC")

      prepend(header,data)
   end

   # I think this may have problems
   def store_selection(row=0, col=0, row_last=0, col_last=0)
      record = 0x001D
      length = 0x000F

      pnn     = 3
      rwAct   = row
      colAct  = col
      irefAct = 0
      cref    = 1

      rwFirst  = row
      colFirst = col

      if row_last != 0
         rwLast = row_last
      else
         rwLast = rwFirst
      end

      if col_last != 0
         colLast = col_last
      else
         colLast = colFirst
      end

      if rwFirst > rwLast
         rwFirst,rwLast = rwLast,rwFirst
      end

      if colFirst > colLast
         colFirst,colLast = colLast,colFirst
      end

      header = [record, length].pack("vv")
      fields = [pnn,rwAct,colAct,irefAct,cref,rwFirst,rwLast,colFirst,colLast]
      data   = fields.pack("CvvvvvvCC")

      append(header, data)
   end

   def write(row, col, data=nil, format=nil)
      if data.nil?
         write_blank(row, col, format)
      elsif data.kind_of?(Array)
         write_row(row, col, data, format)
      elsif data.kind_of?(Numeric)
         write_number(row, col, data, format)
      else
         write_string(row, col, data, format)
      end
   end

   def write_row(row, col, tokens=nil, format=nil)
      if tokens.nil?
         write(row,col,tokens,format)
         return
      end

      tokens.each{ |token|
         if token.kind_of?(Array)
            write_column(row,col,token,format)
         else
            write(row,col,token,format)
         end
         col += 1
      }
   end

   def write_column(row, col, tokens=nil, format=nil)
      if tokens.nil?
         write(row,col,tokens,format)
         return
      end

      tokens.each{ |token|
         if token.kind_of?(Array)
            write_row(row, col, token, format)
         else
            write(row, col, token, format)
         end
         row += 1
      }
   end

   def write_number(row, col, num, format=nil)
      record  = 0x0203
      length  = 0x000E
      
      xf_index = XF(row,col,format)

      raise MaxSizeError if row >= RowMax
      raise MaxSizeError if col >= ColMax

      @dim_rowmin = row if row < @dim_rowmin
      @dim_rowmax = row if row > @dim_rowmax
      @dim_colmin = col if col < @dim_colmin
      @dim_colmax = col if col > @dim_colmax

      header    = [record,length].pack("vv")
      data      = [row,col,xf_index].pack("vvv")
      xl_double = [num].pack("d")

      if BigEndian
         xl_double.reverse!
      end

      append(header,data,xl_double)
   end

   def write_string(row, col, str, format)
      record = 0x0204
      length = 0x0008 + str.length

      xf_index = XF(row, col, format)

      strlen = str.length

      raise MaxSizeError if row >= RowMax
      raise MaxSizeError if col >= ColMax

      @dim_rowmin = row if row < @dim_rowmin
      @dim_rowmax = row if row > @dim_rowmax
      @dim_colmin = col if col < @dim_colmin
      @dim_colmax = col if col > @dim_colmax

      # Truncate strings over 255 characters
      if strlen > StrMax
         str    = str[0..StrMax-1]
         length = 0x0008 + StrMax
         strlen = StrMax
      end

      header = [record, length].pack("vv")
      data   = [row, col, xf_index, strlen].pack("vvvv")

      append(header, data, str)
   end

   # Write a blank cell to the specified row and column (zero indexed).
   # A blank cell is used to specify formatting without adding data.
   def write_blank(row, col, format)
      record = 0x0201
      length = 0x0006

      xf_index = XF(row, col, format)

      raise MaxSizeError if row >= RowMax
      raise MaxSizeError if col >= ColMax

      @dim_rowmin = row if row < @dim_rowmin
      @dim_rowmax = row if row > @dim_rowmax
      @dim_colmin = col if col < @dim_colmin
      @dim_colmax = col if col > @dim_colmax

      header = [record, length].pack("vv")
      data   = [row, col, xf_index].pack("vvv")

      append(header, data)
   end

   # Format a rectangular section of cells.  Note that you should call this
   # method only after you have written data to it, or the formatting will
   # be lost.
   def format_rectangle(x1, y1, x2, y2, format)
      raise TypeError, "invalid format" unless format.kind_of?(Format)

      x1.upto(x2){ |row|
         y1.upto(y2){ |col|
            write_blank(row, col, format)
         }
      }
   end

   def write_url(row, col, url, string=url, format=nil)
      record = 0x01B8
      length = 0x0034 + 2 * (1+url.length)
      
      write_string(row,col,string,format)

      header = [record, length].pack("vv")
      data   = [row, row, col, col].pack("vvvv")

      unknown = "D0C9EA79F9BACE118C8200AA004BA90B02000000"
      unknown += "03000000E0C9EA79F9BACE118C8200AA004BA90B"

      stream = [unknown].pack("H*")

      url = url.split('').join("\0")
      url += "\0\0\0"

      len = url.length
      url_len = [len].pack("V")

      append(header + data)
      append(stream)
      append(url_len)
      append(url)
   end

   def format_row(row, height=nil, format=nil, hidden=false, level=0)
      unless row.kind_of?(Range) || row.kind_of?(Integer)
         raise TypeError, 'row must be an Integer or Range'
      end
      
      if hidden.nil? || hidden == false
         hidden = 0
      end

      record = 0x0208 # record identifier
      length = 0x0010 # number of bytes to follow

      col_first = 0x0000 # first defined column
      col_last  = 0x0000 # last defined column
      irwmac    = 0x0000 # used by Excel to optimize loading
      reserved  = 0x0000 # reservered
      grbit     = 0x0000 # option flags
      xf_index  = nil

      grbit |= level
      if hidden > 0
         grbit |= 0x0020
      end
      grbit |= 0x0040 #unsynched
      grbit |= 0x0100

      @outline_row_level = level if level > @outline_row_level

      if format.nil?
         xf_index = 0x0F
      else
         xf_index = format.xf_index
         grbit |= 0x80
      end

      if height.nil?
         height = 0xff
      else
         height = height * 20
      end

      row = row..row if row.kind_of?(Fixnum) # Avoid an if..else clause :)
      row.each{ |r|
         header = [record, length].pack("vv")
         fields = [r,col_first,col_last,height,irwmac,reserved,grbit,xf_index]
         data = fields.pack("vvvvvvvv")

         @row_formats[r] = format if format
         append(header + data)
      }
   end

   # private - adapted from .37 of Spreadsheet::WriteExcel
   def XF(row, col, xf=nil)
      if xf.kind_of?(Format)
         return xf.xf_index
      elsif @row_formats.has_key?(row) 
         return @row_formats[row].xf_index
      elsif @column_formats.has_key?(col)
         return @column_formats[col].xf_index
      else
         return 0x0F
      end
   end
   
   def format_column(column, width=nil, format=nil)
      unless column.kind_of?(Range) || column.kind_of?(Fixnum)
         raise TypeError
      end

      width = 8.43 if width.nil?
      column = column..column if column.kind_of?(Fixnum)

      column.each{ |e|
         @column_formats[e] = format unless format.nil?
      }
      @colinfo.push([column.begin, column.end, width, format])
   end

end
=begin
= Differences between Worksheet.pm and worksheet.rb
---write_url
   I made this a public method to be called directly by the user if they want
   to write a url string.  A variable number of arguments made it a pain to
   integrate into the 'write' method.
=end
