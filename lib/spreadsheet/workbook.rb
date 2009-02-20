class Workbook < BIFFWriter
   BOF = 11
   EOF = 4
   SheetName = "Sheet"

   attr_accessor :date_system
   attr_reader :formats, :xf_index, :worksheets

   def initialize(file)
      super

      @file       = file
      @format     = Format.new

      @active_sheet = 0
      @first_sheet  = 0
      @biffsize     = 0
      @date_system  = 1900
      @xf_index     = 16

      @worksheets = []
      @formats    = []

      @url_format = add_format(:color=>"blue", :underline=>1)
   end

   def close
      store_workbook
   end

   def add_format(*args)
      if args[0].kind_of?(Hash)
         f = Format.new(args[0], @xf_index)
      elsif args[0].nil?
         f = Format.new
      else
         raise TypeError unless args[0].kind_of?(Format)
         f = args[0]
         f.xf_index = @xf_index
      end
      @xf_index += 1
      @formats.push(f)
      return f
   end

   def add_worksheet(name=nil)
      index = @worksheets.length

      if name.nil?
         name = SheetName + (index + 1).to_s
      end
      
      args = [name,index, @active_sheet, @first_sheet, @url_format]
      ws = Worksheet.new(*args)
      @worksheets[index] = ws
      return ws
   end

   def calc_sheet_offsets
      offset = @datasize
      @worksheets.each{ |sheet|
         offset += BOF + sheet.name.length
      }

      offset += EOF

      @worksheets.each{ |sheet|
         sheet.offset = offset
         offset += sheet.datasize
      }

      @biffsize = offset
   end

   def store_workbook
      @worksheets.each{ |sheet|
         sheet.close
      }

      store_bof(0x0005)
      store_window1
      store_date_system
      store_all_fonts
      store_all_num_formats
      store_all_xfs
      store_all_styles
      calc_sheet_offsets

      @worksheets.each{ |sheet|
         store_boundsheet(sheet.name, sheet.offset)
      }

      store_eof
      store_ole_file
   end
   
   def store_ole_file
      OLEWriter.open(@file){ |ole|
         ole.set_size(@biffsize)
         ole.write_header
         ole.print(@data)
         @worksheets.each{ |sheet|
            ole.print(sheet.data)
         }
      }
   end

   def store_window1
      record    = 0x003D
      length    = 0x0012

      xWn       = 0x0000
      yWn       = 0x0000
      dxWn      = 0x25BC
      dyWn      = 0x1572

      grbit     = 0x0038
      ctabsel   = 0x0001
      wTabRatio = 0x0258

      itabFirst = @first_sheet
      itabCur   = @active_sheet

      header = [record,length].pack("vv")
      fields = [xWn,yWn,dxWn,dyWn,grbit,itabCur,itabFirst,ctabsel,wTabRatio]
      data   = fields.pack("vvvvvvvvv")

      append(header,data)
   end

   def store_all_fonts
      font = @format.font_biff
      for n in 1..5
         append(font)
      end

      fonts = Hash.new(0)
      index = 6
      key = @format.font_key
      fonts[key] = 0

      @formats.each{ |format|
         key = format.font_key
         if fonts.has_key?(key)
            format.font_index = fonts[key]
         else
            fonts[key] = index
            format.font_index = index
            index += 1
            append(format.font_biff)
         end
      }
   end

   def store_xf(style)
      name   = 0x00E0
      length = 0x0010

      ifnt      = 0x0000
      ifmt      = 0x0000
      align     = 0x0020
      icv       = 0x20C0
      fill      = 0x0000
      brd_line  = 0x0000
      brd_color = 0x0000

      header = [name, length].pack("vv")
      fields = [ifnt,ifmt,style,align,icv,fill,brd_line,brd_color]
      data   = fields.pack("vvvvvvvv")

      append(header, data);
   end

   def store_all_num_formats
      index = 164

      num_formats_hash = {}
      num_formats_array = []

      @formats.each{ |format|
         num_format = format.num_format
         next if num_format.kind_of?(Numeric)
         if num_formats_hash.has_key?(num_format)
            format.num_format = num_formats_hash[num_format]
         else
            num_formats_hash[num_format] = index
            format.num_format = index
            num_formats_array.push(num_format)
            index += 1
         end
      }

      index = 164
      num_formats_array.each{ |num_format|
         store_num_format(num_format,index)
         index += 1
      }
   end

   def store_all_xfs
      xf = @format.xf_biff(0xFFF5)
      for n in 1..15
         append(xf)
      end

      xf = @format.xf_biff(0x0001)
      append(xf)
         
      @formats.each{ |format|
         xf = format.xf_biff(0x0001)
         append(xf)
      }
   end

   def store_style
      record = 0x0293
      length = 0x0004

      ixfe    = 0x8000
      builtin = 0x00
      iLevel  = 0xff

      header = [record, length].pack("vv")
      data   = [ixfe, builtin, iLevel].pack("vCC")

      append(header, data)
   end

   alias store_all_styles store_style

   def store_boundsheet(sheet_name, offset)
      name   = 0x0085
      length = 0x07 + sheet_name.length

      grbit = 0x0000
      cch   = sheet_name.length

      header = [name, length].pack("vv")
      data   = [offset, grbit, cch].pack("VvC")

      append(header, data, sheet_name)
   end

   def store_num_format(format, ifmt)
      record = 0x041E
      cch    = format.length
      length = 0x03 + cch

      header = [record, length].pack("vv")
      data   = [ifmt, cch].pack("vC")

      append(header, data, format)
   end

   def store_date_system
      record = 0x0022
      length = 0x0002
      
      f1904 = 0
      f1904 = 1 if @date_system == 1904

      header = [record, length].pack("vv")
      data   = [f1904].pack("v")

      append(header, data)
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
