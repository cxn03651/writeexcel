#
# BIFFwriter - An abstract base class for Excel workbooks and worksheets.
#
#
# Used in conjunction with Spreadsheet::WriteExcel
#
# Copyright 2000-2008, John McNamara, jmcnamara@cpan.org
#
class BIFFWriter

   BIFF_Version = 0x0600
   byte_order = ''

#   BigEndian    = [1].pack("I") == [1].pack("N")

   attr_reader :byte_order, :data, :datasize

   ######################################################################
   # The args here aren't used by BIFFWriter, but they are needed by its 
   # subclasses.  I don't feel like creating multiple constructors.
   ######################################################################

  def initialize(*args)
    @byte_order      = ''
    @data            = ''
    @datasize        = 0
    @limit           = 8224
    @ignore_continue = 0
  # set_byte_order
  end

  def prepend(*args)
    d = args.join
    if d.length > @limit
      d = add_continue d
    end

    @datasize += d.length
    @data      = d << @data

    return d
   end

  def append(*args)
    d = args.join
    if d.length > @limit
      d = add_continue d
    end

    @datasize += d.length
    @data      = @data << d

    return d
  end

# _store_bof($type)
#
# $type = 0x0005, Workbook
# $type = 0x0010, Worksheet
#
# Writes Excel BOF record to indicate the beginning of a stream or
# sub-stream in the BIFF file.
  def store_bof(type = 0x0005)
    record  = 0x0809      # Record identifier
    length  = 0x0010      # Number of bytes to follow

    build   = 0x0DBB
    year    = 0x07CC

    bfh     = 0x00000041
    sfo     = 0x0000000;

    header  = [record,length].pack("vv")
    data    = [BIFF_Version,type,build,year,bfh,sfo].pack("vvvvVV") 

    prepend(header, data)
  end

# _store_eof()
# 
# Writes Excel EOF record to indicate the end of a BIFF stream.
  def store_eof
    record = 0x000A
    length = 0x0000
    header = [record,length].pack("vv")
 
    append(header)
  end
end   
