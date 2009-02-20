class BIFFWriter

   BIFF_Version = 0x0500
   BigEndian    = [1].pack("I") == [1].pack("N")

   attr_reader :byte_order, :data, :datasize

   ######################################################################
   # The args here aren't used by BIFFWriter, but they are needed by its 
   # subclasses.  I don't feel like creating multiple constructors.
   ######################################################################
   def initialize(*args)
      @data       = ""
      @datasize   = 0
   end

   def prepend(*args)
      @data = args.join << @data
      @datasize += args.join.length
   end

   def append(*args)
      @data << args.join
      @datasize += args.join.length
   end

   def store_bof(type = 0x0005)
      record  = 0x0809
      length  = 0x0008
      build   = 0x096C
      year    = 0x07C9

      header  = [record,length].pack("vv")
      data    = [BIFF_Version,type,build,year].pack("vvvv") 

      prepend(header, data)
   end

   def store_eof
      record = 0x000A
      length = 0x0000
      header = [record,length].pack("vv")
 
      append(header)
   end
end   
