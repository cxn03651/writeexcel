# olewriter.rb
#
# This class should never be instantiated directly. The entire class, and all
# its methods should be considered private.

class MaxSizeError < StandardError; end

class OLEWriter

   # Not meant for public consumption
   MaxSize    = 7087104 # Use Spreadsheet::WriteExcel::Big to exceed this
   BlockSize  = 4096
   BlockDiv   = 512
   ListBlocks = 127

   attr_reader :biff_size, :book_size, :big_blocks, :list_blocks
   attr_reader :root_start, :size_allowed

   # Accept an IO or IO-like object or a filename (as a String)
   def initialize(arg)
      if arg.kind_of?(String)
        @io = File.open(arg, "w")
      else
        @io = arg
      end
      @io.binmode

      @biff_only     = false
      @size_allowed  = true
      @biff_size     = 0
      @book_size     = 0
      @big_blocks    = 0
      @list_blocks   = 0
      @root_start    = 0
      @block_count   = 4
   end

   # Imitate IO.open behavior
   def self.open(arg)
     if block_given?
       ole = self.new(arg)
       result = yield(ole)
       ole.close
       result
     else
       self.new(arg)
     end
   end

   # Delegate 'write' and 'print' to the internal IO object.
   def write(*args, &block)
     @io.write(*args, &block)
   end
   def print(*args, &block)
     @io.print(*args, &block)
   end

   # Set the size of the data to be written to the OLE stream
   #
   # @big_blocks = (109 depot block x (128 -1 marker word)
   #               - (1 x end words)) = 13842
   # MaxSize = @big_blocks * 512 bytes = 7087104
   #
   def set_size(size = BlockSize)
      if size > MaxSize
         return @size_allowed = false
      end

      @biff_size = size

      if biff_size > BlockSize
         @book_size = size
      else
         @book_size = BlockSize
      end

      @size_allowed = true
   end

   # Calculate various sizes needed for the OLE stream
   def calculate_sizes
      @big_blocks  = (@book_size.to_f/BlockDiv.to_f).ceil
      @list_blocks = (@big_blocks / ListBlocks) + 1
      @root_start  = @big_blocks
   end

   # Write root entry, big block list and close the filehandle.
   def close
      if @size_allowed == true
         write_padding
         write_property_storage
         write_big_block_depot
      end
      @io.close
   end 

   # Write the OLE header block
   def write_header
      return if @biff_only == true
      calculate_sizes
      root_start = @root_start
      num_lists  = @list_blocks

      id              = [0xD0CF11E0, 0xA1B11AE1].pack("NN")
      unknown1        = [0x00, 0x00, 0x00, 0x00].pack("VVVV")
      unknown2        = [0x3E, 0x03].pack("vv")
      unknown3        = [-2].pack("v")
      unknown4        = [0x09].pack("v")
      unknown5        = [0x06, 0x00, 0x00].pack("VVV")
      num_bbd_blocks  = [num_lists].pack("V")
      root_startblock = [root_start].pack("V")
      unknown6        = [0x00, 0x1000].pack("VV")
      sbd_startblock  = [-2].pack("V")
      unknown7        = [0x00, -2 ,0x00].pack("VVV")

      write(id)
      write(unknown1)
      write(unknown2)
      write(unknown3)
      write(unknown4)
      write(unknown5)
      write(num_bbd_blocks)
      write(root_startblock)
      write(unknown6)
      write(sbd_startblock)
      write(unknown7)

      unused = [-1].pack("V")

      1.upto(num_lists){
         root_start += 1
         write([root_start].pack("V"))
      }

      num_lists.upto(108){
         write(unused)
      }
   end

   # Write a big block depot
   def write_big_block_depot
      num_blocks   = @big_blocks
      num_lists    = @list_blocks
      total_blocks = num_lists * 128
      used_blocks  = num_blocks + num_lists + 2
      
      marker          = [-3].pack("V")
      end_of_chain    = [-2].pack("V")
      unused          = [-1].pack("V")

      1.upto(num_blocks-1){|n|
         write([n].pack("V"))
      }

      write end_of_chain
      write end_of_chain

      1.upto(num_lists){ write(marker) }

      used_blocks.upto(total_blocks){ write(unused) }
      
   end

   # Write property storage
   def write_property_storage
   
      #########  name         type  dir start size
      write_pps('Root Entry', 0x05,  1,   -2, 0x00)
      write_pps('Book',       0x02, -1, 0x00, @book_size)
      write_pps("",           0x00, -1, 0x00, 0x0000)
      write_pps("",           0x00, -1, 0x00, 0x0000)
   end

   # Write property sheet in property storage
   def write_pps(name, type, dir, start, size)
      length = 0
      ord_name = []
      unless name.empty? 
         name += "\0"
         ord_name = name.unpack("c*")
         length = name.length * 2
      end

      rawname        = ord_name.pack("v*")
      zero           = [0].pack("C")
      
      pps_sizeofname = [length].pack("v")   #0x40
      pps_type       = [type].pack("v")     #0x42
      pps_prev       = [-1].pack("V")       #0x44
      pps_next       = [-1].pack("V")       #0x48
      pps_dir        = [dir].pack("V")      #0x4c
      
      unknown = [0].pack("V")
      
      pps_ts1s       = [0].pack("V")        #0x64
      pps_ts1d       = [0].pack("V")        #0x68
      pps_ts2s       = [0].pack("V")        #0x6c
      pps_ts2d       = [0].pack("V")        #0x70
      pps_sb         = [start].pack("V")    #0x74
      pps_size       = [size].pack("V")     #0x78

      write(rawname)
      for n in 1..64-length
         write(zero)
      end
      write(pps_sizeofname)
      write(pps_type)
      write(pps_prev)
      write(pps_next)
      write(pps_dir)
      for n in 1..5
         write(unknown)
      end
      write(pps_ts1s)
      write(pps_ts1d)
      write(pps_ts2s)
      write(pps_ts2d)
      write(pps_sb)
      write(pps_size)
      write(unknown)
   end

   # Pad the end of the file
   def write_padding
      min_size = 512
      min_size = BlockSize if @biff_size < BlockSize

      if @biff_size % min_size != 0
         padding = min_size - (@biff_size % min_size)
         write("\0" * padding)
      end
   end
end
