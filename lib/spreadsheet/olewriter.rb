# olewriter.rb
#
# This class should never be instantiated directly. The entire class, and all
# its methods should be considered private.

class MaxSizeError < StandardError; end

class OLEWriter

   # Not meant for public consumption
   MaxSize    = 7087104
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
   # - (1 x end words)) = 13842
   #
   # MaxSize = @big_blocks * 512 bytes = 7087104
   def set_size(size = BlockSize)
      raise MaxSizeError if size > MaxSize

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

      write([0xD0CF11E0, 0xA1B11AE1].pack("NN"))
      write([0x00, 0x00, 0x00, 0x00].pack("VVVV"))
      write([0x3E, 0x03, -2, 0x09].pack("vvvv"))
      write([0x06, 0x00, 0x00].pack("VVV"))
      write([@list_blocks, root_start].pack("VV"))
      write([0x00, 0x1000,-2].pack("VVV"))
      write([0x00, -2 ,0x00].pack("VVV"))

      unused = [-1].pack("V")

      1.upto(@list_blocks){
         root_start += 1
         write([root_start].pack("V"))
      }

      @list_blocks.upto(108){
         write(unused)
      }
   end

   # Write a big block depot
   def write_big_block_depot
      total_blocks = @list_blocks * 128
      used_blocks  = @big_blocks + @list_blocks + 2
      
      marker = [-3].pack("V")
      eoc    = [-2].pack("V")
      unused = [-1].pack("V")

      num_blocks = @big_blocks - 1

      1.upto(num_blocks){|n|
         write([n].pack("V"))
      }

      write eoc
      write eoc

      1.upto(@list_blocks){ write(marker) }

      used_blocks.upto(total_blocks){ write(unused) }
      
   end

   # Write property storage
   def write_property_storage
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

      zero = [0].pack("C")
      unknown = [0].pack("V")
      
      write(ord_name.pack("v*"))

      for n in 1..64-length
         write(zero)
      end

      write([length,type,-1,-1,dir].pack("vvVVV"))

      for n in 1..5
         write(unknown)
      end
      
      for n in 1..4
         write([0].pack("V"))
      end

      write([start,size].pack("VV"))
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
