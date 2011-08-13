module Writeexcel

class Worksheet < BIFFWriter
  require 'writeexcel/helper'

  class Collection
    def initialize
      @items = {}
    end

    def <<(item)
      @items[item.row] = { item.col => item }
    end

    def array
      return @array if @array

      @array = []
      @items.keys.sort.each do |row|
        @items[row].keys.sort.each do |col|
          @array << @items[row][col]
        end
      end
      @array
    end

  end

  class Comments < Collection
    attr_writer :visible

    def initialize
      super
      @visible  = false
    end

    def visible?
      @visible
    end
  end

  class Comment
    attr_reader :row, :col, :string, :encoding, :author, :author_encoding, :visible, :color, :vertices

    def initialize(worksheet, row, col, string, options = {})
      @worksheet = worksheet
      @row, @col = row, col
      @params = params_with(options)
      @string, @params[:encoding] = string_and_encoding(string, @params[:encoding], 'comment')

      # Limit the string to the max number of chars (not bytes).
      max_len = 32767
      max_len = max_len * 2 if @params[:encoding] != 0

      if @string.bytesize > max_len
        @string = @string[0 .. max_len]
      end
      @encoding        = @params[:encoding]
      @author          = @params[:author]
      @author_encoding = @params[:author_encoding]
      @visible         = @params[:visible]
      @color           = @params[:color]
      @vertices        = calc_vertices
    end

    #
    # Write the worksheet NOTE record that is part of cell comments.
    #
    def note_record(obj_id)   #:nodoc:
      comment_author = author
      comment_author_enc = author_encoding
      ruby_19 { comment_author = [comment_author].pack('a*') if comment_author.ascii_only? }
      record      = 0x001C               # Record identifier
      length      = 0x000C               # Bytes to follow

      comment_author     = '' unless comment_author
      comment_author_enc = 0  unless author_encoding

      # Use the visible flag if set by the user or else use the worksheet value.
      # The flag is also set in store_mso_opt_comment() but with the opposite
      # value.
      if visible
        comment_visible = visible != 0                 ? 0x0002 : 0x0000
      else
        comment_visible = @worksheet.comments_visible? ? 0x0002 : 0x0000
      end

      # Get the number of chars in the author string (not bytes).
      num_chars  = comment_author.bytesize
      num_chars  = num_chars / 2 if comment_author_enc != 0 && comment_author_enc

      # Null terminate the author string.
      comment_author =
        ruby_18 { comment_author + "\0" } ||
        ruby_19 { comment_author.force_encoding('BINARY') + "\0".force_encoding('BINARY') }

      # Pack the record.
      data    = [row, col, comment_visible, obj_id, num_chars, comment_author_enc].pack("vvvvvC")

      length  = data.bytesize + comment_author.bytesize
      header  = [record, length].pack("vv")

      header << data << comment_author
    end

    private

    def params_with(options)
      params = default_params.update(options)

      # Ensure that a width and height have been set.
      params[:width]  = default_width  unless params[:width]  && params[:width] != 0
      params[:width]  = params[:width] * params[:x_scale]  if params[:x_scale] != 0
      params[:height] = default_height unless params[:height] && params[:height] != 0
      params[:height] = params[:height] * params[:y_scale] if params[:y_scale] != 0

      params[:author], params[:author_encoding] =
          string_and_encoding(params[:author], params[:author_encoding], 'author')

      # Set the comment background colour.
      params[:color] = background_color(params[:color])

      # Set the default start cell and offsets for the comment. These are
      # generally fixed in relation to the parent cell. However there are
      # some edge cases for cells at the, er, edges.
      #
      params[:start_row] = default_start_row unless params[:start_row]
      params[:y_offset]  = default_y_offset  unless params[:y_offset]
      params[:start_col] = default_start_col unless params[:start_col]
      params[:x_offset]  = default_x_offset  unless params[:x_offset]

      params
    end

    def default_params
      {
        :author          => '',
        :author_encoding => 0,
        :encoding        => 0,
        :color           => nil,
        :start_cell      => nil,
        :start_col       => nil,
        :start_row       => nil,
        :visible         => nil,
        :width           => default_width,
        :height          => default_height,
        :x_offset        => nil,
        :x_scale         => 1,
        :y_offset        => nil,
        :y_scale         => 1
      }
    end

    def default_width
      128
    end

    def default_height
      74
    end

    def default_start_row
      case @row
      when 0     then 0
      when 65533 then 65529
      when 65534 then 65530
      when 65535 then 65531
      else            @row -1
      end
    end

    def default_y_offset
      case @row
      when 0     then 2
      when 65533 then 4
      when 65534 then 4
      when 65535 then 2
      else            7
      end
    end

    def default_start_col
      case @col
      when 253   then 250
      when 254   then 251
      when 255   then 252
      else            @col + 1
      end
    end

    def default_x_offset
      case @col
      when 253   then 49
      when 254   then 49
      when 255   then 49
      else            15
      end
    end

    def string_and_encoding(string, encoding, type)
      string = convert_to_ascii_if_ascii(string)
      if encoding != 0
        raise "Uneven number of bytes in #{type} string" if string.bytesize % 2 != 0
        # Change from UTF-16BE to UTF-16LE
        string = string.unpack('n*').pack('v*')
      # Handle utf8 strings
      else
        if is_utf8?(string)
          string = NKF.nkf('-w16L0 -m0 -W', string)
          ruby_19 { string.force_encoding('UTF-16LE') }
          encoding = 1
        end
      end
      [string, encoding]
    end

    def background_color(color)
      color = Colors.new.get_color(color)
      color = 0x50 if color == 0x7FFF  # Default color.
      color
    end

    # Calculate the positions of comment object.
    def calc_vertices
      @worksheet.position_object( @params[:start_col],
        @params[:start_row],
        @params[:x_offset],
        @params[:y_offset],
        @params[:width],
        @params[:height]
      )
    end
  end
end

end
