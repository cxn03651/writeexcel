class Format

   COLORS = {
      'aqua'    => 0x0F,
      'black'   => 0x08,
      'blue'    => 0x0C,
      'brown'   => 0x10,
      'cyan'    => 0x0F,
      'fuchsia' => 0x0E,
      'gray'    => 0x17,
      'grey'    => 0x17,
      'green'   => 0x11,
      'lime'    => 0x0B,
      'magenta' => 0x0E,
      'navy'    => 0x12,
      'orange'  => 0x1D,
      'purple'  => 0x24,
      'red'     => 0x0A,
      'silver'  => 0x16,
      'white'   => 0x09,
      'yellow'  => 0x0D
   }

   attr_accessor :xf_index

   def initialize(args={}, xf_index=0)
      defaults = {}

      defaults.update(:color     => 0x7FFF, :bold   => 0x0190)
      defaults.update(:fg_color  => 0x40, :pattern  => 0,  :size => 10)
      defaults.update(:bg_color  => 0x41, :rotation => 0,  :font => "Arial")
      defaults.update(:underline => 0,    :italic   => 0,  :top  => 0)
      defaults.update(:bottom    => 0,    :right    => 0,  :left => 0)

      defaults.update(:font_index     => 0, :font_family  => 0)
      defaults.update(:font_strikeout => 0, :font_script  => 0)
      defaults.update(:font_outline   => 0, :left_color   => 0)
      defaults.update(:font_charset   => 0, :right_color  => 0)
      defaults.update(:font_shadow    => 0, :top_color    => 0x40)
      defaults.update(:text_v_align   => 2, :bottom_color => 0x40)
      defaults.update(:text_h_align   => 0, :num_format   => 0)
      defaults.update(:text_justlast  => 0, :text_wrap    => 0)

      ## convenience methods
      defaults.update(:border => 0, :align => 'left')

      ########################################################################
      # We must manually create accessors for these so that they can handle
      # both 0/1 and true/false.
      ########################################################################
      no_acc = [:bold,:italic,:underline,:strikeout,:text_wrap,:text_justlast]
      no_acc.push(:fg_color,:bg_color,:color,:font_outline,:font_shadow)
      no_acc.push(:align,:border)

      args.each{|key,val|
         key = key.to_s.downcase.intern
         val = 1 if val == true
         val = 0 if val == false
         defaults.fetch(key)
         defaults.update(key=>val)
      }

      defaults.each{|key,val|
         unless no_acc.member?(key)
            self.class.send(:attr_accessor,"#{key}")
         end
         send("#{key}=",val)
      }

      @xf_index = xf_index

      yield self if block_given?
   end

   def color=(colour)
      if COLORS.has_key?(colour)
         @color = COLORS[colour]
      else
         if colour.kind_of?(String)
            raise ArgumentError, "unknown color"
         else
            @color = colour
         end
      end
      @color
   end

   def bg_color=(colour)
      if COLORS.has_key?(colour)
         @bg_color = COLORS[colour]
      else
         if colour.kind_of?(String)
            raise ArgumentError, "unknown color"
         else
            @bg_color = colour
         end
      end
      @bg_color
   end

   def fg_color=(colour)
      if COLORS.has_key?(colour)
         @fg_color = COLORS[colour]
      else
         if colour.kind_of?(String)
            raise ArgumentError, "unknown color"
         else
            @fg_color = colour
         end
      end
      @fg_color
   end

   def fg_color
      @fg_color
   end

   def bg_color
      @bg_color
   end

   # Should I return the stringified version of the color if applicable?
   def color
      #COLORS.invert.fetch(@color)
      @color
   end

   def italic
      return true if @italic >= 1
      return false
   end

   def italic=(val)
      val = 1 if val == true
      val = 0 if val == false
      @italic = val
   end

   def font_shadow
      return true if @font_shadow == 1
      return false
   end

   def font_shadow=(val)
      val = 1 if val == true
      val = 0 if val == false
      @font_shadow = val
   end

   def font_outline
      return true if @font_outline == 1
      return false
   end

   def font_outline=(val)
      val = 1 if val == true
      val = 0 if val == false
      @font_outline = val
   end

   def text_justlast
      return true if @text_justlast == 1
      return false
   end

   def text_justlast=(val)
      val = 1 if val == true
      val = 0 if val == false
      @text_justlast = val
   end

   def text_wrap
      return true if @text_wrap == 1
      return false
   end

   def text_wrap=(val)
      val = 1 if val == true
      val = 0 if val == false
      @text_wrap = val
   end

   def strikeout
      return true if @strikeout == 1
      return false
   end

   def strikeout=(val)
      val = 1 if val == true
      val = 0 if val == false
      @strikeout = val
   end

   def underline
      return true if @underline == 1
      return false
   end

   def underline=(val)
      val = 1 if val == true
      val = 0 if val == false
      @underline = val
   end


   def xf_biff(style=0)
      atr_num = 0
      atr_num = 1 if @num_format != 0

      atr_fnt = 0
      atr_fnt = 1 if @font_index != 0

      atr_alc = @text_wrap
      atr_bdr = [@bottom,@top,@left,@right].find{ |n| n > 0 } || 0
      atr_pat = [@fg_color,@bg_color,@pattern].find{ |n| n > 0 } || 0

      atr_prot    = 0

      @bottom_color = 0 if @bottom == 0
      @top_color    = 0 if @top    == 0
      @right_color  = 0 if @right  == 0
      @left_color   = 0 if @left   == 0

      record         = 0x00E0
      length         = 0x0010

      align  = @text_h_align
      align  |= @text_wrap     << 3
      align  |= @text_v_align  << 4
      align  |= @text_justlast << 7
      align  |= @rotation      << 8
      align  |= atr_num        << 10
      align  |= atr_fnt        << 11
      align  |= atr_alc        << 12
      align  |= atr_bdr        << 13
      align  |= atr_pat        << 14
      align  |= atr_prot       << 15

      # Assume a solid fill color if the bg_color or fg_color are set but
      # the pattern value is not set
      if @pattern <= 0x01 and @bg_color != 0x41 and @fg_color == 0x40
         @fg_color = @bg_color
         @bg_color = 0x40
         @pattern  = 1
      end

      if @pattern < 0x01 and @bg_color == 0x41 and @fg_color != 0x40
         @bg_color = 0x40
         @pattern  = 1
      end

      icv   = @fg_color
      icv  |= @bg_color << 7

      fill = @pattern
      fill |= @bottom << 6
      fill |= @bottom_color << 9

      border1 = @top
      border1 |= @left      << 3
      border1 |= @right     << 6
      border1 |= @top_color << 9

      border2 = @left_color
      border2 |= @right_color << 7

      header = [record,length].pack("vv")
      fields = [@font_index,@num_format,style,align,icv,fill,border1,border2]
      data = fields.pack("vvvvvvvv")

      rv = header + data
      return rv
   end

   def font_biff
      dyheight = @size * 20
      cch      = @font.length
      record   = 0x31
      length   = 0x0F + cch
      reserved = 0x00

      grbit = 0x00
      grbit |= 0x02 if @italic > 0
      grbit |= 0x08 if @font_strikeout > 0
      grbit |= 0x10 if @font_outline > 0
      grbit |= 0x20 if @font_shadow > 0

      header = [record,length].pack("vv")
      fields = [dyheight,grbit,@color,@bold,@font_script,@underline,@font_family]
      fields.push(@font_charset,reserved,cch)

      data = fields.pack("vvvvvCCCCC")
      rv = header + data + @font

      return rv
   end

   def font_key
      key = @font.to_s + @size.to_s + @font_script.to_s + @underline.to_s
      key += @font_strikeout.to_s + @bold.to_s + @font_outline.to_s
      key += @font_family.to_s + @font_charset.to_s + @font_shadow.to_s
      key += @color.to_s + @italic.to_s
      return key
   end

   def align=(location = nil)
      return if location.nil?
      return if location.kind_of?(Fixnum)
      location.downcase!

      case location
      when 'left'
        @text_h_align = 1
      when 'center', 'centre'
        @text_h_align = 2
      when 'right'
        @text_h_align = 3
      when 'fill'
        @text_h_align = 4 
      when 'justify'
        @text_h_align = 5 
      when 'merge'
        @text_h_align = 6
      when 'top'
        @text_v_align = 0
      when 'vcentre', 'vcenter'
        @text_v_align = 1
      when 'bottom'
        @text_v_align = 2
      when 'vjustify'
        @text_v_align = 3
      end
   end

   def align
      [@text_h_align,@text_v_align]
   end

   def bold=(weight)
      weight = 1 if weight == true
      weight = 0 if weight == false
      weight = 0x2BC if weight.nil?
      weight = 0x2BC if weight == 1
      weight = 0x190 if weight == 0
      weight = 0x190 if weight < 0x064
      weight = 0x190 if weight > 0x3E8
      @bold = weight
      @bold
   end

   def bold
      return true if @bold >= 1
      return false
   end

   def border=(style)
      [@bottom,@top,@right,@left].each{ |attr| attr = style }
   end

   def border
      [@bottom,@top,@right,@left]
   end

   def border_color=(color)
      [@bottom_color,@top_color,@left_color,@right_color].each{ |a| a = color }
   end

   def border_color
      [@bottom_color,@top_color,@left_color,@right_color]
   end

end
