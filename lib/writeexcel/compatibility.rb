# coding: utf-8
#
# Why would we ever use Ruby 1.8.7 when we can backport with something
# as simple as this?
#
# copied from prawn.
# modified by Hideo NAKAMURA
#
unless defined?(Encoding)
  class Encoding # :nodoc:
    class ConverterNotFoundError < StandardError; end
    class UndefinedConversionError < StandardError; end

    def self.const_missing (name)
      @looked_for ||= {}
      if !@looked_for.has_key?(name)
        begin
          @looked_for[name] = find(name)
        rescue ArgumentError
          @looked_for[name] = nil
        end
      end
      @looked_for[name]
    end

    def self.find (name)
      result   = try_name name
      result ||= try_name name.gsub(/-/, '_') if name =~ /-/
      raise ArgumentError, "unknown encoding name - #{name}" unless result
      result
    end
    def self.try_name (name)
      dname = name.to_s.downcase
      sel = SupportedEncodings.select{ |se| dname == se.name.downcase }
      return sel.first if 1 <= sel.size
      sel = SupportedEncodings.select{ |se| se.name.downcase =~ Regexp.new(dname) }
      return sel.first if 1 <= sel.size
      nil
    end

    attr_accessor :name, :value
    def initialize (name, value)
      @name  = name
      @value = value
    end

    def == (other)
      if other.is_a? Encoding
        other.value == @value
      elsif other.is_a? Fixnum
        other == @value
      else
        other == @name
      end
    end

    SupportedEncodings = [
      Encoding.new('ASCII', 0),
      Encoding.new('US_ASCII', 0),
      Encoding.new('BINARY', 1),
      Encoding.new('ASCII_8BIT', 1),
      Encoding.new('UTF_8', 2),
      Encoding.new('EUCJP', 3),
      Encoding.new('SJIS', 4),
      Encoding.new('UTF_16LE', 5),
      Encoding.new('UTF_16BE', 6)
    ]
    # ASCII    = 0
    # BINARY   = 1
    # UTF_8    = 2
    # EUCJP    = 3
    # SJIS     = 4
    # UTF_16LE = 5
    # UTF_16BE = 6
  end
end

class String #:nodoc:
  def first_line
    self.each_line { |line| return line }
  end
  unless "".respond_to?(:lines)
    alias_method :lines, :to_a
  end
  unless "".respond_to?(:each_char)
    def each_char #:nodoc:
      # copied from jcode
      if block_given?
        scan(/./m) { |x| yield x }
      else
        scan(/./m)
      end
    end
  end

  unless "".respond_to?(:encode)
    @encoding = Encoding::UTF_8

    def encode(encoding) # :nodoc:
      require 'nkf'
      @encoding ||= Encoding::UTF_8
      if @encoding == Encoding::UTF_8
        # supported only $KCODE = 'u'. so @encoding.nil? means UTF_8.
        case encoding
        when /ASCII$/i
          if self.mbchar?('UTF8')
            raise Encoding::UndefinedConversionError
          else
            str = String.new(self)
            str.force_encoding(encoding)
            str
          end
        when /(BINARY|ASCII[-_]8BIT)/i
          if self.mbchar?('UTF8')
            raise Encoding::UndefinedConversionError
          else
            str = String.new(self)
            str.force_encoding(encoding)
            str
          end
        when /UTF_8/i
          raise Encoding::ConverterNotFoundError
        when /EUCJP/i, /SJIS/i
          enc = encoding =~ /EUCJP/i ? 'e' : 's'
          str = NKF.nkf("-#{enc} -m0 -W", self)
          str.force_encoding(encoding)
          str
        when /UTF_16LE/i, /UTF_16BE/i
          raise Encoding::ConverterNotFoundError
        else
          raise "Sorry, encoding #{encoding} is not supported by WriteExcel."
        end
      elsif @encoding == Encoding::ASCII
        case encoding
        when /ASCII/i, /BINARY/i, /EUCJP/i, /SJIS/i
          str = String.new(self)
          str.force_encoding(encoding)
          str
        when /UTF_8/i, /UTF_16LE/i, /UTF_16BE/i
          raise Encoding::ConverterNotFoundError
        else
          raise "Sorry, encoding #{encoding} is not supported by WriteExcel."
        end
      elsif @encoding == Encoding::BINARY
        case encoding
        when /ASCII$/i, /EUCJP/i, /SJIS/i
          if self.ascii_only? || self.mbchar?('UTF8') || self.mbchar?('EUCJP') || self.mbchar?('SJIS')
            raise Encoding::UndefinedConversionError
          else
            str = String.new(self)
            str.force_encoding(encoding)
            str
          end
        when /(BINARY|ASCII[-_]8BIT)/i
          self
        when /UTF_8/i, /UTF_16LE/i, /UTF_16BE/i
          raise Encoding::ConverterNotFoundError
        else
          raise "Sorry, encoding #{encoding} is not supported by WriteExcel."
        end
      elsif @encoding == Encoding::EUCJP || @encoding == Encoding::SJIS
        type   = @encoding == Encoding::EUCJP ? 'EUCJP' : 'SJIS'
        inenc  = @encoding == Encoding::EUCJP ? 'e' : 's'
        case encoding
        when /ASCII$/i
          if self.mbchar?(type)
            raise Encoding::UndefinedConversionError
          else
            str = String.new(self)
            str.force_encoding(encoding)
            str
          end
        when /(BINARY|ASCII[-_]8BIT)/i
          if self.mbchar?(type)
            raise Encoding::UndefinedConversionError
          else
            str = String.new(self)
            str.force_encoding(encoding)
            str
          end
        when /UTF_8/i
          raise Encoding::ConverterNotFoundError
        when /EUCJP/i, /SJIS/i
          outenc = encoding =~ /EUCJP/i ? 'E' : 'S'
          str = NKF.nkf("-#{inenc} -#{outenc}", self)
          str.force_encoding(encoding)
          str
        when /UTF_16LE/i, /UTF_16BE/i
          raise Encoding::ConverterNotFoundError
        else
          raise "Sorry, encoding #{encoding} is not supported by WriteExcel."
        end
      elsif @encoding == Encoding::UTF_16LE || @encoding == Encoding::UTF_16BE
        enc = @encoding == Encoding::UTF_16LE ? 'L' : 'B'
        utf8 = NKF.nkf("-w -m0 -W16#{enc}", self)
        case encoding
        when /ASCII$/i
          if utf8.mbchar?
            raise Encoding::UndefinedConversionError
          else
            str = String.new(self)
            str.force_encoding(encoding)
            str
          end
        when /(BINARY|ASCII[-_]8BIT)/i
          if utf8.mbchar?
            raise Encoding::UndefinedConversionError
          else
            str = String.new(self)
            str.force_encoding(encoding)
            str
          end
        when /UTF_8/i
          raise Encoding::ConverterNotFoundError
        when /EUCJP/i, /SJIS/i
            str = String.new(self)
            str.force_encoding(encoding)
            str
        when /UTF_16LE/i, /UTF_16BE/i
          raise Encoding::ConverterNotFoundError
        else
          raise "Sorry, encoding #{encoding} is not supported by WriteExcel."
        end
      else
      end
    end

    def self_with_encoding(encoding)
      found = Encoding.find(encoding)
      if found
        @encoding = found
      else
        raise "Sorry, encoding #{encoding} is not supported by WriteExcel."
      end
      self
    end
    private :self_with_encoding
  end

  unless "".respond_to?(:encoding)
    def encoding
      @encoding ||= Encoding::UTF_8
    end
  end

  unless "".respond_to?(:bytesize)
    def bytesize # :nodoc:
      self.length
    end
  end

  unless "".respond_to?(:ord)
    def ord
      self[0]
    end
  end

  unless "".respond_to?(:force_encoding)
    def force_encoding(encoding)
      if encoding.respond_to?(:to_str)
        found = Encoding.find(encoding)
        @encoding = found if found
      else
        @encoding = encoding if encoding.is_a?(Encoding)
      end
      self
    end
  end

  unless "".respond_to?(:ascii_only?)
    def ascii_only?
      !!(self =~ /[^!"#\$%&'\(\)\*\+,\-\.\/\:\;<=>\?@0-9A-Za-z_\[\\\]\{\}^` ~\0\n]/)
    end
  end

  if RUBY_VERSION < "1.9"
    unless "".respond_to?(:mbchar?)
      PATTERN_SJIS = '[\x81-\x9f\xe0-\xef][\x40-\x7e\x80-\xfc]' # :nodoc:
      PATTERN_EUC = '[\xa1-\xfe][\xa1-\xfe]' # :nodoc:
      PATTERN_UTF8 = '[\xc0-\xdf][\x80-\xbf]|[\xe0-\xef][\x80-\xbf][\x80-\xbf]' # :nodoc:

      RE_SJIS = Regexp.new(PATTERN_SJIS, 0, 'n') # :nodoc:
      RE_EUC = Regexp.new(PATTERN_EUC, 0, 'n') # :nodoc:
      RE_UTF8 = Regexp.new(PATTERN_UTF8, 0, 'n') # :nodoc:

      def mbchar?(type = nil)# :nodoc:     idea copied from jcode.rb
        if (!type.nil? && type =~ /SJIS/i) || $KCODE == 'SJIS'
          self =~ RE_SJIS
        elsif (!type.nil? && type =~ /EUCJP/i) || $KCODE == 'EUC'
          self =~ RE_EUC
        elsif (!type.nil? && type =~ /UTF_8/i) || $KCODE == 'UTF8'
          self =~ RE_UTF8
        else
          nil
        end
      end
    end
  end
end

unless File.respond_to?(:binread)
  def File.binread(file) #:nodoc:
    File.open(file,"rb") { |f| f.read }
  end
end

if RUBY_VERSION < "1.9"

  def ruby_18 #:nodoc:
    yield
  end

  def ruby_19 #:nodoc:
    false
  end

else

  def ruby_18 #:nodoc:
    false
  end

  def ruby_19 #:nodoc:
    yield
  end
end
