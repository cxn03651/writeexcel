# coding: utf-8
#
# Why would we ever use Ruby 1.8.7 when we can backport with something
# as simple as this?
#
# copied from prawn.
# modified by Hideo NAKAMURA
#
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
    def encode(encoding) # :nodoc:
      require 'nkf'

      if @encoding.nil? || @encoding == Encoding::UTF_8
        # supported only $KCODE = 'u'. so @encoding.nil? means UTF_8.
        case encoding
        when /ASCII/i
          if self.mbchar?('UTF8')
            raise Encoding::UndefinedConversionError
          else
            str = String.new(self)
            str.force_encoding(encoding)
            str
          end
        when /BINARY/i
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
        when /ASCII/i
          if self.ascii_only? || self.mbchar?('UTF8') || self.mbchar?('EUCJP') || self.mbchar?('SJIS')
            raise Encoding::UndefinedConversionError
          else
            str = String.new(self)
            str.force_encoding(encoding)
            str
          end
        when /BINARY/i
          self
        when /EUCJP/i, /SJIS/i
          if self.ascii_only? || self.mbchar?('UTF8') || self.mbchar?('EUCJP') || self.mbchar?('SJIS')
            raise Encoding::UndefinedConversionError
          else
            str = String.new(self)
            str.force_encoding(encoding)
            str
          end
        when /UTF_8/i, /UTF_16LE/i, /UTF_16BE/i
          raise Encoding::ConverterNotFoundError
        else
          raise "Sorry, encoding #{encoding} is not supported by WriteExcel."
        end
      elsif @encoding == Encoding::EUCJP || @encoding == Encoding::SJIS
        type   = @encoding == Encoding::EUCJP ? 'EUCJP' : 'SJIS'
        inenc  = @encoding == Encoding::EUCJP ? 'e' : 's'
        case encoding
        when /ASCII/i
          if self.mbchar?(type)
            raise Encoding::UndefinedConversionError
          else
            str = String.new(self)
            str.force_encoding(encoding)
            str
          end
        when /BINARY/i
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
        when /ASCII/i
          if utf8.mbchar?
            raise Encoding::UndefinedConversionError
          else
            str = String.new(self)
            str.force_encoding(encoding)
            str
          end
        when /BINARY/i
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
      if encoding =~ /ASCII/i
        @encoding = Encoding::ASCII
      elsif encoding =~ /BINARY/i
        @encoding = Encoding::BINARY
      elsif encoding =~ /UTF_8/i
        @encoding = Encoding::UTF_8
      elsif encoding =~ /EUCJP/i
        @encoding = Encoding::EUCJP
      elsif encoding =~ /SJIS/i
        @encoding = Encoding::SJIS
      elsif encoding =~ /UTF_16LE/i
        @encoding = Encoding::UTF_16LE
      elsif encoding =~ /UTF_16BE/i
        @encoding = Encoding::UTF_16BE
      else
        raise "Sorry, encoding #{encoding} is not supported by WriteExcel."
      end
      self
    end
    private :self_with_encoding
  end

  unless "".respond_to?(:encoding)
    def encoding
      @encoding.nil? ? Encoding::UTF_8 : @encoding
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
      @encoding = case encoding
        when /ASCII/i
          Encoding::ASCII
        when /BINARY/i
          Encoding::BINARY
        when /UTF_8/i
          Encoding::UTF_8
        when /EUCJP/i
          Encoding::EUCJP
        when /SJIS/i
          Encoding::SJIS
        when /UTF_16LE/i
          Encoding::UTF_16LE
        when /UTF_16BE/i
          Encoding::UTF_16BE
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

unless defined?(Encoding)
  class Encoding # :nodoc:
    class ConverterNotFoundError < StandardError; end
    class UndefinedConversionError < StandardError; end

    ASCII    = 0
    BINARY   = 1
    UTF_8    = 2
    EUCJP    = 3
    SJIS     = 4
    UTF_16LE = 5
    UTF_16BE = 6
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
