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
      if encoding =~ /UTF-16LE/i
        @encoding = Encoding::UTF_16LE
        NKF.nkf('-w16L0 -m0 -W', self)
      elsif encoding =~ /UTF-16BE/i
        @encoding = Encoding::UTF_16BE
        NKF.nkf('-w16B0 -m0 -W', self)
      elsif encoding =~ /BINARY/i || encoding =~ /US-ASCII/i
        if self.mbchar?
          @encoding = Encoding::UTF_8
        else
          @encoding = Encoding::US_ASCII
        end
        self
      elsif encoding =~ /UTF_8/i
        @encoding = Encoding::UTF_8
        self
      end
    end
  end

  unless "".respond_to?(:encoding)
    def encoding
      if @encoding
        case $KCODE[0]
        when ?s, ?S
          Encoding::SJIS if self.mbchar?
        when ?e, ?E
          Encoding::EUCJP if self.mbchar?
        when ?u, ?U
          Encoding::UTF_8 if self.mbchar?
          @encoding
        else
          Encoding::US_ASCII
          @encoding
        end
      else
        @encoding
      end
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
      self
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

      def mbchar? # :nodoc:     copied from jcode.rb
        case $KCODE[0]
        when ?s, ?S
          self =~ RE_SJIS
        when ?e, ?E
          self =~ RE_EUC
        when ?u, ?U
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
    US_ASCII = 0
    UTF_8    = 1
    UTF_16BE = 2
    UTF_16LE = 3
    SJIS     = 4
    EUCJP    = 5
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
