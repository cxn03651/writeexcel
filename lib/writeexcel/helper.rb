# -*- coding: utf-8 -*-
#
# helper.rb
#
  # Convert to US_ASCII encoding if ascii characters only.
  def convert_to_ascii_if_ascii(str)
    return nil if str.nil?
    ruby_18 do
      enc = str.encoding
      begin
        str = str.encode('ASCII')
      rescue
        str.force_encoding(enc)
      end
    end ||
    ruby_19 do
      if !str.nil? && str.ascii_only?
        str = [str].pack('a*')
      end
    end
    str
  end
  private :convert_to_ascii_if_ascii


  def utf8_to_16be(utf8)
    utf16be = NKF.nkf('-w16B0 -m0 -W', utf8)
    utf16be.force_encoding('UTF-16BE')
  end
  private :utf8_to_16be

  def utf8_to_16le(utf8)
    utf16le = NKF.nkf('-w16L0 -m0 -W', utf8)
    utf16le.force_encoding('UTF-16LE')
  end
  private :utf8_to_16le

  def ascii_to_16be(ascii)
    ascii.unpack("C*").pack("n*")
    ascii.force_encoding('UTF-16BE')
  end
  private :ascii_to_16be
