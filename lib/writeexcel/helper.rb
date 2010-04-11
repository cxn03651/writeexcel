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
