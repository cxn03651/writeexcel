# -*- coding: utf-8 -*-
#
# helper.rb
#
  # Convert to US_ASCII encoding if ascii characters only.
  def convert_to_ascii_if_ascii(str)
    ruby_18 do
      @encoding = str.mbchar? ? Encoding::UTF_8 : Encoding::US_ASCII
    end
    ruby_19 do
      if !str.nil? && str.ascii_only?
        str = [str].pack('a*')
      end
    end
    str
  end
  private :convert_to_ascii_if_ascii
