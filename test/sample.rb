require 'format'

foo = Format.new

foo.set_format_properties(:bg_color => 'red')
p foo.bg_color
