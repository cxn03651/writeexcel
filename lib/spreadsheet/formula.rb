require 'strscan'
require 'excelformulaparser'

class Formula < ExcelFormulaParser

   def initialize
   end

   def scan(formula)
      s = StringScanner.new(formula)
      q = []
      until s.eos?
         # order is important.
         if    s.scan /(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?/
            q.push [:NUMBER, s.matched]
         elsif s.scan /"([^"]|"")*"/
            q.push [:STRING, s.matched]
         elsif s.scan /\$?[A-I]?[A-Z]\$?\d+/
            q.push [:REF2D,  s.matched]
         elsif s.scan /[^!(,]+!\$?[A-I]?[A-Z]\$?\d+/
            q.push [:REF3D , s.matched]
         elsif s.scan /'[^']+'!\$?[A-I]?[A-Z]\$?\d+/
            q.push [:REF3D , s.matched]
         elsif s.scan /\$?[A-I]?[A-Z]\$?(\d+)?:\$?[A-I]?[A-Z]\$?(\d+)?/
            q.push [:RANGE2D , s.matched]
         elsif s.scan /[^!(,]+!\$?[A-I]?[A-Z]\$?(\d+)?:\$?[A-I]?[A-Z]\$?(\d+)?/
            q.push [:RANGE3D , s.matched]
         elsif s.scan /<=/
            q.push [:LE , s.matched]
         elsif s.scan />=/
            q.push [:GE , s.matched]
         elsif s.scan /<>/
            q.push [:NE , s.matched]
         elsif s.scan /</
            q.push [:LT , s.matched]
         elsif s.scan />/
            q.push [:GT , s.matched]
         elsif s.scan /TRUE/
            q.push [:TRUE, s.matched]
         elsif s.scan /FALSE/
            q.push [:FALSE, s.matched]
         elsif s.scan /[A-Z0-9_.]+/
            q.push [:FUNC,   s.matched]
         elsif s.scan /\s+/
            ;
         elsif s.scan /./
            q.push [s.matched, s.matched]
         end
      end
      q.push [:EOL, nil]
   end

   def parse(formula)
      @q = scan(formula)
      @q.push [false, nil]
      do_parse
   end

   def next_token
      @q.shift
   end

   def reverse(expression)
      q = []
      expression.each do |e|
         if e.kind_of?(Array)
            qq = reverse(e)
            qq.each { |ee| q.push ee }
         else
            q.push e
         end
      end
      q
   end

end

if $0 ==__FILE__


parser = Formula.new
puts
puts 'type "Q" to quit.'
puts
while true
  puts
  print '? '
  str = gets.chop!
  break if /q/i =~ str
  begin
    e = parser.parse(str)
    p   parser.reverse(e)
  rescue ParseError
    puts $!
  end
end

end
