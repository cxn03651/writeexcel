#########################################
# tc_formula.rb
#
# Tests for the Formula class (Formula.rb).
#########################################
base = File.basename(Dir.pwd)
if base == "test" || base =~ /spreadsheet/i
   Dir.chdir("..") if base == "test"
   $LOAD_PATH.unshift(Dir.pwd + "/lib/spreadsheet")
   Dir.chdir("test") rescue nil
end

require "test/unit"
require "formula"

class TC_Formula < Test::Unit::TestCase

   def setup
      @formula = Formula.new
   end

   def test_scan
      # scan must return array of token info
      string01 = '1 + 2 * LEN("String")'
      expected01 = [
         [:NUMBER, '1'],
         ['+',     '+'],
         [:NUMBER, '2'],
         ['*',     '*'],
         [:FUNC,   'LEN'],
         ['(',     '('],
         [:STRING, '"String"'],
         [')',     ')'],
         [:EOL,    nil]
      ]
      assert_kind_of(Array, @formula.scan(string01))
      assert_equal(expected01, @formula.scan(string01))

      string02 = 'IF(A1>=0,SIN(0),COS(90))'
      expected02 = [
         [:FUNC,   'IF'],
         ['(',     '('],
         [:REF2D,  'A1'],
         [:GE,     '>='],
         [:NUMBER, '0'],
         [',',     ','],
         [:FUNC,   'SIN'],
         ['(',     '('],
         [:NUMBER, '0'],
         [')',     ')'],
         [',',     ','],
         [:FUNC,   'COS'],
         ['(',     '('],
         [:NUMBER, '90'],
         [')',     ')'],
         [')',     ')'],
         [:EOL,    nil]
      ]
      assert_kind_of(Array, @formula.scan(string02))
      assert_equal(expected02, @formula.scan(string02))
      end

end
