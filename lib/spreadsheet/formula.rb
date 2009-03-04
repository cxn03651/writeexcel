require 'strscan'
require 'excelformulaparser'

class Formula < ExcelFormulaParser

   attr_accessor :byte_order, :workbook, :ext_sheets, :ext_refs, :ext_ref_count


   def initialize(byte_order)
      @byte_order     = byte_order,
      @workbook       = ""
      @ext_sheets     = {}
      @ext_refs       = {}
      @ext_ref_count  = 0
      initialize_hashes
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


   ###############################################################################

   private
   
   ###############################################################################


   ###############################################################################
   #
   # _initialize_hashes()
   #
   def initialize_hashes
   
       # The Excel ptg indices
       ptg = {
           'ptgExp'            => 0x01,
           'ptgTbl'            => 0x02,
           'ptgAdd'            => 0x03,
           'ptgSub'            => 0x04,
           'ptgMul'            => 0x05,
           'ptgDiv'            => 0x06,
           'ptgPower'          => 0x07,
           'ptgConcat'         => 0x08,
           'ptgLT'             => 0x09,
           'ptgLE'             => 0x0A,
           'ptgEQ'             => 0x0B,
           'ptgGE'             => 0x0C,
           'ptgGT'             => 0x0D,
           'ptgNE'             => 0x0E,
           'ptgIsect'          => 0x0F,
           'ptgUnion'          => 0x10,
           'ptgRange'          => 0x11,
           'ptgUplus'          => 0x12,
           'ptgUminus'         => 0x13,
           'ptgPercent'        => 0x14,
           'ptgParen'          => 0x15,
           'ptgMissArg'        => 0x16,
           'ptgStr'            => 0x17,
           'ptgAttr'           => 0x19,
           'ptgSheet'          => 0x1A,
           'ptgEndSheet'       => 0x1B,
           'ptgErr'            => 0x1C,
           'ptgBool'           => 0x1D,
           'ptgInt'            => 0x1E,
           'ptgNum'            => 0x1F,
           'ptgArray'          => 0x20,
           'ptgFunc'           => 0x21,
           'ptgFuncVar'        => 0x22,
           'ptgName'           => 0x23,
           'ptgRef'            => 0x24,
           'ptgArea'           => 0x25,
           'ptgMemArea'        => 0x26,
           'ptgMemErr'         => 0x27,
           'ptgMemNoMem'       => 0x28,
           'ptgMemFunc'        => 0x29,
           'ptgRefErr'         => 0x2A,
           'ptgAreaErr'        => 0x2B,
           'ptgRefN'           => 0x2C,
           'ptgAreaN'          => 0x2D,
           'ptgMemAreaN'       => 0x2E,
           'ptgMemNoMemN'      => 0x2F,
           'ptgNameX'          => 0x39,
           'ptgRef3d'          => 0x3A,
           'ptgArea3d'         => 0x3B,
           'ptgRefErr3d'       => 0x3C,
           'ptgAreaErr3d'      => 0x3D,
           'ptgArrayV'         => 0x40,
           'ptgFuncV'          => 0x41,
           'ptgFuncVarV'       => 0x42,
           'ptgNameV'          => 0x43,
           'ptgRefV'           => 0x44,
           'ptgAreaV'          => 0x45,
           'ptgMemAreaV'       => 0x46,
           'ptgMemErrV'        => 0x47,
           'ptgMemNoMemV'      => 0x48,
           'ptgMemFuncV'       => 0x49,
           'ptgRefErrV'        => 0x4A,
           'ptgAreaErrV'       => 0x4B,
           'ptgRefNV'          => 0x4C,
           'ptgAreaNV'         => 0x4D,
           'ptgMemAreaNV'      => 0x4E,
           'ptgMemNoMemN'      => 0x4F,
           'ptgFuncCEV'        => 0x58,
           'ptgNameXV'         => 0x59,
           'ptgRef3dV'         => 0x5A,
           'ptgArea3dV'        => 0x5B,
           'ptgRefErr3dV'      => 0x5C,
           'ptgAreaErr3d'      => 0x5D,
           'ptgArrayA'         => 0x60,
           'ptgFuncA'          => 0x61,
           'ptgFuncVarA'       => 0x62,
           'ptgNameA'          => 0x63,
           'ptgRefA'           => 0x64,
           'ptgAreaA'          => 0x65,
           'ptgMemAreaA'       => 0x66,
           'ptgMemErrA'        => 0x67,
           'ptgMemNoMemA'      => 0x68,
           'ptgMemFuncA'       => 0x69,
           'ptgRefErrA'        => 0x6A,
           'ptgAreaErrA'       => 0x6B,
           'ptgRefNA'          => 0x6C,
           'ptgAreaNA'         => 0x6D,
           'ptgMemAreaNA'      => 0x6E,
           'ptgMemNoMemN'      => 0x6F,
           'ptgFuncCEA'        => 0x78,
           'ptgNameXA'         => 0x79,
           'ptgRef3dA'         => 0x7A,
           'ptgArea3dA'        => 0x7B,
           'ptgRefErr3dA'      => 0x7C,
           'ptgAreaErr3d'      => 0x7D
       };
   
       # Thanks to Michael Meeks and Gnumeric for the initial arg values.
       #
       # The following hash was generated by "function_locale.pl" in the distro.
       # Refer to function_locale.pl for non-English function names.
       #
       # The array elements are as follow:
       # ptg:   The Excel function ptg code.
       # args:  The number of arguments that the function takes:
       #           >=0 is a fixed number of arguments.
       #           -1  is a variable  number of arguments.
       # class: The reference, value or array class of the function args.
       # vol:   The function is volatile.
       #
       functions  = {
           #                                     ptg  args  class  vol
           'COUNT'                         => [   0,   -1,    0,    0 ],
           'IF'                            => [   1,   -1,    1,    0 ],
           'ISNA'                          => [   2,    1,    1,    0 ],
           'ISERROR'                       => [   3,    1,    1,    0 ],
           'SUM'                           => [   4,   -1,    0,    0 ],
           'AVERAGE'                       => [   5,   -1,    0,    0 ],
           'MIN'                           => [   6,   -1,    0,    0 ],
           'MAX'                           => [   7,   -1,    0,    0 ],
           'ROW'                           => [   8,   -1,    0,    0 ],
           'COLUMN'                        => [   9,   -1,    0,    0 ],
           'NA'                            => [  10,    0,    0,    0 ],
           'NPV'                           => [  11,   -1,    1,    0 ],
           'STDEV'                         => [  12,   -1,    0,    0 ],
           'DOLLAR'                        => [  13,   -1,    1,    0 ],
           'FIXED'                         => [  14,   -1,    1,    0 ],
           'SIN'                           => [  15,    1,    1,    0 ],
           'COS'                           => [  16,    1,    1,    0 ],
           'TAN'                           => [  17,    1,    1,    0 ],
           'ATAN'                          => [  18,    1,    1,    0 ],
           'PI'                            => [  19,    0,    1,    0 ],
           'SQRT'                          => [  20,    1,    1,    0 ],
           'EXP'                           => [  21,    1,    1,    0 ],
           'LN'                            => [  22,    1,    1,    0 ],
           'LOG10'                         => [  23,    1,    1,    0 ],
           'ABS'                           => [  24,    1,    1,    0 ],
           'INT'                           => [  25,    1,    1,    0 ],
           'SIGN'                          => [  26,    1,    1,    0 ],
           'ROUND'                         => [  27,    2,    1,    0 ],
           'LOOKUP'                        => [  28,   -1,    0,    0 ],
           'INDEX'                         => [  29,   -1,    0,    1 ],
           'REPT'                          => [  30,    2,    1,    0 ],
           'MID'                           => [  31,    3,    1,    0 ],
           'LEN'                           => [  32,    1,    1,    0 ],
           'VALUE'                         => [  33,    1,    1,    0 ],
           'TRUE'                          => [  34,    0,    1,    0 ],
           'FALSE'                         => [  35,    0,    1,    0 ],
           'AND'                           => [  36,   -1,    1,    0 ],
           'OR'                            => [  37,   -1,    1,    0 ],
           'NOT'                           => [  38,    1,    1,    0 ],
           'MOD'                           => [  39,    2,    1,    0 ],
           'DCOUNT'                        => [  40,    3,    0,    0 ],
           'DSUM'                          => [  41,    3,    0,    0 ],
           'DAVERAGE'                      => [  42,    3,    0,    0 ],
           'DMIN'                          => [  43,    3,    0,    0 ],
           'DMAX'                          => [  44,    3,    0,    0 ],
           'DSTDEV'                        => [  45,    3,    0,    0 ],
           'VAR'                           => [  46,   -1,    0,    0 ],
           'DVAR'                          => [  47,    3,    0,    0 ],
           'TEXT'                          => [  48,    2,    1,    0 ],
           'LINEST'                        => [  49,   -1,    0,    0 ],
           'TREND'                         => [  50,   -1,    0,    0 ],
           'LOGEST'                        => [  51,   -1,    0,    0 ],
           'GROWTH'                        => [  52,   -1,    0,    0 ],
           'PV'                            => [  56,   -1,    1,    0 ],
           'FV'                            => [  57,   -1,    1,    0 ],
           'NPER'                          => [  58,   -1,    1,    0 ],
           'PMT'                           => [  59,   -1,    1,    0 ],
           'RATE'                          => [  60,   -1,    1,    0 ],
           'MIRR'                          => [  61,    3,    0,    0 ],
           'IRR'                           => [  62,   -1,    0,    0 ],
           'RAND'                          => [  63,    0,    1,    1 ],
           'MATCH'                         => [  64,   -1,    0,    0 ],
           'DATE'                          => [  65,    3,    1,    0 ],
           'TIME'                          => [  66,    3,    1,    0 ],
           'DAY'                           => [  67,    1,    1,    0 ],
           'MONTH'                         => [  68,    1,    1,    0 ],
           'YEAR'                          => [  69,    1,    1,    0 ],
           'WEEKDAY'                       => [  70,   -1,    1,    0 ],
           'HOUR'                          => [  71,    1,    1,    0 ],
           'MINUTE'                        => [  72,    1,    1,    0 ],
           'SECOND'                        => [  73,    1,    1,    0 ],
           'NOW'                           => [  74,    0,    1,    1 ],
           'AREAS'                         => [  75,    1,    0,    1 ],
           'ROWS'                          => [  76,    1,    0,    1 ],
           'COLUMNS'                       => [  77,    1,    0,    1 ],
           'OFFSET'                        => [  78,   -1,    0,    1 ],
           'SEARCH'                        => [  82,   -1,    1,    0 ],
           'TRANSPOSE'                     => [  83,    1,    1,    0 ],
           'TYPE'                          => [  86,    1,    1,    0 ],
           'ATAN2'                         => [  97,    2,    1,    0 ],
           'ASIN'                          => [  98,    1,    1,    0 ],
           'ACOS'                          => [  99,    1,    1,    0 ],
           'CHOOSE'                        => [ 100,   -1,    1,    0 ],
           'HLOOKUP'                       => [ 101,   -1,    0,    0 ],
           'VLOOKUP'                       => [ 102,   -1,    0,    0 ],
           'ISREF'                         => [ 105,    1,    0,    0 ],
           'LOG'                           => [ 109,   -1,    1,    0 ],
           'CHAR'                          => [ 111,    1,    1,    0 ],
           'LOWER'                         => [ 112,    1,    1,    0 ],
           'UPPER'                         => [ 113,    1,    1,    0 ],
           'PROPER'                        => [ 114,    1,    1,    0 ],
           'LEFT'                          => [ 115,   -1,    1,    0 ],
           'RIGHT'                         => [ 116,   -1,    1,    0 ],
           'EXACT'                         => [ 117,    2,    1,    0 ],
           'TRIM'                          => [ 118,    1,    1,    0 ],
           'REPLACE'                       => [ 119,    4,    1,    0 ],
           'SUBSTITUTE'                    => [ 120,   -1,    1,    0 ],
           'CODE'                          => [ 121,    1,    1,    0 ],
           'FIND'                          => [ 124,   -1,    1,    0 ],
           'CELL'                          => [ 125,   -1,    0,    1 ],
           'ISERR'                         => [ 126,    1,    1,    0 ],
           'ISTEXT'                        => [ 127,    1,    1,    0 ],
           'ISNUMBER'                      => [ 128,    1,    1,    0 ],
           'ISBLANK'                       => [ 129,    1,    1,    0 ],
           'T'                             => [ 130,    1,    0,    0 ],
           'N'                             => [ 131,    1,    0,    0 ],
           'DATEVALUE'                     => [ 140,    1,    1,    0 ],
           'TIMEVALUE'                     => [ 141,    1,    1,    0 ],
           'SLN'                           => [ 142,    3,    1,    0 ],
           'SYD'                           => [ 143,    4,    1,    0 ],
           'DDB'                           => [ 144,   -1,    1,    0 ],
           'INDIRECT'                      => [ 148,   -1,    1,    1 ],
           'CALL'                          => [ 150,   -1,    1,    0 ],
           'CLEAN'                         => [ 162,    1,    1,    0 ],
           'MDETERM'                       => [ 163,    1,    2,    0 ],
           'MINVERSE'                      => [ 164,    1,    2,    0 ],
           'MMULT'                         => [ 165,    2,    2,    0 ],
           'IPMT'                          => [ 167,   -1,    1,    0 ],
           'PPMT'                          => [ 168,   -1,    1,    0 ],
           'COUNTA'                        => [ 169,   -1,    0,    0 ],
           'PRODUCT'                       => [ 183,   -1,    0,    0 ],
           'FACT'                          => [ 184,    1,    1,    0 ],
           'DPRODUCT'                      => [ 189,    3,    0,    0 ],
           'ISNONTEXT'                     => [ 190,    1,    1,    0 ],
           'STDEVP'                        => [ 193,   -1,    0,    0 ],
           'VARP'                          => [ 194,   -1,    0,    0 ],
           'DSTDEVP'                       => [ 195,    3,    0,    0 ],
           'DVARP'                         => [ 196,    3,    0,    0 ],
           'TRUNC'                         => [ 197,   -1,    1,    0 ],
           'ISLOGICAL'                     => [ 198,    1,    1,    0 ],
           'DCOUNTA'                       => [ 199,    3,    0,    0 ],
           'ROUNDUP'                       => [ 212,    2,    1,    0 ],
           'ROUNDDOWN'                     => [ 213,    2,    1,    0 ],
           'RANK'                          => [ 216,   -1,    0,    0 ],
           'ADDRESS'                       => [ 219,   -1,    1,    0 ],
           'DAYS360'                       => [ 220,   -1,    1,    0 ],
           'TODAY'                         => [ 221,    0,    1,    1 ],
           'VDB'                           => [ 222,   -1,    1,    0 ],
           'MEDIAN'                        => [ 227,   -1,    0,    0 ],
           'SUMPRODUCT'                    => [ 228,   -1,    2,    0 ],
           'SINH'                          => [ 229,    1,    1,    0 ],
           'COSH'                          => [ 230,    1,    1,    0 ],
           'TANH'                          => [ 231,    1,    1,    0 ],
           'ASINH'                         => [ 232,    1,    1,    0 ],
           'ACOSH'                         => [ 233,    1,    1,    0 ],
           'ATANH'                         => [ 234,    1,    1,    0 ],
           'DGET'                          => [ 235,    3,    0,    0 ],
           'INFO'                          => [ 244,    1,    1,    1 ],
           'DB'                            => [ 247,   -1,    1,    0 ],
           'FREQUENCY'                     => [ 252,    2,    0,    0 ],
           'ERROR.TYPE'                    => [ 261,    1,    1,    0 ],
           'REGISTER.ID'                   => [ 267,   -1,    1,    0 ],
           'AVEDEV'                        => [ 269,   -1,    0,    0 ],
           'BETADIST'                      => [ 270,   -1,    1,    0 ],
           'GAMMALN'                       => [ 271,    1,    1,    0 ],
           'BETAINV'                       => [ 272,   -1,    1,    0 ],
           'BINOMDIST'                     => [ 273,    4,    1,    0 ],
           'CHIDIST'                       => [ 274,    2,    1,    0 ],
           'CHIINV'                        => [ 275,    2,    1,    0 ],
           'COMBIN'                        => [ 276,    2,    1,    0 ],
           'CONFIDENCE'                    => [ 277,    3,    1,    0 ],
           'CRITBINOM'                     => [ 278,    3,    1,    0 ],
           'EVEN'                          => [ 279,    1,    1,    0 ],
           'EXPONDIST'                     => [ 280,    3,    1,    0 ],
           'FDIST'                         => [ 281,    3,    1,    0 ],
           'FINV'                          => [ 282,    3,    1,    0 ],
           'FISHER'                        => [ 283,    1,    1,    0 ],
           'FISHERINV'                     => [ 284,    1,    1,    0 ],
           'FLOOR'                         => [ 285,    2,    1,    0 ],
           'GAMMADIST'                     => [ 286,    4,    1,    0 ],
           'GAMMAINV'                      => [ 287,    3,    1,    0 ],
           'CEILING'                       => [ 288,    2,    1,    0 ],
           'HYPGEOMDIST'                   => [ 289,    4,    1,    0 ],
           'LOGNORMDIST'                   => [ 290,    3,    1,    0 ],
           'LOGINV'                        => [ 291,    3,    1,    0 ],
           'NEGBINOMDIST'                  => [ 292,    3,    1,    0 ],
           'NORMDIST'                      => [ 293,    4,    1,    0 ],
           'NORMSDIST'                     => [ 294,    1,    1,    0 ],
           'NORMINV'                       => [ 295,    3,    1,    0 ],
           'NORMSINV'                      => [ 296,    1,    1,    0 ],
           'STANDARDIZE'                   => [ 297,    3,    1,    0 ],
           'ODD'                           => [ 298,    1,    1,    0 ],
           'PERMUT'                        => [ 299,    2,    1,    0 ],
           'POISSON'                       => [ 300,    3,    1,    0 ],
           'TDIST'                         => [ 301,    3,    1,    0 ],
           'WEIBULL'                       => [ 302,    4,    1,    0 ],
           'SUMXMY2'                       => [ 303,    2,    2,    0 ],
           'SUMX2MY2'                      => [ 304,    2,    2,    0 ],
           'SUMX2PY2'                      => [ 305,    2,    2,    0 ],
           'CHITEST'                       => [ 306,    2,    2,    0 ],
           'CORREL'                        => [ 307,    2,    2,    0 ],
           'COVAR'                         => [ 308,    2,    2,    0 ],
           'FORECAST'                      => [ 309,    3,    2,    0 ],
           'FTEST'                         => [ 310,    2,    2,    0 ],
           'INTERCEPT'                     => [ 311,    2,    2,    0 ],
           'PEARSON'                       => [ 312,    2,    2,    0 ],
           'RSQ'                           => [ 313,    2,    2,    0 ],
           'STEYX'                         => [ 314,    2,    2,    0 ],
           'SLOPE'                         => [ 315,    2,    2,    0 ],
           'TTEST'                         => [ 316,    4,    2,    0 ],
           'PROB'                          => [ 317,   -1,    2,    0 ],
           'DEVSQ'                         => [ 318,   -1,    0,    0 ],
           'GEOMEAN'                       => [ 319,   -1,    0,    0 ],
           'HARMEAN'                       => [ 320,   -1,    0,    0 ],
           'SUMSQ'                         => [ 321,   -1,    0,    0 ],
           'KURT'                          => [ 322,   -1,    0,    0 ],
           'SKEW'                          => [ 323,   -1,    0,    0 ],
           'ZTEST'                         => [ 324,   -1,    0,    0 ],
           'LARGE'                         => [ 325,    2,    0,    0 ],
           'SMALL'                         => [ 326,    2,    0,    0 ],
           'QUARTILE'                      => [ 327,    2,    0,    0 ],
           'PERCENTILE'                    => [ 328,    2,    0,    0 ],
           'PERCENTRANK'                   => [ 329,   -1,    0,    0 ],
           'MODE'                          => [ 330,   -1,    2,    0 ],
           'TRIMMEAN'                      => [ 331,    2,    0,    0 ],
           'TINV'                          => [ 332,    2,    1,    0 ],
           'CONCATENATE'                   => [ 336,   -1,    1,    0 ],
           'POWER'                         => [ 337,    2,    1,    0 ],
           'RADIANS'                       => [ 342,    1,    1,    0 ],
           'DEGREES'                       => [ 343,    1,    1,    0 ],
           'SUBTOTAL'                      => [ 344,   -1,    0,    0 ],
           'SUMIF'                         => [ 345,   -1,    0,    0 ],
           'COUNTIF'                       => [ 346,    2,    0,    0 ],
           'COUNTBLANK'                    => [ 347,    1,    0,    0 ],
           'ROMAN'                         => [ 354,   -1,    1,    0 ]
      }
   
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
