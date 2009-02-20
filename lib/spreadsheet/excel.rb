$LOAD_PATH.unshift(File.dirname(__FILE__))

require "biffwriter"
require "olewriter"
require "workbook"
require "worksheet"
require "format"

module Spreadsheet
   class Excel < Workbook
      VERSION = "0.3.5.1"
   end
end
