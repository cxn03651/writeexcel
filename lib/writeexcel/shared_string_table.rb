class Workbook < BIFFWriter
  require 'writeexcel/properties'
  require 'writeexcel/helper'

  class SharedString
    attr_reader :string, :str_id

    def initialize(string, str_id)
      @string, @str_id = string, str_id
    end
  end

  class SharedStringTable
    attr_reader :str_total

    def initialize
      @shared_string_table = []
      @string_to_shared_string = {}
      @str_total = 0
    end

    def has_string?(string)
      !!@string_to_shared_string[string]
    end

    def <<(string)
      @str_total += 1
      unless has_string?(string)
        shared_string = SharedString.new(string, str_unique)
        @shared_string_table << shared_string
        @string_to_shared_string[string] = shared_string
      end
      id(string)
    end

    def strings
      @shared_string_table.collect { |shared_string| shared_string.string }
    end

    def id(string)
      @string_to_shared_string[string].str_id
    end

    def str_unique
      @shared_string_table.size
    end
  end
end
