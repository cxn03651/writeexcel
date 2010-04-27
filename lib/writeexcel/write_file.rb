# -*- coding: utf-8 -*-

class WriteFile
  ###############################################################################
  #
  # _prepend($data)
  #
  # General storage function
  #
  def prepend(*args)
    data = args.collect{ |arg| arg.dup.force_encoding('ASCII-8BIT') }.join
    data = add_continue(data) if data.bytesize > @limit

    @datasize += data.bytesize
    @data      = data + @data

    data
  end

  ###############################################################################
  #
  # _append($data)
  #
  # General storage function
  #
  def append(*args)
    data = args.collect{ |arg| arg.dup.force_encoding('ASCII-8BIT') }.join
    # Add CONTINUE records if necessary
    data = add_continue(data) if data.bytesize > @limit
    if @using_tmpfile
      @filehandle.write data
      @datasize += data.bytesize
    else
      @datasize += data.bytesize
      @data      = @data + data
    end

    data
  end
end
