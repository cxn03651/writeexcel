# -*- coding: utf-8 -*-

class WriteFile
  ###############################################################################
  #
  # _prepend($data)
  #
  # General storage function
  #
  def prepend(*args)
    data = join_data(args)
    @data = data + @data

    data
  end

  ###############################################################################
  #
  # _append($data)
  #
  # General storage function
  #
  def append(*args)
    data = join_data(args)

    if @using_tmpfile
      @filehandle.write(data)
    else
      @data += data
    end

    data
  end

  private

  def join_data(args)
    data =
      ruby_18 { args.join } ||
      ruby_19 { args.collect{ |arg| arg.dup.force_encoding('ASCII-8BIT') }.join }
    # Add CONTINUE records if necessary
    data = add_continue(data) if data.bytesize > @limit

    @datasize += data.bytesize

    data
  end
end
