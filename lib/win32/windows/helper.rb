require 'ffi'

module Windows
  module TaskSchedulerHelper
    extend FFI::Library

    FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200
    FORMAT_MESSAGE_FROM_SYSTEM    = 0x00001000
    FORMAT_MESSAGE_MAX_WIDTH_MASK = 0x000000FF

    ffi_lib :kernel32

    attach_function :FormatMessage, :FormatMessageA,
    [:ulong, :pointer, :ulong, :ulong, :pointer, :ulong, :pointer], :ulong

    def win_error(function, err=FFI.errno)
      err_msg = ''
      flags = FORMAT_MESSAGE_IGNORE_INSERTS |
        FORMAT_MESSAGE_FROM_SYSTEM |
        FORMAT_MESSAGE_MAX_WIDTH_MASK

      # 0x0409 == MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US)
      # We use English for errors because Ruby uses English for errors.

      FFI::MemoryPointer.new(:char, 1024) do |buf|
        len = FormatMessage(flags, nil, err , 0x0409, buf, buf.size, nil)
        err_msg = function + ': ' + buf.read_string(len).strip
      end

      err_msg
    end

    def ole_error(function, err)
      regex = /OLE error code:(.*?)\sin/
      match = regex.match(err.to_s)

      if match
        error = match.captures.first.hex
        win_error(function, error)
      else
        "#{function}: #{err.to_s}"
      end
    end
  end
end
