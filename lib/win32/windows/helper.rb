require 'ffi'

module Windows
  module Helper
    extend FFI::Library

    ffi_lib :kernel32

    attach_function :FormatMessage, :FormatMessageA,
    [:uint32, :pointer, :uint32, :uint32, :pointer, :uint32, :pointer], :uint32

    FORMAT_MESSAGE_IGNORE_INSERTS    = 0x00000200
    FORMAT_MESSAGE_FROM_SYSTEM       = 0x00001000

    def win_error(function, err=FFI.errno)
      error_msg = ''
      flags = FORMAT_MESSAGE_FROM_SYSTEM |
              FORMAT_MESSAGE_IGNORE_INSERTS
      FFI::MemoryPointer.new(:char, 1024) do |buf|

        FormatMessage(flags, nil, err , 0x0409, buf, 1024, nil)

        error_msg = function + ': ' + buf.read_string.strip
      end

      error_msg
    end

    def ole_error(function, err)
      regex = /OLE error code:(.*?)\sin/
      match = regex.match(err.to_s)

      if match
        error = match.captures.first.hex
        win_error(function, error)
      else
        msg
      end
    end
  end
end
