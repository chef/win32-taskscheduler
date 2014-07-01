require 'ffi'

module Windows
  module Helper
    extend FFI::Library

    ffi_lib :kernel32

    attach_function :FormatMessage, :FormatMessageA,
    [:uint32, :pointer, :uint32, :uint32, :pointer, :uint32, :pointer], :uint32

    def win_error(function, err=FFI.errno)
      flags = 0x00001000 | 0x00000200
      buf = FFI::MemoryPointer.new(:char, 1024)

      FormatMessage(flags, nil, err , 0x0409, buf, 1024, nil)

      function + ': ' + buf.read_string.strip
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
