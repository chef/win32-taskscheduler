require_relative 'helper'

module FFI
  class Pointer
    def read_wstring(num_wchars = nil)
      if num_wchars.nil?
        # Find the length of the string
        length = 0
        last_char = nil
        while last_char != "\000\000"
          length += 1
          last_char = get_bytes(0, length * 2)[-2..-1]
        end

        num_wchars = length
      end

      wide_to_utf8(get_bytes(0, num_wchars * 2))
    end

    def wide_to_utf8(wstring)
      # ensure it is actually UTF-16LE
      # Ruby likes to mark binary data as ASCII-8BIT
      wstring = wstring.force_encoding('UTF-16LE')

      # encode it all as UTF-8 and remove trailing CRLF and NULL characters
      wstring.encode('UTF-8').strip
    end
  end
end

module Win32
  class TaskScheduler
    module SID
      extend Win32::TaskScheduler::Helper
      ERROR_INSUFFICIENT_BUFFER = 122

      def self.LocalSystem
        from_string_sid('S-1-5-18')
      end

      def self.NtLocal
        from_string_sid('S-1-5-19')
      end

      def self.NtNetwork
        from_string_sid('S-1-5-20')
      end

      def self.BuiltinAdministrators
        from_string_sid('S-1-5-32-544')
      end

      def self.BuiltinUsers
        from_string_sid('S-1-5-32-545')
      end

      def self.Guests
        from_string_sid('S-1-5-32-546')
      end

      # Converts a string-format security identifier (SID) into a valid, functional SID
      # and returns a hash with account_name and account_simple_name
      # @see https://docs.microsoft.com/en-us/windows/desktop/api/sddl/nf-sddl-convertstringsidtosidw
      #
      def self.from_string_sid(string_sid)
        result = FFI::MemoryPointer.new :pointer
        unless ConvertStringSidToSidW(utf8_to_wide(string_sid), result)
          raise FFI::LastError.error
        end
        result_pointer = result.read_pointer
        domain, name, use = account(result_pointer)
        LocalFree(result_pointer)
        account_names(domain, name, use)
      end

      def self.utf8_to_wide(ustring)
        # ensure it is actually UTF-8
        # Ruby likes to mark binary data as ASCII-8BIT
        ustring = (ustring + '').force_encoding('UTF-8')

        # ensure we have the double-null termination Windows Wide likes
        ustring += "\000\000" if ustring.empty? || ustring[-1].chr != "\000"

        # encode it all as UTF-16LE AKA Windows Wide Character AKA Windows Unicode
        ustring.encode('UTF-16LE')
      end

      # Accepts a security identifier (SID) as input.
      # It retrieves the name of the account for this SID and the name of the
      # first domain on which this SID is found
      # @see https://docs.microsoft.com/en-us/windows/desktop/api/winbase/nf-winbase-lookupaccountsidw
      #
      def self.account(sid)
        sid = sid.pointer if sid.respond_to?(:pointer)
        name_size = FFI::Buffer.new(:long).write_long(0)
        referenced_domain_name_size = FFI::Buffer.new(:long).write_long(0)

        if LookupAccountSidW(nil, sid, nil, name_size, nil, referenced_domain_name_size, nil)
          raise 'Expected ERROR_INSUFFICIENT_BUFFER from LookupAccountSid, and got no error!'
        elsif FFI::LastError.error != ERROR_INSUFFICIENT_BUFFER
          raise FFI::LastError.error
        end

        name = FFI::MemoryPointer.new :char, (name_size.read_long * 2)
        referenced_domain_name = FFI::MemoryPointer.new :char, (referenced_domain_name_size.read_long * 2)
        use = FFI::Buffer.new(:long).write_long(0)
        unless LookupAccountSidW(nil, sid, name, name_size, referenced_domain_name, referenced_domain_name_size, use)
          raise FFI::LastError.error
        end
        [referenced_domain_name.read_wstring(referenced_domain_name_size.read_long), name.read_wstring(name_size.read_long), use.read_long]
      end

      # Formats domain, name and returns a hash with
      # account_name and account_simple_name
      #
      def self.account_names(domain, name, _use)
        account_name = !domain.to_s.empty? ? "#{domain}\\#{name}" : name
        account_simple_name = name
        { account_name: account_name, account_simple_name: account_simple_name }
      end

      SERVICE_ACCOUNT_USERS = [self.LocalSystem, self.NtLocal, self.NtNetwork].map do |user|
        [user[:account_simple_name].upcase, user[:account_name].upcase]
      end.flatten.freeze

      BUILT_IN_GROUPS = [self.BuiltinAdministrators, self.BuiltinUsers, self.Guests].map do |user|
        [user[:account_simple_name].upcase, user[:account_name].upcase]
      end.flatten.freeze
    end
  end
end
