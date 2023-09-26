module Win32
  class TaskScheduler
    module TimeCalcHelper

      # No of days in given month. Deliberately placed 0 in the
      # beginning to avoid any miscalculations
      #
      DAYS_IN_A_MONTH = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31].freeze

      # No of days in a month for given year
      #
      # @param [Integer] month
      # @param [Integer] year
      # @return [Integer] No of days
      #
      def days_in_month(month, year)
        month == 2 && is_leap_year?(year) ? 29 : DAYS_IN_A_MONTH[month]
      end

      # Checks weather the given year is a leap year or not
      #
      # @param [Integer] year
      # @return [Boolean]
      #
      def is_leap_year?(year)
        (((year % 4).zero? && !(year % 100).zero?) || (year % 400).zero?)
      end

      # Calculates the total minutes within given PTM format time
      #
      # @param [Integer] time_str
      # @return [Integer] Time duration in minutes
      #
      def time_in_minutes(time_str)
        time_in_seconds(time_str) / 60
      end

      # Calculates the total seconds within given PTM format time
      #
      # @param [Integer] time_str the time in PTM format
      # @return [Integer] Time duration in seconds
      #
      def time_in_seconds(time_str)
        dt_tm_hash = time_details(time_str)
        curr_time = Time.now

        # Basic time variables
        future_year = curr_time.year + dt_tm_hash[:year].to_i
        future_month = curr_time.month + dt_tm_hash[:month].to_i
        future_day = curr_time.day + dt_tm_hash[:day].to_i
        future_hr = curr_time.hour + dt_tm_hash[:hour].to_i
        future_min = curr_time.min + dt_tm_hash[:min].to_i
        future_sec = curr_time.sec + dt_tm_hash[:sec].to_i

        # 'extra value' calculations for these time variables
        future_sec, future_min = extra_time(future_sec, future_min, 60)
        future_min, future_hr = extra_time(future_min, future_hr, 60)
        future_hr, future_day = extra_time(future_hr, future_day, 24)

        # explicit method to calculate overloaded days;
        # They may stretch upto years; heance leap year & months are into consideration
        future_day, future_month, future_year = extra_days(future_day, future_month, future_year, curr_time.month, curr_time.year)

        future_month, future_year = extra_months(future_month, future_year, curr_time.month, curr_time.year)

        future_time = Time.new(future_year, future_month, future_day, future_hr, future_min, future_sec)

        # Difference in time will return seconds
        future_time.to_i - curr_time.to_i
      end

      # Adjusts the overlapping seconds and returns actual minutes and seconds
      #
      # @example
      #   extra_time(65, 2, 60) #=> => [5, 3]
      #
      def extra_time(low_rank, high_rank, div_val)
        a, b = low_rank.divmod(div_val)
        high_rank += a; low_rank = b
        [low_rank, high_rank]
      end

      # Adjusts the overlapping months and returns actual month and year
      #
      def extra_months(month_count, year_count, _init_month, _init_year)
        year, month_count = month_count.divmod(12)
        if year.positive? && month_count.zero?
          month_count = 12
          year -= 1
        end
        year_count += year
        [month_count, year_count]
      end

      # Adjusts the overlapping years and months and returns actual days, month and year
      #
      def extra_days(days_count, month_count, year_count, init_month, init_year)
        # Will keep increamenting them with surplus days
        days = days_count
        mth = init_month
        yr = init_year

        loop do
          days -= days_in_month(mth, yr)
          break if days <= 0

          mth += 1
          if mth > 12
            mth = 1; yr += 1
          end
          days_count = days
        end

        # Setting actual incremented values
        month_count += (mth - init_month)
        year_count += (yr - init_year)

        [days_count, month_count, year_count]
      end

      # Extracts a hash out of given PTM formatted time
      #
      # @param [String] time_str
      # @return [Hash<:year, :month, :day, :hour, :min, :sec>] With their values in Integer
      #
      # @example
      #   time_details("PT3S") #=> {sec: 3}
      #
      def time_details(time_str)
        tm_detail = {}
        if time_str.to_s != ""
          # time_str will be like "PxxYxxMxxDTxxHxxMxxS"
          # Ignoring 'P' and extracting date and time
          dt, tm = time_str[1..-1].split("T")

          # Replacing strings
          if dt.to_s != ""
            dt["Y"] = "year" if dt["Y"]; dt["M"] = "month" if dt["M"]; dt["D"] = "day" if dt["D"]
            dt_tm_string_to_hash(dt, tm_detail)
          end

          if tm.to_s != ""
            tm["H"] = "hour" if tm["H"]; tm["M"] = "min" if tm["M"]; tm["S"] = "sec" if tm["S"]
            dt_tm_string_to_hash(tm, tm_detail)
          end
        end
        tm_detail
      end

      # Converts the given date/time string to the hash
      #
      # @param [String] str
      # @param [Hash] tm_detail May be loaded
      # @return [Hash]
      #
      # @example
      #   dt_tm_string_to_hash("10year3month", {}) #=> {:year=>"10", :month=>"3"}
      #
      def dt_tm_string_to_hash(str, tm_detail)
        str.split(/(\d+)/)[1..-1].each_slice(2).each_with_object(tm_detail) { |i, h| h[i.last.to_sym] = i.first; }
      end
    end
  end
end
