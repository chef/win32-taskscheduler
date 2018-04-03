module Windows
  module TimeCalcHelper

    # Returns actual no of days for given month;
    # Array with a 0 is defined to give actual result without
    # any manipulation. eg, DAY_OF_MONTH[1] = 31
    # 0(NUMBER) is kept to avoid exceptions during calculations
    DAY_OF_MONTH = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

    # Returns no of days in a given month of a year
    def day_in_month(month, year)
      (month == 2 && is_leap_year?(year)) ? 29 : DAY_OF_MONTH[month]
    end

    # Year is leap when it is a multiple of 4 and not a multiple of 100.
    # But it can be a multiple of 400
    def is_leap_year?(year)
      (((year % 4).zero? && !(year % 100).zero?) || (year % 400).zero?)
    end

    # Returns total time duration in minutes
    def time_in_minutes(time_str)
      time_in_seconds(time_str) / 60
    end

    # Calculates total time duration in seconds
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

    # a will contain extra value of low_rank (in high_rank(eg min));
    # b will hold actual low_rank value(ie sec) Example:
    # low_rank = 65, high_rank = 2, div_val = 60
    # Hence a = 1; b = 5
    def extra_time(low_rank, high_rank, div_val)
      a, b = low_rank.divmod(div_val)
      high_rank += a; low_rank = b
      [low_rank, high_rank]
    end

    def extra_months(month_count, year_count, init_month, init_year)
      year, month_count = month_count.divmod(12)
      if year.positive? && month_count.zero?
        month_count = 12
        year -= 1
      end
      year_count += year
      [month_count, year_count]
    end

    # Returns no of actual days with all overloaded months & Years
    def extra_days(days_count, month_count, year_count, init_month, init_year)
      # Will keep increamenting them with surplus days
      days = days_count
      mth = init_month
      yr = init_year

      loop do
       days -= day_in_month(mth, yr)
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

    # Extracts "P_Y_M_DT_H_M_S" format and
    # Returns a hash with applicable values of
    # (keys =>) [:year, :month, :day, :hour, :min, :sec]
    # Example: "PT3S" => {sec: 3}
    def time_details(time_str)
      tm_detail = {}
      if time_str.to_s != ''
        # time_str will be like "PxxYxxMxxDTxxHxxMxxS"
        # Ignoring 'P' and extracting date and time
        dt, tm = time_str[1..-1].split('T')

        # Replacing strings
        if dt.to_s != ''
          dt['Y'] = 'year' if dt['Y']; dt['M'] = 'month' if dt['M']; dt['D'] = 'day' if dt['D']
          dt_tm_array_to_hash(dt, tm_detail)
        end

        if tm.to_s != ''
          tm['H'] = 'hour' if tm['H']; tm['M'] = 'min' if tm['M']; tm['S'] = 'sec' if tm['S']
          dt_tm_array_to_hash(tm, tm_detail)
        end
      end
      tm_detail
    end

    # Method to convert date/time array to hash
    def dt_tm_array_to_hash(arr, tm_detail)
      arr.split(/(\d+)/)[1..-1].each_slice(2).inject(tm_detail) { |h, i| h[i.last.to_sym] = i.first; h }
    end
  end
end
