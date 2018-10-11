require 'win32/taskscheduler/time_calc_helper'
require 'spec_helper'

RSpec.describe Win32::TaskScheduler::TimeCalcHelper do
  let(:object) { klass.new }
  let(:klass) do
    Class.new do
      include Win32::TaskScheduler::TimeCalcHelper
    end
  end

  describe 'Format Test:' do
    context 'An Invalid Date-Time string' do
      time_str = 'An Invalid String'
      let(:time_details) { object.time_details(time_str) }
      it 'Returns an empty hash' do
        expect(time_details).to be_a(Hash)
        expect(time_details).to be_empty
      end
    end

    context 'A valid Date string' do
      time_str = 'P1Y2M3D'
      let(:time_details) { object.time_details(time_str) }
      it 'Returns a valid Hash with only Date values' do
        expect(time_details).to be_a(Hash)
        expect(time_details[:year]).to eql('1')
        expect(time_details[:month]).to eql('2')
        expect(time_details[:day]).to eql('3')
        expect(time_details[:hour]).to be_nil
        expect(time_details[:min]).to be_nil
        expect(time_details[:sec]).to be_nil
      end
    end

    context 'A valid Time string' do
      time_str = 'PT4H5M6S'
      let(:time_details) { object.time_details(time_str) }
      it 'Returns a valid Hash with only Time values' do
        expect(time_details).to be_a(Hash)
        expect(time_details[:year]).to be_nil
        expect(time_details[:month]).to be_nil
        expect(time_details[:day]).to be_nil
        expect(time_details[:hour]).to eql('4')
        expect(time_details[:min]).to eql('5')
        expect(time_details[:sec]).to eql('6')
      end
    end

    context 'A valid Date-Time string' do
      time_str = 'P1Y2M3DT4H5M6S'
      let(:time_details) { object.time_details(time_str) }
      it 'returns a valid Hash with values' do
        expect(time_details).to be_a(Hash)
        expect(time_details[:year]).to eql('1')
        expect(time_details[:month]).to eql('2')
        expect(time_details[:day]).to eql('3')
        expect(time_details[:hour]).to eql('4')
        expect(time_details[:min]).to eql('5')
        expect(time_details[:sec]).to eql('6')
      end
    end

    context 'An Unformatted Date-Time string' do
      time_str = 'P2M3D1YT6S5M4H'
      let(:time_details) { object.time_details(time_str) }
      it 'Returns a valid Hash with values' do
        expect(time_details).to be_a(Hash)
        expect(time_details[:year]).to eql('1')
        expect(time_details[:month]).to eql('2')
        expect(time_details[:day]).to eql('3')
        expect(time_details[:hour]).to eql('4')
        expect(time_details[:min]).to eql('5')
        expect(time_details[:sec]).to eql('6')
      end
    end

    context 'A Date-Time string with few parameters' do
      time_str = 'P1Y2MT3M4S'
      let(:time_details) { object.time_details(time_str) }
      it 'Returns a valid Hash with selected values' do
        expect(time_details).to be_a(Hash)
        expect(time_details[:year]).to eql('1')
        expect(time_details[:month]).to eql('2')
        expect(time_details[:day]).to be_nil
        expect(time_details[:hour]).to be_nil
        expect(time_details[:min]).to eql('3')
        expect(time_details[:sec]).to eql('4')
      end
    end
  end

  describe 'Year conversion:' do
    # Time.now to be a Non-Leap Year(0001-01-01 00:00:00), So that
    # Tests suites will be independant of their execution Time
    before(:each) do
      allow(Time).to receive(:now).and_return(Time.new(1))
    end

    let(:current_time) { Time.new(1) }
    context 'on given year' do
      time_str = 'P1Y'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected year' do
        expect(next_time.year).to eql(current_time.year + 1)
        expect(next_time.month).to eql(current_time.month)
        expect(next_time.day).to eql(current_time.day)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end

    context 'on given months' do
      time_str = 'P12M'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected year' do
        expect(next_time.year).to eql(current_time.year + 1)
        expect(next_time.month).to eql(current_time.month)
        expect(next_time.day).to eql(current_time.day)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end

    context 'on given days' do
      time_str = 'P365D'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected year' do
        expect(next_time.year).to eql(current_time.year + 1)
        expect(next_time.month).to eql(current_time.month)
        expect(next_time.day).to eql(current_time.day)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end

    context 'by given year and month' do
      time_str = 'P1Y12M'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected year' do
        expect(next_time.year).to eql(current_time.year + 2)
        expect(next_time.month).to eql(current_time.month)
        expect(next_time.day).to eql(current_time.day)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end

    context 'by given year, months and days' do
      time_str = 'P1Y12M365D'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'Returns the actual incremented day' do
        expect(next_time.year).to eql(current_time.year + 3)
        expect(next_time.month).to eql(current_time.month)
        expect(next_time.day).to eql(current_time.day)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end
  end

  describe 'Month conversion:' do
    # Time.now to be a Non-Leap Year(0001-01-01 00:00:00), So that
    # Tests suites will be independant of their execution Time
    before(:each) do
      allow(Time).to receive(:now).and_return(Time.new(1))
    end

    let(:current_time) { Time.new(1) }
    context 'on given month' do
      time_str = 'P1M'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected month' do
        expect(next_time.year).to eql(current_time.year)
        expect(next_time.month).to eql(current_time.month + 1)
        expect(next_time.day).to eql(current_time.day)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end

    context 'on given days' do
      # Since we are running test on 1st month, next month will occur in 31 days
      time_str = 'P31D'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected month' do
        expect(next_time.year).to eql(current_time.year)
        expect(next_time.month).to eql(current_time.month + 1)
        expect(next_time.day).to eql(current_time.day)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end

    context 'by given month and days' do
      # Since we are running test on 1st month, next month will occur in 31 days
      time_str = 'P1M31D'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected month' do
        expect(next_time.year).to eql(current_time.year)
        expect(next_time.month).to eql(current_time.month + 2)
        expect(next_time.day).to eql(current_time.day)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end
  end

  describe 'Day conversion:' do
    # Time.now to be a Non-Leap Year(0001-01-01 00:00:00), So that
    # Tests suites will be independant of their execution Time
    before(:each) do
      allow(Time).to receive(:now).and_return(Time.new(1))
    end

    let(:current_time) { Time.new(1) }
    context 'on given days' do
      time_str = 'P1D'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected days' do
        expect(next_time.year).to eql(current_time.year)
        expect(next_time.month).to eql(current_time.month)
        expect(next_time.day).to eql(current_time.day + 1)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end

    context 'on given hours' do
      time_str = 'PT24H'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected days' do
        expect(next_time.year).to eql(current_time.year)
        expect(next_time.month).to eql(current_time.month)
        expect(next_time.day).to eql(current_time.day + 1)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end

    context 'on given minutes' do
      time_str = 'PT1440M'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected days' do
        expect(next_time.year).to eql(current_time.year)
        expect(next_time.month).to eql(current_time.month)
        expect(next_time.day).to eql(current_time.day + 1)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end

    context 'on given seconds' do
      time_str = 'PT86400S'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected days' do
        expect(next_time.year).to eql(current_time.year)
        expect(next_time.month).to eql(current_time.month)
        expect(next_time.day).to eql(current_time.day + 1)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end

    context 'By the given days, hours, minutes and seconds' do
      time_str = 'P1DT24H1440M86400S'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected days' do
        expect(next_time.year).to eql(current_time.year)
        expect(next_time.month).to eql(current_time.month)
        expect(next_time.day).to eql(current_time.day + 4)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end
  end

  describe 'Minute conversion:' do
    # Time.now to be a Non-Leap Year(0001-01-01 00:00:00), So that
    # Tests suites will be independant of their execution Time
    before(:each) do
      allow(Time).to receive(:now).and_return(Time.new(1))
    end

    let(:current_time) { Time.new(1) }
    context 'on given minute' do
      time_str = 'PT1M'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected minutes' do
        expect(next_time.year).to eql(current_time.year)
        expect(next_time.month).to eql(current_time.month)
        expect(next_time.day).to eql(current_time.day)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min + 1)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end

    context 'on given seconds' do
      time_str = 'PT60S'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected minutes' do
        expect(next_time.year).to eql(current_time.year)
        expect(next_time.month).to eql(current_time.month)
        expect(next_time.day).to eql(current_time.day)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min + 1)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end

    context 'By given minutes and seconds' do
      time_str = 'PT1M60S'
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected minutes' do
        expect(next_time.year).to eql(current_time.year)
        expect(next_time.month).to eql(current_time.month)
        expect(next_time.day).to eql(current_time.day)
        expect(next_time.hour).to eql(current_time.hour)
        expect(next_time.min).to eql(current_time.min + 2)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end
  end

  describe 'Date-Time conversions:' do
    # For a single string given in leap/non-leap year
    time_str = 'P400Y500M600DT700H800M900S'

    context 'In a non leap year' do
      let(:current_time) { Time.new(1) }
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected values' do
        allow(Time).to receive(:now).and_return(Time.new(1))

        expect(next_time.year).to eql(current_time.year + 443)
        expect(next_time.month).to eql(current_time.month + 4)
        expect(next_time.day).to eql(current_time.day + 21)
        expect(next_time.hour).to eql(current_time.hour + 17)
        expect(next_time.min).to eql(current_time.min + 35)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end

    context 'In a leap year' do
      let(:current_time) { Time.new(0) }
      let(:time_seconds) { object.time_in_seconds(time_str) }
      let(:next_time) { current_time + time_seconds }
      it 'returns expected values' do
        allow(Time).to receive(:now).and_return(Time.new(0))

        # 1 Day less for the same string
        expect(next_time.year).to eql(current_time.year + 443)
        expect(next_time.month).to eql(current_time.month + 4)
        expect(next_time.day).to eql(current_time.day + 20)
        expect(next_time.hour).to eql(current_time.hour + 17)
        expect(next_time.min).to eql(current_time.min + 35)
        expect(next_time.sec).to eql(current_time.sec)
      end
    end
  end
end
