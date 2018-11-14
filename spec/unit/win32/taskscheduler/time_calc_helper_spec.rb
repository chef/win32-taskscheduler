require "spec_helper"
require "win32/taskscheduler/time_calc_helper"

RSpec.describe Win32::TaskScheduler::TimeCalcHelper do
  let(:object) { klass.new }
  let(:klass) do
    Class.new do
      include Win32::TaskScheduler::TimeCalcHelper
    end
  end

  describe "#is_leap_year?" do
    it "require an year in integer format" do
      expect { object.is_leap_year? }.to raise_error(ArgumentError)
      expect { object.is_leap_year?("year") }.to raise_error(NoMethodError)
    end

    it "returns true for leap years" do
      year = 2000
      expect(object.is_leap_year?(year)).to be_truthy
      year = 2004
      expect(object.is_leap_year?(year)).to be_truthy
    end

    it "returns false for non-leap years" do
      year = 1900
      expect(object.is_leap_year?(year)).to be_falsy
      year = 2001
      expect(object.is_leap_year?(year)).to be_falsy
    end
  end

  describe "#days_in_month" do
    it "require month and year in integer format" do
      expect { object.days_in_month }.to raise_error(ArgumentError)
      expect { object.days_in_month("month", "year") }.to raise_error(TypeError)
    end

    context "leap year" do
      year = 2000
      it "January will have 31 days" do
        month = 1
        expect(object.days_in_month(month, year)).to eql(31)
      end

      it "February will have 29 days" do
        month = 2
        expect(object.days_in_month(month, year)).to eql(29)
      end

      it "November will have 30 days" do
        month = 11
        expect(object.days_in_month(month, year)).to eql(30)
      end
    end

    context "non-leap year" do
      year = 2003
      it "January will have 31 days" do
        month = 1
        expect(object.days_in_month(month, year)).to eql(31)
      end

      it "February will have 28 days" do
        month = 2
        expect(object.days_in_month(month, year)).to eql(28)
      end

      it "November will have 30 days" do
        month = 11
        expect(object.days_in_month(month, year)).to eql(30)
      end
    end
  end

  describe "#time_details" do
    it "require a time string in String format" do
      expect { object.time_details }.to raise_error(ArgumentError)
      expect { object.time_details(1234) }.to raise_error(TypeError)
    end

    it "returns an empty hash if no string is passed" do
      expect(object.time_details(nil)).to be_a(Hash)
      expect(object.time_details(nil)).to be_empty
    end

    context "of MSDN time string" do
      it "returns hash with date" do
        time_str = "P10Y10M10D"
        time_hsh = { year: "10", month: "10", day: "10" }
        expect(object.time_details(time_str)).to eq(time_hsh)
      end

      it "returns hash with time" do
        time_str = "PT10H10M10S"
        time_hsh = { hour: "10", min: "10", sec: "10" }
        expect(object.time_details(time_str)).to eq(time_hsh)
      end

      it "returns hash with date-time" do
        time_str = "P10Y10M10DT10H10M10S"
        time_hsh = { year: "10", month: "10", day: "10",
                     hour: "10", min: "10", sec: "10" }
        expect(object.time_details(time_str)).to eq(time_hsh)
      end
    end
  end

  describe "#time_in_seconds" do
    it "require a time string in String format" do
      expect { object.time_in_seconds }.to raise_error(ArgumentError)
      expect { object.time_in_seconds(1234) }.to raise_error(TypeError)
    end

    it "returns zero if no string is passed" do
      expect(object.time_in_seconds(nil)).to be_zero
    end

    context "in leap year" do
      before do
        time_now = Time.new(2004)
        allow(Time).to receive(:now).and_return(time_now)
      end

      it "returns seconds for date" do
        time_str = "P10Y10M10D"
        expect(object.time_in_seconds(time_str)).to eq(342_748_800)
      end

      it "returns seconds for time" do
        time_str = "PT10H10M10S"
        expect(object.time_in_seconds(time_str)).to eq(36_610)
      end

      it "returns seconds for date-time" do
        time_str = "P10Y10M10DT10H10M10S"
        expect(object.time_in_seconds(time_str)).to eq(342_785_410)
      end
    end
    context "in non-leap year" do
      before do
        time_now = Time.new(2003)
        allow(Time).to receive(:now).and_return(time_now)
      end

      it "returns seconds for date" do
        time_str = "P10Y10M10D"
        expect(object.time_in_seconds(time_str)).to eq(342_748_800)
      end

      it "returns seconds for time" do
        time_str = "PT10H10M10S"
        expect(object.time_in_seconds(time_str)).to eq(36_610)
      end

      it "returns seconds for date-time" do
        time_str = "P10Y10M10DT10H10M10S"
        expect(object.time_in_seconds(time_str)).to eq(342_785_410)
      end
    end
  end

  describe "#time_in_minutes" do
    it "require a time string in String format" do
      expect { object.time_in_minutes }.to raise_error(ArgumentError)
      expect { object.time_in_minutes(1234) }.to raise_error(TypeError)
    end

    it "returns zero if no string is passed" do
      expect(object.time_in_minutes(nil)).to be_zero
    end

    context "in leap year" do
      before do
        time_now = Time.new(2004)
        allow(Time).to receive(:now).and_return(time_now)
      end

      it "returns minutes for date" do
        time_str = "P10Y10M10D"
        expect(object.time_in_minutes(time_str)).to eq(5_712_480)
      end

      it "returns minutes for time" do
        time_str = "PT10H10M10S"
        expect(object.time_in_minutes(time_str)).to eq(610)
      end

      it "returns minutes for date-time" do
        time_str = "P10Y10M10DT10H10M10S"
        expect(object.time_in_minutes(time_str)).to eq(5_713_090)
      end
    end

    context "in non-leap year" do
      before do
        time_now = Time.new(2003)
        allow(Time).to receive(:now).and_return(time_now)
      end

      it "returns minutes for date" do
        time_str = "P10Y10M10D"
        expect(object.time_in_minutes(time_str)).to eq(5_712_480)
      end

      it "returns minutes for time" do
        time_str = "PT10H10M10S"
        expect(object.time_in_minutes(time_str)).to eq(610)
      end

      it "returns minutes for date-time" do
        time_str = "P10Y10M10DT10H10M10S"
        expect(object.time_in_minutes(time_str)).to eq(5_713_090)
      end
    end
  end
end
