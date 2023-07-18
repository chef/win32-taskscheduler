require "spec_helper"
require "win32/taskscheduler"
require "win32/taskscheduler/constants"

RSpec.describe Win32::TaskScheduler, :windows_only do
  describe "Ensuring trigger constants" do
    subject(:ts) { Win32::TaskScheduler }
    describe "to handle scheduled tasks" do
      %i{ONCE DAILY WEEKLY MONTHLYDATE MONTHLYDOW}.each do |const|
        it { should be_const_defined(const) }
      end
    end

    describe "to handle other types" do
      %i{AT_LOGON AT_SYSTEMSTART ON_IDLE ON_SESSION_STATE_CHANGE}.each do |const|
        it { should be_const_defined(const) }
      end
    end
  end
end
