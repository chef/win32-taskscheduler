require 'spec_helper'

require 'win32/taskscheduler'
require 'win32/windows/constants'
require 'win32/windows/helper'
require 'win32/windows/time_calc_helper'

RSpec.describe Win32::TaskScheduler do

  let(:task) { "test_task" }
  let(:application) { "notepad.exe" }
  let(:tsk_time) { Time.now }
  let(:trigger) { { start_year: tsk_time.year, start_month: tsk_time.month, start_day: tsk_time.day,
                    start_hour: tsk_time.hour, start_minute: tsk_time.min } }
  
  describe "Check constants" do
    subject(:ts) { Win32::TaskScheduler }
    context "to handle scheduled triggers" do
      it { should be_const_defined(:ONCE) }
      it { should be_const_defined(:DAILY) }
      it { should be_const_defined(:WEEKLY) }
      it { should be_const_defined(:MONTHLYDATE) }
      it { should be_const_defined(:MONTHLYDOW) }
    end

    context "to handle other trigger type" do
      it { should be_const_defined(:AT_LOGON) }
      it { should be_const_defined(:AT_SYSTEMSTART) }
      it { should be_const_defined(:ON_IDLE) }
    end
  end

  describe 'One time task' do
    let(:task_scheduler) { Win32::TaskScheduler.new }

    before do 
      trigger[:trigger_type] = Win32::TaskScheduler::ONCE
      task_scheduler.new_work_item(task, trigger)
      task_scheduler.application_name = application
    end

    after { delete_task }

    let(:trigger_details) { task_scheduler.trigger(0) }

    it 'is enabled' do
      task_scheduler.activate(task)
      expect(task_scheduler.exists?(task)).to be true
      expect(task_scheduler.enabled?).to be true
    end

    it 'is working for application' do
      expect(task_scheduler.application_name).to eql(application)
    end

    it 'returns correct trigger information' do
      expect(trigger_details).to be_a(Hash)
      expect(trigger_details[:trigger_type]).to eql(Win32::TaskScheduler::ONCE)
    end

    it 'returns start time' do
      expect(trigger_details[:start_year].to_i).to eql(trigger[:start_year].to_i)
      expect(trigger_details[:start_month].to_i).to eql(trigger[:start_month].to_i)
      expect(trigger_details[:start_day].to_i).to eql(trigger[:start_day].to_i)
      expect(trigger_details[:start_hour].to_i).to eql(trigger[:start_hour].to_i)
      expect(trigger_details[:start_minute].to_i).to eql(trigger[:start_minute].to_i)
    end

    it 'returns end time' do
      expect(trigger_details[:end_year]).to be_nil
      expect(trigger_details[:end_month]).to be_nil
      expect(trigger_details[:end_day]).to be_nil
    end
  end

  def delete_task
    task_scheduler.delete(task)
  end
end
