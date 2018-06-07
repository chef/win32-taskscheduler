require 'spec_helper'

require 'win32/taskscheduler'
require 'win32/windows/constants'
require 'win32/windows/helper'
require 'win32/windows/time_calc_helper'


RSpec.describe Win32::TaskScheduler, :windows_only  do
  let(:task) { "test_task" }
  let(:application) { "notepad.exe" }
  let(:tsk_time) { Time.now }
  let(:trigger) { { start_year: tsk_time.year, start_month: tsk_time.month, start_day: tsk_time.day,
                    start_hour: tsk_time.hour, start_minute: tsk_time.min } }

  describe 'Ensuring trigger constants' do
    subject(:ts) { Win32::TaskScheduler }
    context 'to handle scheduled tasks' do
      it { should be_const_defined(:ONCE) }
      it { should be_const_defined(:DAILY) }
      it { should be_const_defined(:WEEKLY) }
      it { should be_const_defined(:MONTHLYDATE) }
      it { should be_const_defined(:MONTHLYDOW) }
    end

    context 'to handle other types' do
      it { should be_const_defined(:AT_LOGON) }
      it { should be_const_defined(:AT_SYSTEMSTART) }
      it { should be_const_defined(:ON_IDLE) }
    end
  end

  describe '#tasks' do
    let(:task_scheduler) { Win32::TaskScheduler.new }
    it 'returns an array' do
      expect(task_scheduler.tasks).to be_a(Array)
    end

    it 'returns tasks name post creation' do
      create_task
      expect(task_scheduler.tasks).to include(task)
    end

    it 'is an alias with enum' do
      expect(task_scheduler.tasks).to include(task)
    end

    it 'does not return the task name post deletion' do
      delete_task
      expect(task_scheduler.tasks).not_to include(task)
    end
  end

  describe '#exists?' do
    let(:task_scheduler) { Win32::TaskScheduler.new }
    it 'returns false when a task is not created' do
      expect(task_scheduler.exists?(task)).to be(false)
    end
    it 'returns true when a task is created' do
      create_task
      expect(task_scheduler.exists?(task)).to be(true)
    end
    it 'returns true when a task is deleted' do
      delete_task
      expect(task_scheduler.exists?(task)).to be(false)
    end
  end

  describe '#get_task' do
    let(:task_scheduler) { Win32::TaskScheduler.new }
    it 'requires an argument' do
      expect{ task_scheduler.get_task }.to raise_error(ArgumentError)
    end

    it 'raises an error when a string is not passed' do
      expect{ task_scheduler.get_task(1) }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      expect{ task_scheduler.get_task(task) }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'returns task when it is present' do
      create_task
      expect(task_scheduler.get_task(task)).to be_a(WIN32OLE)
    end

    it 'raises an error post task deletion' do
      delete_task
      expect{ task_scheduler.get_task(task) }.to raise_error(Win32::TaskScheduler::Error)
    end
  end

  describe '#activate' do
    let(:task_scheduler) { Win32::TaskScheduler.new }
    it 'requires an argument' do
      expect{ task_scheduler.activate }.to raise_error(ArgumentError)
    end

    it 'raises an error when a string is not passed' do
      expect{ task_scheduler.activate(1) }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      expect{ task_scheduler.activate(task) }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'enables the specified task' do
      create_task
      expect(task_scheduler.activate(task)).to be_a(WIN32OLE)
      expect(task_scheduler.enabled?).to be(true)
    end

    it 'raises an error post task deletion' do
      delete_task
      expect{ task_scheduler.activate(task) }.to raise_error(Win32::TaskScheduler::Error)
    end
  end

  describe '#delete' do
    let(:task_scheduler) { Win32::TaskScheduler.new }
    it 'requires an argument' do
      expect{ task_scheduler.delete }.to raise_error(ArgumentError)
    end

    it 'raises an error when a string is not passed' do
      expect{ task_scheduler.delete(1) }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      expect{ task_scheduler.delete(task) }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'deletes the specified task' do
      create_task
      expect(delete_task).to be_nil
      expect(task_scheduler.exists?(task)).to be(false)
    end
  end

  def create_task
    trigger[:trigger_type] = Win32::TaskScheduler::ONCE
    task_scheduler.new_work_item(task, trigger)
    task_scheduler.application_name = application
  end

  def delete_task
    task_scheduler.delete(task)
  end
end
