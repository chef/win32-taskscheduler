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

  describe '#enabled?' do
    let(:task_scheduler) { Win32::TaskScheduler.new }    
    it 'raises an error when a task is not found' do
      expect{ task_scheduler.enabled? }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'returns true for enabled task' do
      create_task
      expect(task_scheduler.enabled?).to eql(true)
      delete_task
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

  describe '#run' do
    # Need to check this method. Not wokring for win10
  end

  describe '#terminate' do
    # Need to check this method. Not wokring for win10
  end

  describe '#application_name' do
    let(:task_scheduler) { Win32::TaskScheduler.new }
    let(:test_app){'cmd.exe'}
    it 'setter raises an error when a string is not passed' do
      expect{ task_scheduler.application_name=(1) }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      expect{ task_scheduler.application_name=('app') }.to raise_error(Win32::TaskScheduler::Error)
      expect{ task_scheduler.application_name }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter raises an error if no active tasks' do
      expect{ task_scheduler.application_name }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter and setter works' do
      create_task
      expect(task_scheduler.application_name=(test_app)).to eql(test_app)
      expect(task_scheduler.application_name).to eql(test_app)
      delete_task
    end
  end

  describe '#parameters' do
    let(:task_scheduler) { Win32::TaskScheduler.new }
    let(:test_app){'cmd1.exe'}
    it 'setter raises an error when a string is not passed' do
      expect{ task_scheduler.parameters=(1) }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      expect{ task_scheduler.parameters=('app') }.to raise_error(Win32::TaskScheduler::Error)
      expect{ task_scheduler.parameters }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter raises an error if no active tasks' do
      expect{ task_scheduler.parameters }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter and setter works' do
      create_task
      expect(task_scheduler.parameters=(test_app)).to eql(test_app)
      expect(task_scheduler.parameters).to eql(test_app)
      delete_task
    end
  end

  describe '#working_directory' do
    let(:task_scheduler) { Win32::TaskScheduler.new }
    let(:test_dir){ Dir.pwd }
    it 'setter raises an error when a string is not passed' do
      expect{ task_scheduler.working_directory=(1) }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      expect{ task_scheduler.working_directory=('app') }.to raise_error(Win32::TaskScheduler::Error)
      expect{ task_scheduler.working_directory }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter raises an error if no active tasks' do
      expect{ task_scheduler.working_directory }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter and setter works' do
      create_task
      expect(task_scheduler.working_directory=(test_dir)).to eql(test_dir)
      expect(task_scheduler.working_directory).to eql(test_dir)
      delete_task
    end
  end

  describe '#priority' do
    let(:task_scheduler) { Win32::TaskScheduler.new }
    let(:priority_val){ Win32::TaskScheduler::HIGH_PRIORITY_CLASS }
    it 'setter raises an error when an integer is not passed' do
      expect{ task_scheduler.priority=('highest') }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      expect{ task_scheduler.priority=(priority_val) }.to raise_error(Win32::TaskScheduler::Error)
      expect{ task_scheduler.priority }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter raises an error if no active tasks' do
      expect{ task_scheduler.priority }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter and setter works' do
      create_task
      expect(task_scheduler.priority=(priority_val)).to eql(priority_val)
      expect(task_scheduler.priority).to eql('highest')
      delete_task
    end
  end

  describe '#comment' do
    let(:task_scheduler) { Win32::TaskScheduler.new }
    let(:comment_val){ 'Test Comment' }
    it 'setter raises an error when a String is not passed' do
      expect{ task_scheduler.comment=(1) }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      expect{ task_scheduler.comment=(comment_val) }.to raise_error(Win32::TaskScheduler::Error)
      expect{ task_scheduler.comment }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter raises an error if no active tasks' do
      expect{ task_scheduler.comment }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter and setter works' do
      create_task
      expect(task_scheduler.comment=(comment_val)).to eql(comment_val)
      expect(task_scheduler.comment).to eql(comment_val)
      delete_task
    end

    it 'is an alias for description' do
      create_task
      expect(task_scheduler.description=(comment_val)).to eql(comment_val)
      expect(task_scheduler.description).to eql(comment_val)
      delete_task
    end
  end


  describe '#max_run_time' do
    let(:task_scheduler) { Win32::TaskScheduler.new }
    let(:max_run_time_val){ 1244145000000 } # Just a random time in miliseconds
    it 'setter raises an error when an integer is not passed' do
      expect{ task_scheduler.max_run_time=('time') }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      expect{ task_scheduler.max_run_time=(max_run_time_val) }.to raise_error(Win32::TaskScheduler::Error)
      expect{ task_scheduler.max_run_time }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter raises an error if no active tasks' do
      expect{ task_scheduler.max_run_time }.to raise_error(Win32::TaskScheduler::Error)
    end

    # TODO: max_run_time setter expects time in miliseconds(may be an overhead),
    # and parser of getter is not working. It may be required to take a look again
    # Well, getter may implement `time_in_seconds` of TimeCalcHelper
    it 'getter and setter works' do
      create_task
      expect(task_scheduler.max_run_time=(max_run_time_val)).to eql(max_run_time_val)
      expect(task_scheduler.max_run_time).to eql(max_run_time_val)
      delete_task
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
