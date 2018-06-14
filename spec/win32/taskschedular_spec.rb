require 'spec_helper'

require 'win32/taskscheduler'
require 'win32/windows/constants'
require 'win32/windows/helper'
require 'win32/windows/time_calc_helper'


RSpec.describe Win32::TaskScheduler, :windows_only  do
  let(:application) { "notepad.exe" }
  let(:tsk_time) { Time.now }
  let(:dummy_task_name) { "test_task_dummy" }
  let(:trigger) { { start_year: tsk_time.year, start_month: tsk_time.month, start_day: tsk_time.day,
                    start_hour: tsk_time.hour, start_minute: tsk_time.min } }

  describe 'Ensuring trigger constants' do
    subject(:ts) { Win32::TaskScheduler }
    describe 'to handle scheduled tasks' do
      it { should be_const_defined(:ONCE) }
      it { should be_const_defined(:DAILY) }
      it { should be_const_defined(:WEEKLY) }
      it { should be_const_defined(:MONTHLYDATE) }
      it { should be_const_defined(:MONTHLYDOW) }
    end

    describe 'to handle other types' do
      it { should be_const_defined(:AT_LOGON) }
      it { should be_const_defined(:AT_SYSTEMSTART) }
      it { should be_const_defined(:ON_IDLE) }
    end
  end

  describe '#constructor' do
    let(:ts) { Win32::TaskScheduler }
    let(:task) { "test_task" }
    let(:folder) { '\\' }
    let(:force) { false }
    before { trigger[:trigger_type] = Win32::TaskScheduler::ONCE }
    context 'argument' do
      context 'requires upto 4 parameters' do
        it 'runs with no arguments' do
          expect(ts.new).to be_a(ts)
        end
        it 'runs with an argument: task' do
          expect(ts.new(task)).to be_a(ts)
        end
        it 'raises error when with two arguments: task, trigger and trigger_type is blank' do
          trigger[:trigger_type] = nil
          expect{ ts.new(task, trigger) }.to raise_error(ArgumentError)
        end
        it 'runs with two arguments: task, trigger and trigger_type is not blank' do
          expect(ts.new(task, trigger)).to be_a(ts)
          delete_task
        end
        it 'runs with three arguments: task, trigger, folder' do
          expect(ts.new(task, trigger, folder)).to be_a(ts)
          delete_task
        end
        it 'runs with four arguments: task, trigger, folder, force' do
          expect(ts.new(task, trigger, folder, force)).to be_a(ts)
          delete_task
        end
        it 'raises an error for more than four arguments' do
          expect{ ts.new(task, trigger, folder, force, 1) }.to raise_error(ArgumentError)
          expect{ ts.new(task, trigger, folder, force, 'abc') }.to raise_error(ArgumentError)
        end
      end
      context 'task' do
        it 'default value is nil' do
          task_scheduler = ts.new(task)
          expect(task_scheduler).to be_a(ts)
        end
        it 'does not create any task if trigger is absent' do
          task_scheduler = ts.new(task)
          expect(task_scheduler).to be_a(ts)
          expect(task_scheduler.exists?(task)).to be(false)
        end
        it 'can create a task if trigger is present' do
          task_scheduler = ts.new(task, trigger)
          expect(task_scheduler).to be_a(ts)
          expect(task_scheduler.exists?(task)).to be(true)
          delete_task
          expect(task_scheduler.exists?(task)).to be(false)
        end
      end
      context 'trigger' do
        it 'default value is nil' do
          task_scheduler = ts.new(task)
          expect(task_scheduler).to be_a(ts)
        end
        it 'can create a task if task argument is present' do
          task_scheduler = ts.new(task, trigger)
          expect(task_scheduler).to be_a(ts)
          expect(task_scheduler.exists?(task)).to be(true)
          delete_task
          expect(task_scheduler.exists?(task)).to be(false)
        end
        it 'does not create any task if task argument is absent' do
          task_scheduler = ts.new(nil, trigger)
          expect(task_scheduler).to be_a(ts)
          expect(task_scheduler.exists?(task)).to be(false)
        end
      end
      context 'folder' do
        let(:service) { WIN32OLE.new('Schedule.Service') }
        let(:root_path) { "\\" }
        let(:test_path) { "\\Foo" }
        before { service.Connect }
        it 'default value is root' do
          task_scheduler = ts.new(task, trigger)
          expect(service.GetFolder(root_path).GetTask(task)).to be_a(WIN32OLE)
          delete_task
        end
        it 'raises an error if path separators (\\\) are absent' do
          some_invalid_path = "Test"
          expect{ ts.new(task, trigger, some_invalid_path) }.to raise_error(ArgumentError)
        end
        context 'when existing folder is given' do
          before do
            @root = service.GetFolder(root_path)
            @root.CreateFolder(test_path)
            @folder = service.GetFolder(test_path)
            @count = @folder.GetTasks(0).Count
          end
          after do
            @folder.DeleteTask(task,0) if !@count.zero?
            @root.DeleteFolder(test_path, 0)
          end
          it 'creates the task at specified path' do
            task_scheduler = ts.new(task, trigger, test_path)
            @count = @folder.GetTasks(0).Count
            expect(task_scheduler).to be_a(ts)
            expect(@folder.GetTask(task)).to be_a(WIN32OLE)
            expect(@count).not_to be_zero
          end
          it 'creates the task at specified path when force is true' do
            task_scheduler = ts.new(task, trigger, test_path, true)
            @count = @folder.GetTasks(0).Count
            expect(task_scheduler).to be_a(ts)
            expect(@folder.GetTask(task)).to be_a(WIN32OLE)
            expect(@count).not_to be_zero
          end
          it 'creates the task at specified path when force is false' do
            task_scheduler = ts.new(task, trigger, test_path, false)
            @count = @folder.GetTasks(0).Count
            expect(task_scheduler).to be_a(ts)
            expect(@folder.GetTask(task)).to be_a(WIN32OLE)
            expect(@count).not_to be_zero
          end
        end
        context 'when folder does not exists' do
          it 'raises an error when force is not given' do
            expect{ ts.new(task, trigger, test_path) }.to raise_error(ArgumentError)
          end
          it 'raises an error when force is false' do
            expect{ ts.new(task, trigger, test_path, false) }.to raise_error(ArgumentError)
          end
          it 'creates a task with a folder specified when force is true' do
            task_scheduler = ts.new(task, trigger, test_path, true)
            root = service.GetFolder(root_path)
            folder = service.GetFolder(test_path)
            task_count = folder.GetTasks(0).Count
            folder_count = folder.GetFolders(0).Count
            expect(task_scheduler).to be_a(ts)
            expect(folder.GetTask(task)).to be_a(WIN32OLE)
            expect(task_count).not_to be_zero
            expect(folder_count).to be_zero
            folder.DeleteTask(task,0)
            root.DeleteFolder(test_path, 0)
          end
          it 'does not create any extra folder/tasks when force is true without update' do
            nested_path = "\\Foo\\Bar"
            task_scheduler = ts.new(task, trigger, nested_path, true)
            root = service.GetFolder(root_path)
            folder = service.GetFolder(nested_path)
            task_count = folder.GetTasks(0).Count
            folder_count = folder.GetFolders(0).Count
            expect(task_scheduler).to be_a(ts)
            expect(folder.GetTask(task)).to be_a(WIN32OLE)
            expect(task_count).not_to be_zero
            expect(folder_count).to be_zero
            folder.DeleteTask(task,0)
            root.DeleteFolder(nested_path, 0)
            root.DeleteFolder(test_path, 0)
          end
          it 'does not create any extra folder/tasks when force is true with update' do
            skip 'Skipping since extra nested folders are getting created'
            nested_path = "\\Foo\\Bar"
            task_scheduler = ts.new(task, trigger, nested_path, true)
            task_scheduler.max_run_time= 10000
            root = service.GetFolder(root_path)
            folder = service.GetFolder(nested_path)
            task_count = folder.GetTasks(0).Count
            folder_count = folder.GetFolders(0).Count
            expect(task_scheduler).to be_a(ts)
            expect(folder.GetTask(task)).to be_a(WIN32OLE)
            expect(task_count).not_to be_zero
            expect(folder_count).to be_zero
            folder.DeleteTask(task,0)
            root.DeleteFolder(nested_path, 0)
            root.DeleteFolder(test_path, 0)
          end
        end
      end
    end
  end

  before(:context) do
    @task = "test_task"
    @task_scheduler = Win32::TaskScheduler.new
  end

  # Cleanup TaskScheduler object
  after(:context) do
    delete_task
  end

  context '#tasks' do
    it 'returns an array' do
      expect(@task_scheduler.tasks).to be_a(Array)
    end

    it 'returns tasks name post creation' do
      create_task
      expect(@task_scheduler.tasks).to include(@task)
    end

    it 'is an alias with enum' do
      expect(@task_scheduler.tasks).to include(@task)
    end

    it 'does not return the task name post deletion' do
      delete_task
      expect(@task_scheduler.tasks).not_to include(@task)
    end
  end

  context '#exists?' do
    it 'returns false when a task is not created' do
      expect(@task_scheduler.exists?(@task)).to be(false)
    end
    it 'returns true when a task is created' do
      create_task
      expect(@task_scheduler.exists?(@task)).to be(true)
    end
    it 'returns true when a task is deleted' do
      delete_task
      expect(@task_scheduler.exists?(@task)).to be(false)
    end
  end

  context '#get_task' do
    it 'requires an argument' do
      expect{ @task_scheduler.get_task }.to raise_error(ArgumentError)
    end

    it 'raises an error when a string is not passed' do
      expect{ @task_scheduler.get_task(1) }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      expect{ @task_scheduler.get_task(@task) }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'returns task when it is present' do
      create_task
      expect(@task_scheduler.get_task(@task)).to be_a(WIN32OLE)
    end

    it 'raises an error post task deletion' do
      delete_task
      expect{ @task_scheduler.get_task(@task) }.to raise_error(Win32::TaskScheduler::Error)
    end
  end

  context '#enabled?' do
    it 'raises an error when a task is not found' do
      @task_scheduler.instance_variable_set(:@task, nil)
      expect{ @task_scheduler.enabled? }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'returns true for enabled task' do
      create_task
      expect(@task_scheduler.enabled?).to eql(true)
    end  
  end

  context '#activate' do
    it 'requires an argument' do
      expect{ @task_scheduler.activate }.to raise_error(ArgumentError)
    end

    it 'raises an error when a string is not passed' do
      expect{ @task_scheduler.activate(1) }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      expect{ @task_scheduler.activate(dummy_task_name) }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'enables the specified task' do
      create_task
      expect(@task_scheduler.activate(@task)).to be_a(WIN32OLE)
      expect(@task_scheduler.enabled?).to be(true)
    end

    it 'raises an error post task deletion' do
      delete_task
      expect{ @task_scheduler.activate(@task) }.to raise_error(Win32::TaskScheduler::Error)
    end
  end

  context '#delete' do
    it 'requires an argument' do
      expect{ @task_scheduler.delete }.to raise_error(ArgumentError)
    end

    it 'raises an error when a string is not passed' do
      expect{ @task_scheduler.delete(1) }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      expect{ @task_scheduler.delete(@task) }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'deletes the specified task' do
      create_task
      expect(delete_task).to be_nil
      expect(@task_scheduler.exists?(@task)).to be(false)
    end
  end

  context '#run' do
    # Need to check this method. Not wokring for win10
  end

  context '#terminate' do
    # Need to check this method. Not wokring for win10
  end

  context '#application_name' do
    let(:test_app){'cmd.exe'}
    it 'setter raises an error when a string is not passed' do
      expect{ @task_scheduler.application_name=(1) }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      @task_scheduler.instance_variable_set(:@task, nil)
      expect{ @task_scheduler.application_name=('app') }.to raise_error(Win32::TaskScheduler::Error)
      expect{ @task_scheduler.application_name }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter raises an error if no active tasks' do
      expect{ @task_scheduler.application_name }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter and setter are working successfully' do
      create_task
      expect(@task_scheduler.application_name=(test_app)).to eql(test_app)
      expect(@task_scheduler.application_name).to eql(test_app)
    end
  end

  context '#parameters' do
    let(:test_app){'cmd1.exe'}
    it 'setter raises an error when a string is not passed' do
      expect{ @task_scheduler.parameters=(1) }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      @task_scheduler.instance_variable_set(:@task, nil)
      expect{ @task_scheduler.parameters=('app') }.to raise_error(Win32::TaskScheduler::Error)
      expect{ @task_scheduler.parameters }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter raises an error if no active tasks' do
      expect{ @task_scheduler.parameters }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter and setter are working successfully' do
      create_task
      expect(@task_scheduler.parameters=(test_app)).to eql(test_app)
      expect(@task_scheduler.parameters).to eql(test_app)
    end
  end

  context '#working_directory' do
    let(:test_dir){ Dir.pwd }
    it 'setter raises an error when a string is not passed' do
      expect{ @task_scheduler.working_directory=(1) }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      @task_scheduler.instance_variable_set(:@task, nil)
      expect{ @task_scheduler.working_directory=('app') }.to raise_error(Win32::TaskScheduler::Error)
      expect{ @task_scheduler.working_directory }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter raises an error if no active tasks' do
      expect{ @task_scheduler.working_directory }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter and setter are working successfully' do
      create_task
      expect(@task_scheduler.working_directory=(test_dir)).to eql(test_dir)
      expect(@task_scheduler.working_directory).to eql(test_dir)
    end
  end

  context '#priority' do
    let(:priority_val){ Win32::TaskScheduler::HIGH_PRIORITY_CLASS }
    it 'setter raises an error when an integer is not passed' do
      expect{ @task_scheduler.priority=('highest') }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      @task_scheduler.instance_variable_set(:@task, nil)
      expect{ @task_scheduler.priority=(priority_val) }.to raise_error(Win32::TaskScheduler::Error)
      expect{ @task_scheduler.priority }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter raises an error if no active tasks' do
      expect{ @task_scheduler.priority }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter and setter are working successfully' do
      create_task
      expect(@task_scheduler.priority=(priority_val)).to eql(priority_val)
      expect(@task_scheduler.priority).to eql('highest')
    end
  end

  context '#comment' do
    let(:comment_val){ 'Test Comment' }
    it 'setter raises an error when a String is not passed' do
      expect{ @task_scheduler.comment=(1) }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      @task_scheduler.instance_variable_set(:@task, nil)
      expect{ @task_scheduler.comment=(comment_val) }.to raise_error(Win32::TaskScheduler::Error)
      expect{ @task_scheduler.comment }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter raises an error if no active tasks' do
      expect{ @task_scheduler.comment }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter and setter are working successfully' do
      create_task
      expect(@task_scheduler.comment=(comment_val)).to eql(comment_val)
      expect(@task_scheduler.comment).to eql(comment_val)
    end

    it 'is an alias for description' do
      create_task
      expect(@task_scheduler.description=(comment_val)).to eql(comment_val)
      expect(@task_scheduler.description).to eql(comment_val)
    end
  end


  context '#max_run_time' do
    let(:max_run_time_val){ 1244145000000 } # Just a random time in miliseconds
    it 'setter raises an error when an integer is not passed' do
      expect{ @task_scheduler.max_run_time=('time') }.to raise_error(TypeError)
    end

    it 'raises an error when a task is not found' do
      @task_scheduler.instance_variable_set(:@task, nil)
      expect{ @task_scheduler.max_run_time=(max_run_time_val) }.to raise_error(Win32::TaskScheduler::Error)
      expect{ @task_scheduler.max_run_time }.to raise_error(Win32::TaskScheduler::Error)
    end

    it 'getter raises an error if no active tasks' do
      expect{ @task_scheduler.max_run_time }.to raise_error(Win32::TaskScheduler::Error)
    end

    # TODO: max_run_time setter expects time in miliseconds(may be an overhead),
    # and parser of getter is not working. It may be required to take a look again
    # Well, getter may implement `time_in_seconds` of TimeCalcHelper
    it 'getter and setter are working successfully' do
      skip 'skipping due to incorrect method logic'
      create_task
      expect(@task_scheduler.max_run_time=(max_run_time_val)).to eql(max_run_time_val)
      expect(@task_scheduler.max_run_time).to eql(max_run_time_val)
    end
  end

  def create_task
    trigger[:trigger_type] = Win32::TaskScheduler::ONCE
    @task_scheduler.new_work_item(@task, trigger)
    @task_scheduler.application_name = application
  end

  def delete_task
    @task_scheduler.delete(@task)
  end
end
