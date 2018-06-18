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

  context '#account_information' do
    it 'returns nil when a task is not found' do
      expect(@task_scheduler.account_information).to be_nil
    end
    it 'returns User ID for the task' do
      create_task
      expect(@task_scheduler.account_information).to be_a(String)
    end
    it 'returns nil post task deletion' do
      skip 'instance variable @task needs to be set as nil in #delete_task in taskscheduler.rb'
      delete_task
      expect(@task_scheduler.account_information).to be_nil
    end
  end

  context '#set_account_information' do
    let(:numeric_pwd) {12345}
    let(:string_pwd) {"pwd@123"}
    it 'require user_id in String format' do
      expect{ @task_scheduler.set_account_information(numeric_pwd, string_pwd) }.to raise_error(TypeError)
    end
    context 'system user' do
      let(:user_id) {"SYSTEM"}
      it 'does not require password to be String' do
        expect{ @task_scheduler.set_account_information(user_id, numeric_pwd) }.to raise_error(Win32::TaskScheduler::Error)
        expect{ @task_scheduler.set_account_information(user_id, string_pwd) }.to raise_error(Win32::TaskScheduler::Error)
      end
      it 'raises an error when a task is not found' do
        expect{ @task_scheduler.set_account_information(user_id, string_pwd) }.to raise_error(Win32::TaskScheduler::Error)
      end
      it 'is able to set credentials for the task' do
        skip 'credentials required'
        create_task
        expect(@task_scheduler.set_account_information(user_id, string_pwd)).to eql(true)
        expect(@task_scheduler.account_information).to eql(user_id)
        delete_task
      end
    end
    context 'non-system user' do
      let(:user_id) {"Guest"}
      it 'require password to be String' do
        expect{ @task_scheduler.set_account_information(user_id, numeric_pwd) }.to raise_error(TypeError)
      end
      it 'raises an error when a task is not found' do
        expect{ @task_scheduler.set_account_information(user_id, string_pwd) }.to raise_error(Win32::TaskScheduler::Error)
      end
      it 'is able to set credentials for the task for string password' do
        skip 'credentials required'
        create_task
        expect(@task_scheduler.set_account_information(user_id, string_pwd)).to eql(true)
        expect(@task_scheduler.account_information).to eql(user_id)
        delete_task
      end
    end
  end

  context '#author' do
    let(:author_id){ "Author" }
    it 'setter requires String argument' do
      expect{ @task_scheduler.author=(123) }.to raise_error(TypeError)
    end
    it 'raises an error when a task is not found' do
      @task_scheduler.instance_variable_set(:@task, nil)
      expect{ @task_scheduler.author=(author_id) }.to raise_error(Win32::TaskScheduler::Error)
      expect{ @task_scheduler.author }.to raise_error(Win32::TaskScheduler::Error)
    end
    it 'getter raises an error if no active tasks' do
      @task_scheduler.instance_variable_set(:@task, nil)
      expect{ @task_scheduler.author }.to raise_error(Win32::TaskScheduler::Error)
    end
    it 'getter and setter are working successfully' do
      create_task
      expect(@task_scheduler.author=(author_id)).to eql(author_id)
      expect(@task_scheduler.author).to eql(author_id)
    end
    it 'is an alias with #creator' do
      create_task
      expect(@task_scheduler.creator=(author_id)).to eql(author_id)
      expect(@task_scheduler.creator).to eql(author_id)
    end
  end

  context '#trigger_count' do
    it 'raises an error when task is not found' do
      @task_scheduler.instance_variable_set(:@task, nil)
      expect{ @task_scheduler.trigger_count }.to raise_error(Win32::TaskScheduler::Error)
    end
    it 'returns no of trigger of a task' do
      create_task
      expect(@task_scheduler.trigger_count).to eql(1)
    end
  end

  context '#trigger_string' do
    it 'requires an integer argument' do
      expect{ @task_scheduler.trigger_string("First") }.to raise_error(TypeError)
    end
    it 'raises an error when task is not found' do
      @task_scheduler.instance_variable_set(:@task, nil)
      expect{ @task_scheduler.trigger_string(1) }.to raise_error(Win32::TaskScheduler::Error)
    end
    it 'returns trigger string' do
      create_task
      expect(@task_scheduler.trigger_string(0)).to be_a(String)
    end
    it 'returns error for invalid index' do
      create_task
      expect{ @task_scheduler.trigger_string(2) }.to raise_error(Win32::TaskScheduler::Error)
    end
  end

  context '#trigger' do
    let(:user_id) { (ENV['USERDOMAIN'] && ENV['USERNAME']) ? (ENV['USERDOMAIN'] + '\\' + ENV['USERNAME']) : "SYSTEM" }
    let(:index){ 0 }
    let(:time_limit) { 20 }
    let(:day) { Win32::TaskScheduler::FIRST }
    let(:sunday) { Win32::TaskScheduler::SUNDAY }
    let(:monday) { Win32::TaskScheduler::MONDAY }
    let(:january) { Win32::TaskScheduler::JANUARY }
    let(:february) { Win32::TaskScheduler::FEBRUARY }
    let(:first_week) { Win32::TaskScheduler::FIRST_WEEK }
    let(:second_week) { Win32::TaskScheduler::SECOND_WEEK }
    it 'getter requires Numeric argument' do
      expect{ @task_scheduler.trigger("First") }.to raise_error(TypeError)
    end
    it 'setter requires a Hash argument' do
      expect{ @task_scheduler.trigger=("First") }.to raise_error(TypeError)
    end
    it 'raises an error when a task is not found' do
      @task_scheduler.instance_variable_set(:@task, nil)
      expect{ @task_scheduler.trigger(index) }.to raise_error(Win32::TaskScheduler::Error)
    end
    it 'getter and setter are working successfully' do
      create_task
      expect(@task_scheduler.trigger=(trigger)).to eql(trigger)
      expect(@task_scheduler.trigger(index)).to be_a(Hash)
      expect(@task_scheduler.trigger(index)).not_to be_empty
    end
    context 'ONCE' do
      before do
        trigger[:trigger_type] = Win32::TaskScheduler::ONCE
        @task_scheduler.new_work_item(@task, trigger)
      end
      after { trigger[:trigger_type] = nil }
      it 'supports the param random_minutes_interval' do
        trigger[:random_minutes_interval] = time_limit
        expect(@task_scheduler.trigger=(trigger)).to eql(trigger)
        expect(@task_scheduler.trigger(index)).to include(random_minutes_interval: time_limit)
      end
    end
    context 'DAILY' do
      before do
        trigger[:trigger_type] = Win32::TaskScheduler::DAILY
        @task_scheduler.new_work_item(@task, trigger)
      end
      after { trigger[:trigger_type] = nil }
      it 'supports the param days_interval within type' do
        trigger[:type] = {days_interval: time_limit}
        expect(@task_scheduler.trigger=(trigger)).to eql(trigger)
        expect(@task_scheduler.trigger(index)).to include(type: {days_interval: time_limit})
      end
      it 'supports the param random_minutes_interval' do
        trigger[:random_minutes_interval] = time_limit
        expect(@task_scheduler.trigger=(trigger)).to eql(trigger)
        expect(@task_scheduler.trigger(index)).to include(random_minutes_interval: time_limit)
      end
    end
    context 'WEEKLY' do
      before do
        trigger[:trigger_type] = Win32::TaskScheduler::WEEKLY
        trigger[:type] = {days_of_week: sunday}
        @task_scheduler.new_work_item(@task, trigger)
      end
      after { trigger[:trigger_type] = nil, trigger[:type] = nil }
      it 'requires the param days_of_week within type' do
        trigger[:type] = {days_of_week: nil}
        expect{@task_scheduler.trigger=(trigger)}.to raise_error(Win32::TaskScheduler::Error)
      end
      it 'supports the param weeks_interval within type' do
        week_type = {weeks_interval: time_limit, days_of_week: sunday}
        trigger[:type] = week_type
        expect(@task_scheduler.trigger=(trigger)).to eql(trigger)
        expect(@task_scheduler.trigger(index)).to include(type: week_type)
      end
      it 'supports the param random_minutes_interval' do
        trigger[:random_minutes_interval] = time_limit
        expect(@task_scheduler.trigger=(trigger)).to eql(trigger)
        expect(@task_scheduler.trigger(index)).to include(random_minutes_interval: time_limit)
      end
    end
    context 'MONTHLYDATE' do
      before do
        trigger[:trigger_type] = Win32::TaskScheduler::MONTHLYDATE
        trigger[:type] = {months: january, days: day}
        @task_scheduler.new_work_item(@task, trigger)
      end
      after { trigger[:trigger_type] = nil, trigger[:type] = nil}
      it 'requires the param months and days within type' do
        trigger[:type] = {months: nil, days: nil}
        expect{@task_scheduler.trigger=(trigger)}.to raise_error(Win32::TaskScheduler::Error)
      end
      it 'supports the param random_minutes_interval' do
        trigger[:random_minutes_interval] = time_limit
        expect(@task_scheduler.trigger=(trigger)).to eql(trigger)
        expect(@task_scheduler.trigger(index)).to include(random_minutes_interval: time_limit)
      end
    end
    context 'MONTHLYDOW' do
      before do
        trigger[:trigger_type] = Win32::TaskScheduler::MONTHLYDOW
        @task_scheduler.new_work_item(@task, trigger)
      end
      after { trigger[:trigger_type] = nil, trigger[:type] = nil}
      it 'defaults to Sunday in First Week of January' do
        month_type = {months: january, days_of_week: sunday, weeks_of_month: first_week}
        expect(@task_scheduler.trigger(index)).to include(type: month_type)
      end
      it 'can specify months, weekday, and week of the month within type' do
        month_type = {months: february, days_of_week: monday, weeks_of_month: second_week}
        trigger[:type] = month_type
        expect(@task_scheduler.trigger=(trigger)).to eql(trigger)
        expect(@task_scheduler.trigger(index)).to include(type: month_type)
      end
      it 'supports the param random_minutes_interval' do
        trigger[:random_minutes_interval] = time_limit
        expect(@task_scheduler.trigger=(trigger)).to eql(trigger)
        expect(@task_scheduler.trigger(index)).to include(random_minutes_interval: time_limit)
      end
    end
    context 'AT_LOGON' do
      before do
        trigger[:trigger_type] = Win32::TaskScheduler::AT_LOGON
        @task_scheduler.new_work_item(@task, trigger)
      end
      after { trigger[:trigger_type] = nil }
      it 'returns user_id only if it is passed along with trigger' do
        trigger[:user_id] = nil
        expect(@task_scheduler.trigger=(trigger)).to eql(trigger)
        expect(@task_scheduler.trigger(index)).not_to include(:user_id)
        trigger[:user_id] = user_id
        expect(@task_scheduler.trigger=(trigger)).to eql(trigger)
        expect(@task_scheduler.trigger(index)).to include(user_id: user_id)
      end
      it 'supports the param delay_duration' do
        trigger[:delay_duration] = time_limit
        expect(@task_scheduler.trigger=(trigger)).to eql(trigger)
        expect(@task_scheduler.trigger(index)).to include(delay_duration: time_limit)
      end
    end
    context 'AT_SYSTEMSTART' do
      before do
        trigger[:trigger_type] = Win32::TaskScheduler::AT_SYSTEMSTART
        @task_scheduler.new_work_item(@task, trigger)
      end
      after { trigger[:trigger_type] = nil }
      it 'supports the param delay_duration' do
        trigger[:delay_duration] = time_limit
        expect(@task_scheduler.trigger=(trigger)).to eql(trigger)
        expect(@task_scheduler.trigger(index)).to include(delay_duration: time_limit)
      end
    end
    context 'ON_IDLE' do
      before do
        trigger[:trigger_type] = Win32::TaskScheduler::ON_IDLE
        @task_scheduler.new_work_item(@task, trigger)
      end
      after { trigger[:trigger_type] = nil }
      it 'supports the param execution_time_limit' do
        trigger[:execution_time_limit] = time_limit
        expect(@task_scheduler.trigger=(trigger)).to eql(trigger)
        expect(@task_scheduler.trigger(index)).to include(execution_time_limit: time_limit)
      end
    end
  end

  context '#principals' do
    let(:principals){ { id: "Author", display_name: "DispName", logon_type: Win32::TaskScheduler::TASK_LOGON_SERVICE_ACCOUNT,
                        run_level: Win32::TaskScheduler::TASK_RUNLEVEL_HIGHEST } }
    it 'raises an error when a task is not found' do
      @task_scheduler.instance_variable_set(:@task, nil)
      expect{ @task_scheduler.configure_principals(principals) }.to raise_error(Win32::TaskScheduler::Error)
      expect{ @task_scheduler.principals }.to raise_error(Win32::TaskScheduler::Error)
    end
    it 'getter and setter are working successfully' do
      create_task
      expect(@task_scheduler.configure_principals(principals)).to eql(principals)
      expect(@task_scheduler.principals).to include(principals)
    end
  end

  context '#settings' do
    let(:settings){ { allow_demand_start: true, restart_interval: "",
                      restart_count: 0, multiple_instances: 2, stop_if_going_on_batteries: true,
                      disallow_start_if_on_batteries: true, allow_hard_terminate: true,
                      start_when_available: false, run_only_if_network_available: false,
                      enabled: true, delete_expired_task_after: "", 
                      priority: 7, compatibility: 2, hidden: false, 
                      run_only_if_idle: false, wake_to_run: false,
                      disallow_start_on_remote_app_session: false, use_unified_scheduling_engine: false,
                      maintenance_settings: nil, volatile: false, 
                      # SKIP: Comenting below options since they are not working properly
                      # execution_time_limit: 20,
                      # network_settings: { name: "", id: "" },
                      # idle_settings: { idle_duration: 10, wait_timeout: 1, 
                      #                  stop_on_idle_end: true, restart_on_idle: false }
                    }
                  }
    it 'raises an error when a task is not found' do
      @task_scheduler.instance_variable_set(:@task, nil)
      expect{ @task_scheduler.configure_settings(settings) }.to raise_error(Win32::TaskScheduler::Error)
      expect{ @task_scheduler.settings }.to raise_error(Win32::TaskScheduler::Error)
    end
    it 'getter and setter are working successfully' do
      create_task
      expect(@task_scheduler.configure_settings(settings)).to eql(settings)
      expect(@task_scheduler.settings).to include(settings)
    end
  end

  private
  def create_task
    trigger[:trigger_type] = Win32::TaskScheduler::ONCE
    @task_scheduler.new_work_item(@task, trigger)
    @task_scheduler.application_name = application
  end

  def delete_task
    @task_scheduler.delete(@task)
  end
end
