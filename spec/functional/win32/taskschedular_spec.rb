require 'spec_helper'
require 'win32ole'
require 'win32/taskscheduler'

RSpec.describe Win32::TaskScheduler, :windows_only do
  before { create_test_folder }
  after { clear_them }
  before { load_task_variables }

  describe '#constructor' do
    let(:ts) { Win32::TaskScheduler }
    context 'Task' do
      it 'Does not creates task when default(nil)' do
        task = nil
        expect(ts.new(task)).to be_a(ts)
        expect(no_of_tasks).to eq(0)
      end

      it 'Does not creates task without trigger' do
        expect(ts.new(@task)).to be_a(ts)
        expect(no_of_tasks).to eq(0)
      end

      it 'Creates a task with trigger' do
        expect(ts.new(@task, @trigger)).to be_a(ts)
        expect(no_of_tasks).to eq(1)
      end
    end

    context 'Trigger' do
      it 'Does not creates task when default(nil)' do
        trigger = nil
        expect(ts.new(@task, trigger)).to be_a(ts)
        expect(no_of_tasks).to eq(0)
      end

      it 'Raises error without trigger type' do
        @trigger[:trigger_type] = nil
        expect { ts.new(@task, @trigger) }.to raise_error(ArgumentError)
        expect(no_of_tasks).to eq(0)
      end

      it 'Creates a task with trigger type' do
        expect(ts.new(@task, @trigger)).to be_a(ts)
        expect(no_of_tasks).to eq(1)
      end
    end

    context 'Folder' do
      let(:folder) { '\\Foo' }
      it 'Raises error when nil' do
        expect { ts.new(@task, @trigger, nil) }.to raise_error(NoMethodError)
        expect(no_of_tasks).to eq(0)
      end

      it "Raises error when path separators(\\\) are absent" do
        invalid_path = 'Foo'
        expect { ts.new(@task, @trigger, invalid_path) }.to raise_error(ArgumentError)
        expect(no_of_tasks).to eq(0)
      end

      it 'Creates a task when default(root)' do
        expect(ts.new(@task, @trigger, @folder)).to be_a(ts)
        expect(no_of_tasks).to eq(1)
      end

      context 'When force is true' do
        let(:force) { true }
        it 'Creates a task at specified folder' do
          expect(ts.new(@task, @trigger, folder, force)).to be_a(ts)
          expect(no_of_tasks(folder)).to eq(1)
        end
      end

      context 'When force is false' do
        let(:force) { false }
        it 'Raises an error when folder does not exists' do
          expect { ts.new(@task, @trigger, folder, force) }.to raise_error(ArgumentError)
        end

        it 'Creates a task at specified folder if it exists' do
          new_folder = @test_folder.CreateFolder(folder)
          expect(ts.new(@task, @trigger, new_folder.Path, force)).to be_a(ts)
          expect(no_of_tasks(folder)).to eq(1)
        end
      end
    end
  end

  describe '#tasks' do
    before { create_task }
    it 'Returns Task Names' do
      expect(@ts.tasks).to include(@task)
    end

    it 'is an alias with enum' do
      expect(@ts.enum).to include(@task)
    end
  end

  describe '#exists?' do
    let(:folder) { '\\Foo' }
    let(:force) { true }
    before { create_task }

    context 'valid task path returns true' do
      it 'at root' do
        expect(@ts.exists?(@task)).to be_truthy
      end

      it 'at nested folder' do
        @ts = Win32::TaskScheduler.new(@task, @trigger, folder, force)
        task = @test_path + folder + '\\' + @task
        expect(@ts.exists?(task)).to be_truthy
      end
    end

    context 'invalid task path returns false' do
      it 'at root' do
        expect(@ts.exists?('invalid')).to be_falsy
      end

      it 'at nested folder' do
        @ts = Win32::TaskScheduler.new(@task, @trigger, folder, force)
        task = folder + '\\' + 'invalid'
        expect(@ts.exists?(task)).to be_falsy
      end
    end
  end

  describe '#get_task' do
    let(:folder) { '\\Foo' }
    let(:force) { true }
    before { create_task }

    context 'from root' do
      it 'Returns the Task if exists' do
        expect(@ts.get_task(@task)).to be_a(@current_task.class)
      end

      it 'Raises an error if task does not exists' do
        expect { @ts.get_task('invalid') }.to raise_error(tasksch_err)
      end
    end

    context 'from nested folder' do
      it 'Returns the Task if exists' do
        skip 'Implementation Pending'
      end

      it 'Raises an error if task does not exists' do
        skip 'Implementation Pending'
      end
    end
  end

  describe '#activate' do
    let(:folder) { '\\Foo' }
    let(:force) { true }
    before { create_task }

    context 'from root' do
      it 'Activates the Task if exists' do
        expect(@ts.activate(@task)).to be_a(@current_task.class)
        expect(@current_task.Enabled).to be_truthy
      end

      it 'Raises an error if task does not exists' do
        expect { @ts.activate('invalid') }.to raise_error(tasksch_err)
      end
    end

    context 'from nested folder' do
      it 'Activates the Task if exists' do
        skip 'Implementation Pending'
      end

      it 'Raises an error if task does not exists' do
        skip 'Implementation Pending'
      end
    end
  end

  describe '#delete' do
    let(:folder) { '\\Foo' }
    let(:force) { true }
    before { create_task }

    context 'from root' do
      it 'Deletes the Task if exists' do
        expect(no_of_tasks).to eq(1)
        @ts.delete(@task)
        expect(no_of_tasks).to eq(0)
      end

      it 'Raises an error if task does not exists' do
        expect { @ts.delete('invalid') }.to raise_error(tasksch_err)
      end
    end

    context 'from nested folder' do
      it 'Activates the Task if exists' do
        skip 'Implementation Pending'
      end

      it 'Raises an error if task does not exists' do
        skip 'Implementation Pending'
      end
    end
  end

  describe '#run' do
    before { create_task }
    after { stop_the_app }

    it 'Execute(Start) the Task' do
      @ts.run
      expect(@ts.status).to eq('running')
        .or eq('queued')
        .or eq('ready') # It takes time to load sometimes
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.run }.to raise_error(tasksch_err)
    end
  end

  describe '#terminate' do
    before { create_task }
    after { stop_the_app }

    it 'terminates the Task if exists' do
      @ts.terminate
      expect(@ts.status).to eq('ready')
    end

    it 'has an alias stop' do
      @ts.stop
      expect(@ts.status).to eq('ready')
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.terminate }.to raise_error(tasksch_err)
    end
  end

  describe '#application_name' do
    before { create_task }
    it 'Returns the application name of task' do
      check = @app
      expect(@ts.application_name).to eq(check)
    end

    it 'Sets the application name of task' do
      stub_user
      check = 'cmd.exe'
      expect(@ts.application_name = check).to eq(check)
      expect(@ts.application_name).to eq(check)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.application_name }.to raise_error(tasksch_err)
    end
  end

  describe '#parameters' do
    before { create_task }
    it 'Returns the parameters of task' do
      check = ''
      expect(@ts.parameters).to eq(check)
    end

    it 'Sets the parameters of task' do
      stub_user
      check = 'cmd.exe'
      expect(@ts.parameters = check).to eq(check)
      expect(@ts.parameters).to eq(check)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.parameters }.to raise_error(tasksch_err)
    end
  end

  describe '#working_directory' do
    before { create_task }
    it 'Returns the working directory of task' do
      check = ''
      expect(@ts.working_directory).to eq(check)
    end

    it 'Sets the working directory of task' do
      stub_user
      check = Dir.pwd
      expect(@ts.working_directory = check).to eq(check)
      expect(@ts.working_directory).to eq(check)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.working_directory }.to raise_error(tasksch_err)
    end
  end

  describe '#priority' do
    before { create_task }
    it 'Returns default priority of task' do
      check = 'below_normal_7'
      expect(@ts.priority).to eq(check)
    end

    it '0 Sets the priority of task to critical' do
      stub_user
      check = 0
      expect(@ts.priority = check).to eq(check)
      expect(@ts.priority).to eq('critical')
    end

    it '1 Sets the priority of task to highest' do
      stub_user
      check = 1
      expect(@ts.priority = check).to eq(check)
      expect(@ts.priority).to eq('highest')
    end

    it '2 Sets the priority of task to above_normal_2' do
      stub_user
      check = 2
      expect(@ts.priority = check).to eq(check)
      expect(@ts.priority).to eq('above_normal_2')
    end

    it '3 Sets the priority of task to above_normal_3' do
      stub_user
      check = 3
      expect(@ts.priority = check).to eq(check)
      expect(@ts.priority).to eq('above_normal_3')
    end

    it '4 Sets the priority of task to normal_4' do
      stub_user
      check = 4
      expect(@ts.priority = check).to eq(check)
      expect(@ts.priority).to eq('normal_4')
    end

    it '5 Sets the priority of task to normal_5' do
      stub_user
      check = 5
      expect(@ts.priority = check).to eq(check)
      expect(@ts.priority).to eq('normal_5')
    end

    it '6 Sets the priority of task to normal_6' do
      stub_user
      check = 6
      expect(@ts.priority = check).to eq(check)
      expect(@ts.priority).to eq('normal_6')
    end

    it '7 Sets the priority of task to below_normal_7' do
      stub_user
      check = 7
      expect(@ts.priority = check).to eq(check)
      expect(@ts.priority).to eq('below_normal_7')
    end

    it '8 Sets the priority of task to below_normal_8' do
      stub_user
      check = 8
      expect(@ts.priority = check).to eq(check)
      expect(@ts.priority).to eq('below_normal_8')
    end

    it '9 Sets the priority of task to lowest' do
      stub_user
      check = 9
      expect(@ts.priority = check).to eq(check)
      expect(@ts.priority).to eq('lowest')
    end

    it '10 Sets the priority of task to idle' do
      stub_user
      check = 10
      expect(@ts.priority = check).to eq(check)
      expect(@ts.priority).to eq('idle')
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.priority }.to raise_error(tasksch_err)
    end
  end

  describe '#comment' do
    before { create_task }
    it 'Returns the comment of task' do
      check = 'Sample task for testing purpose'
      expect(@ts.comment).to eq(check)
    end

    it 'Sets the Comment(Description) of task' do
      stub_user
      check = 'Description To Test'
      expect(@ts.comment = check).to eq(check)
      expect(@ts.comment).to eq(check)
    end

    it 'alias with Description' do
      stub_user
      check = 'Description To Test'
      expect(@ts.description = check).to eq(check)
      expect(@ts.comment).to eq(check)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.comment }.to raise_error(tasksch_err)
    end
  end

  describe '#author' do
    before { create_task }
    it 'Returns the author of task' do
      check = 'Rspec'
      expect(@ts.author).to eq(check)
    end

    it 'Sets the Author of task' do
      stub_user
      check = 'Author'
      expect(@ts.author = check).to eq(check)
      expect(@ts.author).to eq(check)
    end

    it 'alias with Creator' do
      stub_user
      check = 'Description To Test'
      expect(@ts.creator = check).to eq(check)
      expect(@ts.author).to eq(check)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.author }.to raise_error(tasksch_err)
    end
  end

  describe '#max_run_time' do
    before { create_task }
    it 'Returns the Execution Time Limit of task' do
      skip 'Due to logical error in implementation'
      check = (72 * 60 * 60) # Task: PT72H
      expect(@ts.max_run_time).to eq(check)
    end

    it 'Sets the max_run_time of task' do
      skip 'Due to logical error in implementation'
      check = 1_244_145_000_000
      expect(@ts.max_run_time = check).to eq(check)
      expect(@ts.max_run_time).to eq(check)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.max_run_time }.to raise_error(tasksch_err)
    end
  end

  describe '#enabled?' do
    before { create_task }
    it 'Returns true when Task is enabled' do
      expect(@ts.enabled?).to be_truthy
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.enabled? }.to raise_error(tasksch_err)
    end
  end

  describe '#next_run_time' do
    before { create_task }
    it 'Returns a Time object that indicates the next time the task will run' do
      expect(@ts.next_run_time).to be_a(Time)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.next_run_time }.to raise_error(tasksch_err)
    end
  end

  describe '#most_recent_run_time' do
    before { create_task }

    it 'Returns nil if the task has never run' do
      skip 'Error in implementation; Time.parse requires String'
      expect(@ts.most_recent_run_time).to be_nil
    end

    it 'Returns Time object indicating the most recent time the task ran' do
      skip 'Error in implementation; Time.parse requires String'
      @ts.run
      expect(@ts.most_recent_run_time).to be_a(Time)
      @ts.stop
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.most_recent_run_time }.to raise_error(tasksch_err)
    end
  end

  describe '#account_information' do
    before { create_task }
    context 'Service Account Users' do
      let(:service_user) { 'LOCAL SERVICE' }
      let(:password) { nil }
      it 'Does not require any password' do
        expect(@ts.set_account_information(service_user, password)). to be_truthy
        expect(@ts.account_information.upcase).to eq(service_user)
      end
      it 'passing account_name will display account_simple_name' do
        user = 'NT AUTHORITY\LOCAL SERVICE'
        expect(@ts.set_account_information(user, password)). to be_truthy
        expect(@ts.account_information.upcase).to eq(service_user)
      end
      it 'Raises an error if password is given' do
        password = 'XYZ'
        expect { @ts.set_account_information(service_user, password) }. to raise_error(tasksch_err)
      end
    end
    context 'Built in Groups' do
      let(:group_user) { 'USERS' }
      let(:password) { nil }
      it 'Does not require any password' do
        expect(@ts.set_account_information(group_user, password)). to be_truthy
        expect(@ts.account_information.upcase).to eq(group_user)
      end
      it 'passing account_name will display account_simple_name' do
        user = 'BUILTIN\USERS'
        expect(@ts.set_account_information(user, password)). to be_truthy
        expect(@ts.account_information.upcase).to eq(group_user)
      end
      it 'Raises an error if password is given' do
        password = 'XYZ'
        expect { @ts.set_account_information(group_user, password) }. to raise_error(tasksch_err)
      end
    end
    context 'Non-system users' do
      it 'Require a password' do
        user = ENV['user']
        password = nil
        expect { @ts.set_account_information(user, password) }. to raise_error(tasksch_err)
        expect(@ts.account_information).to include(user)
      end
    end
    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      user = 'User'
      password = 'XXXX'
      expect { @ts.set_account_information(user, password) }. to raise_error(tasksch_err)
      expect(@ts.account_information).to be_nil
    end
  end

  describe '#status' do
    before { create_task }
    it 'Returns tasks status' do
      expect(@ts.status).to be_a(String)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.status }.to raise_error(tasksch_err)
    end
  end

  describe '#exit_code' do
    before { create_task }
    it 'Returns the exit code from the last scheduled run' do
      expect(@ts.exit_code).to be_a(Integer)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.exit_code }.to raise_error(tasksch_err)
    end
  end

  describe '#trigger_count' do
    before { create_task }
    it 'Returns No of triggers associated with the task' do
      expect(@ts.trigger_count).to eq(1)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.trigger_count }.to raise_error(tasksch_err)
    end
  end

  describe '#trigger_string' do
    before { create_task }
    it 'Returns a string that describes the current trigger at '\
       'the specified index for the active task' do
      expect(@ts.trigger_string(0)).to be_a(String)
    end

    it 'Raises an error if trigger is not found at the given index' do
      expect { @ts.trigger_string(1) }.to raise_error(tasksch_err)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.trigger_string(0) }.to raise_error(tasksch_err)
    end
  end

  describe '#idle_settings' do
    before { create_task }
    it 'Returns a hash containing idle settings of the current task' do
      settings = @ts.idle_settings
      expect(settings).to be_a(Hash)
      expect(settings).to include(stop_on_idle_end: true)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.idle_settings }. to raise_error(tasksch_err)
    end
  end

  describe '#network_settings' do
    before { create_task }
    it 'Returns a hash containing network settings of the current task' do
      settings = @ts.network_settings
      expect(settings).to be_a(Hash)
      expect(settings).to include(name: '')
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.network_settings }. to raise_error(tasksch_err)
    end
  end

  describe '#settings' do
    before { create_task }
    it 'Returns a hash containing all the settings of the current task' do
      settings = @ts.settings
      expect(settings).to be_a(Hash)
      expect(settings).to include(enabled: true)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.settings }. to raise_error(tasksch_err)
    end
  end

  describe '#configure_settings' do
    before { create_task }
    before { stub_user }

    it 'Require a hash' do
      expect { @ts.configure_settings('XYZ') }. to raise_error(TypeError)
      expect(@ts.configure_settings({})).to eq({})
    end

    it 'Raises an error if invalid setting is passed' do
      expect { @ts.configure_settings(invalid: true) }. to raise_error(TypeError)
    end

    it 'Sets settings of the task and returns a hash' do
      settings = { allow_demand_start: false,
                   disallow_start_if_on_batteries: false,
                   enabled: false,
                   stop_if_going_on_batteries: false }

      expect(@ts.configure_settings(settings)).to eq(settings)
      expect(@ts.settings).to include(settings)
    end

    it 'Sets idle settings of the task' do
      # It accepts time in minutes
      # and returns the result in ISO8601 format
      settings = { idle_duration: 11,
                   wait_timeout: 10,
                   stop_on_idle_end: true,
                   restart_on_idle: false }

      result_settings = { idle_duration: 'PT11M',
                          wait_timeout: 'PT10M',
                          stop_on_idle_end: true,
                          restart_on_idle: false }

      @ts.configure_settings(settings)
      expect(@ts.idle_settings).to eq(result_settings)
    end

    it 'Set settings and idle settings of the task' do
      settings = {
        allow_demand_start: false,
        disallow_start_if_on_batteries: false,
        stop_if_going_on_batteries: false
      }

      idle_settings = {
        stop_on_idle_end: true,
        restart_on_idle: false
      }

      @ts.configure_settings(settings.merge(idle_settings))
      expect(@ts.settings).to include(settings)
      expect(@ts.idle_settings).to include(idle_settings)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.settings }. to raise_error(tasksch_err)
    end
  end

  describe '#principals' do
    before { create_task }
    it 'Returns a hash containing all the principal information of the current task' do
      principals = @ts.principals
      expect(principals).to be_a(Hash)
      expect(principals).to include(id: 'Author')
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.principals }. to raise_error(tasksch_err)
    end
  end

  describe '#principals' do
    before { create_task }
    it 'Returns a hash containing all the principal information of the current task' do
      principals = @ts.principals
      expect(principals).to be_a(Hash)
      expect(principals).to include(id: 'Author')
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.principals }. to raise_error(tasksch_err)
    end
  end

  describe '#configure_principals' do
    before { create_task }
    before { stub_user }

    it 'Sets the principals for current active task' do
      principals = @ts.principals
      principals[:user_id] = 'SYSTEM'
      principals[:display_name] = 'Rspec'
      principals[:run_level] = Win32::TaskScheduler::TASK_RUNLEVEL_HIGHEST
      principals[:logon_type] = Win32::TaskScheduler::TASK_LOGON_SERVICE_ACCOUNT

      @ts.configure_principals(principals)
      expect(@ts.principals).to eq(principals)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.principals }. to raise_error(tasksch_err)
    end
  end

  describe '#trigger' do
    before { create_task }

    it 'Returns a hash that describes the trigger '\
       'at the given index for the current task' do
      trigger = @ts.trigger(0)
      expect(trigger).to be_a(Hash)
      expect(trigger).not_to be_empty
    end

    it 'Raises an error if trigger is not found at the given index' do
      expect { @ts.trigger(1) }.to raise_error(tasksch_err)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.trigger(0) }.to raise_error(tasksch_err)
    end
  end

  describe '#trigger=' do
    before { create_task }
    before { stub_user }
    all_triggers.each do |type, trigger_hash|
      it "Updates the first trigger for type: #{type}" do
        @ts.trigger = trigger_hash
        returned_trigger = @ts.trigger(0)
        expect(returned_trigger).to eq(trigger_hash)
        expect(@ts.trigger_count).to eq(1)
      end

      it "Updates the trigger at given index for type: #{type}" do
        skip 'Implementation Pending'
        @ts.trigger = trigger_hash
        returned_trigger = @ts.trigger(0)
        expect(returned_trigger).to eq(trigger_hash)
        expect(@ts.trigger_count).to eq(1)
      end

      it 'Raises an error if task does not exists' do
        @ts.instance_variable_set(:@task, nil)
        expect { @ts.trigger = trigger_hash }.to raise_error(tasksch_err)
      end
    end
  end

  describe '#new_work_item' do
    before { create_task }
    context 'Creates a new task' do
      all_triggers.each do |type, trigger_hash|
        it "for the triger: #{type}" do
          task_count = 1 # One task already exists
          expect(@ts.new_work_item("Task_#{type}", trigger_hash)).to be_a(@current_task.class)
          task_count += 1
          expect(no_of_tasks).to eq(task_count)
        end
      end
    end

    context 'Updates the existing task' do
      all_triggers.each do |type, trigger_hash|
        it "for the triger: #{type}" do
          expect(@ts.new_work_item(@task, trigger_hash)).to be_a(@current_task.class)
          expect(no_of_tasks).to eq(1)
        end

        it 'Raises an error if task does not exists' do
          @ts.instance_variable_set(:@task, nil)
          expect { @ts.trigger = trigger_hash }.to raise_error(tasksch_err)
        end
      end
    end
  end

  private

  def load_task_variables
    time = Time.now
    # Ensuring root path will be test path
    allow_any_instance_of(Win32::TaskScheduler).to receive(:root_path).and_return(@test_path)
    @app = 'iexplore.exe'
    @task = 'test_task'
    @folder = @test_path
    @force = false
    @trigger = { start_year: time.year, start_month: time.month,
                 start_day: time.day, start_hour: time.hour,
                 start_minute: time.min,
                 # Will update this in test cases when required
                 trigger_type: Win32::TaskScheduler::ONCE }
    @ts = Win32::TaskScheduler.new
  end

  # Sets the user Id as nil, hence SYSTEM will be considered
  # Alternatively, we may to provide the login users password
  def stub_user
    # @ts.instance_variable_set(:@password, 'user_login_password')
    allow(@ts).to receive(:task_user_id).and_return(nil)
  end

  # Will stop the appliaction only if it is running
  def stop_the_app
    `tasklist /FI "IMAGENAME eq #{@app}" 2>NUL | find /I /N "#{@app}">NUL
    if "%ERRORLEVEL%"=="0" taskkill /f /im #{@app}`
  end
end
