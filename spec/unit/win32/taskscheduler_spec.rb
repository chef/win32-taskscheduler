require "spec_helper"
require "win32ole"
require "win32/taskscheduler"

RSpec.describe Win32::TaskScheduler, :windows_only do
  before { create_test_folder }
  after { clear_them }
  before { load_task_variables }

  # Helper method to register task with specific user and password
  def register_task(user = 'SYSTEM', password = nil)
    # Make sure @test_folder and @task_definition are initialized in load_task_variables
    @current_task = @test_folder.RegisterTaskDefinition(
      @task,             # Task name
      @task_definition,  # Task definition
      6,                 # TASK_CREATE_OR_UPDATE flag
      user,              # User to run task as
      password,          # Password (nil if system user)
      3                  # Flags, as per your original code
    )
    @ts.instance_variable_set(:@task, @current_task) if @ts
  end

  describe "#Constructor" do
    let(:ts) { Win32::TaskScheduler }
    context "no of arguments" do
      it "zero" do
        expect(ts.new).to be_a(ts)
        expect(no_of_tasks).to eq(0)
      end

      it "one: task" do
        expect(ts.new(@task)).to be_a(ts)
        expect(no_of_tasks).to eq(0)
      end

      it "two: task, trigger; trigger_type is blank" do
        @trigger[:trigger_type] = nil
        expect { ts.new(@task, @trigger) }.to raise_error(ArgumentError)
        expect(no_of_tasks).to eq(0)
      end

      it "two: task, trigger; trigger_type is not blank; Creates a task" do
        expect(ts.new(@task, @trigger)).to be_a(ts)
        expect(no_of_tasks).to eq(1)
      end

      it "three: task, trigger, folder; Creates a task" do
        expect(ts.new(@task, @trigger, @folder)).to be_a(ts)
        expect(no_of_tasks).to eq(1)
      end

      it "four: task, trigger, folder, force; Creates a task" do
        expect(ts.new(@task, @trigger, @folder, @force)).to be_a(ts)
        expect(no_of_tasks).to eq(1)
      end

      it "raises an error for more than four arguments" do
        expect { ts.new(@task, @trigger, @folder, @force, 1) }.to raise_error(ArgumentError)
        expect { ts.new(@task, @trigger, @folder, @force, "abc") }.to raise_error(ArgumentError)
        expect(no_of_tasks).to eq(0)
      end
    end
  end

  describe "#Tasks" do
    before { create_task }
    it "Returns an Array" do
      expect(@ts.tasks).to be_a(Array)
    end

    it "Returns an Empty Array if no task is present" do
      delete_tasks_in(@test_folder)
      expect(@ts.tasks).to be_empty
    end
  end

  describe "#exists?" do
    let(:folder) { '\\Foo' }
    let(:force) { true }
    before { create_task }

    context "At Root" do
      it "Requires only task name" do
        expect(@ts.exists?(@task)).to be_truthy
      end

      it "Does not require root to be appended" do
        task = @test_path + @folder + '\\' + @task
        expect(@ts.exists?(task)).to be_falsey
      end
    end

    context "At Nested folder" do
      it "Returns false for non existing folder" do
        task = folder + '\\' + @task
        expect(@ts.exists?(task)).to be_falsey
      end

      it "Returns true for existing folder" do
        @ts = Win32::TaskScheduler.new(@task, @trigger, folder, force)
        task = @test_path + folder + '\\' + @task
        expect(@ts.exists?(task)).to be_truthy
      end
    end
  end

  describe "#get_task" do
    before { create_task }
    it "Requires a string: Task Name" do
      expect { @ts.get_task(0) }.to raise_error(TypeError)
    end
  end

  describe "#activate" do
    before { create_task }
    it "Requires a string: Task Name" do
      expect { @ts.activate(0) }.to raise_error(TypeError)
    end
  end

  describe "#delete" do
    before { create_task }
    it "Requires a string: Task Name" do
      expect { @ts.delete(0) }.to raise_error(TypeError)
    end
  end

  describe "#logon_type" do
    let(:user_id) { "User" }
    context "With Password" do
      let(:password) { "Password" }
      it "Returns PASSWORD flag for non-system users" do
        expect(@ts.send(:logon_type, user_id, password)).to eq(Win32::TaskScheduler::TASK_LOGON_PASSWORD)
      end

      it "Returns GROUP flag for group users" do
        user_id = "Guests"
        expect(@ts.send(:logon_type, user_id, password)).to eq(Win32::TaskScheduler::TASK_LOGON_GROUP)
      end

      it "Returns SERVICE_ACCOUNT flag for service-account users" do
        user_id = "SYSTEM"
        expect(@ts.send(:logon_type, user_id, password)).to eq(Win32::TaskScheduler::TASK_LOGON_SERVICE_ACCOUNT)
      end
    end

    context "Without Password" do
      let(:password) { nil }
      it "Returns INTERACTIVE_TOKEN flag for non-system users" do
        expect(@ts.send(:logon_type, user_id, password)).to eq(Win32::TaskScheduler::TASK_LOGON_INTERACTIVE_TOKEN)
      end

      it "Returns GROUP flag for group users" do
        user_id = "Guests"
        expect(@ts.send(:logon_type, user_id, password)).to eq(Win32::TaskScheduler::TASK_LOGON_GROUP)
      end

      it "Returns SERVICE_ACCOUNT flag for service-account users" do
        user_id = "SYSTEM"
        expect(@ts.send(:logon_type, user_id, password)).to eq(Win32::TaskScheduler::TASK_LOGON_SERVICE_ACCOUNT)
      end
    end
  end

  describe "#check_credential_requirements" do
    let(:password) { "ABCXYZ" }
    let(:nil_pass) { nil }

    context "System Users" do
      let(:user_id) { "SYSTEM" }
      it "does not require a password" do
        expect { @ts.send(:check_credential_requirements, user_id, nil_pass) }.not_to raise_error
      end
      it "raises error when password is sent" do
        expect { @ts.send(:check_credential_requirements, user_id, password) }.to raise_error(ArgumentError)
      end
    end

    context "Non-System Users" do
      let(:user_id) { "User" }
      context "For Interactive tasks" do
        before { @ts.instance_variable_set(:@interactive, true) }
        it "does not require a password" do
          expect { @ts.send(:check_credential_requirements, user_id, nil_pass) }.not_to raise_error
        end
        it "does not raises error when password is sent" do
          expect { @ts.send(:check_credential_requirements, user_id, password) }.not_to raise_error
        end
      end

      context "For Non-Interactive tasks" do
        before { @ts.instance_variable_set(:@interactive, false) }
        it "raises error when password is not sent" do
          expect { @ts.send(:check_credential_requirements, user_id, nil_pass) }.to raise_error(ArgumentError)
        end
        it "require a password" do
          expect { @ts.send(:check_credential_requirements, user_id, password) }.not_to raise_error
        end
      end
    end
  end

  # ** Added this block with the fix for your failing test **
  describe "#account_information" do
    context "Non-system users when task is Non-Interactive Require a password" do
      let(:user) { "ContainerAdministrator" }
      let(:password) { "nill" }
      before do
        register_task(user, password)
        @ts.instance_variable_set(:@interactive, false)
      end

      it "returns account information including user" do
        expect(@ts.account_information).to include(user)
      end
    end
  end

  private

  def load_task_variables
    time = Time.now
    allow_any_instance_of(Win32::TaskScheduler).to receive(:root_path).and_return(@test_path)
    @app = "notepad.exe"
    @task = "test_task"
    @folder = @test_path
    @force = false
    @trigger = { start_year: time.year, start_month: time.month,
                 start_day: time.day, start_hour: time.hour,
                 start_minute: time.min,
                 trigger_type: Win32::TaskScheduler::ONCE }
    @ts = Win32::TaskScheduler.new

    # Initialize these so register_task works correctly
    # Make sure the COM objects are available and these paths are correct for your environment
    scheduler = WIN32OLE.new('Schedule.Service')
    scheduler.Connect
    @test_folder ||= scheduler.GetFolder(@test_path)
    @task_definition ||= scheduler.NewTask(0)
  end
end
