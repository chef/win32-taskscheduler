require 'spec_helper'
require 'win32ole'
require 'win32/taskscheduler'
require 'byebug'

RSpec.describe Win32::TaskScheduler, :windows_only do
  before { create_test_folder }
  after { clear_them }
  before { load_task_variables }

  describe '#Constructor' do
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

      it 'Raises error when path separators(\\\) are absent' do
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

  describe '#Tasks' do
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
        task = folder + '\\' + @task
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
        skip 'Code missing'
      end

      it 'Raises an error if task does not exists' do
        skip 'Code missing'
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
        skip 'Code missing'
      end

      it 'Raises an error if task does not exists' do
        skip 'Code missing'
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
        skip 'Code missing'
      end

      it 'Raises an error if task does not exists' do
        skip 'Code missing'
      end
    end
  end

  describe '#run' do
    before { create_task }
    it 'Execute(Start) the Task' do
      @ts.run
      expect(app_running?).to be_truthy
      stop_the_app
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.run }.to raise_error(tasksch_err)
    end
  end

  describe '#terminate' do
    before { create_task }
    it 'terminates the Task if exists' do
      @ts.run
      @ts.terminate
      expect(app_running?).to be_falsy
    end

    it 'has an alias stop' do
      @ts.run
      @ts.terminate
      expect(app_running?).to be_falsy
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
      check = 'cmd.exe'
      expect(@ts.application_name = check).to eq(check)
      expect(@ts.application_name).to eq(check)
    end

    it 'Raises an error if task does not exists' do
      @ts.instance_variable_set(:@task, nil)
      expect { @ts.application_name }.to raise_error(tasksch_err)
    end
  end

  private

  def load_task_variables
    time = Time.now
    # Ensuring root path will be test path
    allow_any_instance_of(Win32::TaskScheduler).to receive(:root_path).and_return(@test_path)
    @app = 'notepad.exe'
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

  # Determines if notepad is running
  def app_running?
    status = `tasklist | find "notepad.exe"`
    !status.empty?
  end

  def stop_the_app
    `taskkill /IM notepad.exe`
  end
end
