require 'spec_helper'
require 'win32ole'
require 'win32/taskscheduler'

RSpec.describe Win32::TaskScheduler, :windows_only do
  before { create_test_folder }
  after { clear_them }
  before { load_task_variables }

  describe '#Constructor' do
    let(:ts) { Win32::TaskScheduler }
    context 'no of arguments' do
      it 'zero' do
        expect(ts.new).to be_a(ts)
        expect(no_of_tasks).to eq(0)
      end

      it 'one: task' do
        expect(ts.new(@task)).to be_a(ts)
        expect(no_of_tasks).to eq(0)
      end

      it 'two: task, trigger; trigger_type is blank' do
        @trigger[:trigger_type] = nil
        expect { ts.new(@task, @trigger) }.to raise_error(ArgumentError)
        expect(no_of_tasks).to eq(0)
      end

      it 'two: task, trigger; trigger_type is not blank; Creates a task' do
        expect(ts.new(@task, @trigger)).to be_a(ts)
        expect(no_of_tasks).to eq(1)
      end

      it 'three: task, trigger, folder; Creates a task' do
        expect(ts.new(@task, @trigger, @folder)).to be_a(ts)
        expect(no_of_tasks).to eq(1)
      end

      it 'four: task, trigger, folder, force; Creates a task' do
        expect(ts.new(@task, @trigger, @folder, @force)).to be_a(ts)
        expect(no_of_tasks).to eq(1)
      end

      it 'raises an error for more than four arguments' do
        expect { ts.new(@task, @trigger, @folder, @force, 1) }.to raise_error(ArgumentError)
        expect { ts.new(@task, @trigger, @folder, @force, 'abc') }.to raise_error(ArgumentError)
        expect(no_of_tasks).to eq(0)
      end
    end
  end

  describe '#Tasks' do
    before { create_task }
    it 'Returns an Array' do
      expect(@ts.tasks).to be_a(Array)
    end

    it 'Returns an Empty Array if no task is present' do
      delete_tasks_in(@test_folder)
      expect(@ts.tasks).to be_empty
    end
  end

  describe '#exists?' do
    let(:folder) { '\\Foo' }
    let(:force) { true }
    before { create_task }

    context 'At Root' do
      it 'Requires only task name' do
        expect(@ts.exists?(@task)).to be_truthy
      end

      it 'Does not require root to be appended' do
        task = @test_path + @folder + '\\' + @task
        expect(@ts.exists?(task)).to be_falsey
      end
    end

    context 'At Nested folder' do
      it 'Returns false for non existing folder' do
        task = folder + '\\' + @task
        expect(@ts.exists?(task)).to be_falsey
      end

      it 'Returns true for existing folder' do
        @ts = Win32::TaskScheduler.new(@task, @trigger, folder, force)
        task = @test_path + folder + '\\' + @task
        expect(@ts.exists?(task)).to be_truthy
      end
    end
  end

  describe '#get_task' do
    before { create_task }
    it 'Requires a string: Task Name' do
      expect { @ts.get_task(0) }.to raise_error(TypeError)
    end
  end

  describe '#activate' do
    before { create_task }
    it 'Requires a string: Task Name' do
      expect { @ts.activate(0) }.to raise_error(TypeError)
    end
  end

  describe '#delete' do
    before { create_task }
    it 'Requires a string: Task Name' do
      expect { @ts.delete(0) }.to raise_error(TypeError)
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
end
