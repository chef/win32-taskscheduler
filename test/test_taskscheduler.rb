##########################################################################
# test_taskscheduler.rb
#
# Test suite for the win32-taskscheduler package. You should run this
# via the 'rake test' task.
##########################################################################
require 'win32/taskscheduler'
require 'socket'
require 'test-unit'
require 'etc'
include Win32

class TC_TaskScheduler < Test::Unit::TestCase
  def self.startup
    @@host = Socket.gethostname
    @@user = Etc.getlogin
  end

  def setup
    @task     = 'foo'
    @job_file = "C:\\WINDOWS\\Tasks\\test.job"

    @trigger = {
      'start_year'   => 2009,
      'start_month'  => 4,
      'start_day'    => 11,
      'start_hour'   => 7,
      'start_minute' => 14,
      'trigger_type' => TaskScheduler::DAILY,
      'type'         => { 'days_interval' => 1 }
    }

    @ts = TaskScheduler.new
  end

  test "version constant is set to expected value" do
    assert_equal('0.3.0', TaskScheduler::VERSION)
  end

  test "account_information method basic functionality" do
    assert_respond_to(@ts, :account_information)
  end

  test "account_information returns the task owner" do
    @ts.activate(@task)
    assert_equal("#{@@host}\\#{@@user}", @ts.account_information)
  end

  test "account_information returns nil if no task has been activated" do
    assert_nil(@ts.account_information)
  end

  test "account_information does not accept any arguments" do
    assert_raise(ArgumentError){ @ts.account_information('foo') }
  end

  test "set_account_information basic functionality" do
    assert_respond_to(@ts, :set_account_information)
  end

  test "set_account_information works as expected" do
    assert_nothing_raised{ @ts.set_account_information('test', 'XXXX') }
    assert_equal('test', @ts.account_information)
  end

  test "set_account_information requires two arguments" do
    assert_raise(ArgumentError){ @ts.set_account_information }
    assert_raise(ArgumentError){ @ts.set_account_information('x') }
    assert_raise(ArgumentError){ @ts.set_account_information('x', 'y', 'z') }
  end

  test "arguments to set_account_information must be strings" do
    assert_raise(TypeError){ @ts.set_account_information(1, 'XXX') }
    assert_raise(TypeError){ @ts.set_account_information('x', 1) }
  end

  test "activate method basic functionality" do
    assert_respond_to(@ts, :activate)
  end

  test "activate behaves as expected" do
      assert_nothing_raised{ @ts.activate(@task) }
   end

  test "calling activate on the same object multiple times has no effect" do
    assert_nothing_raised{ @ts.activate(@task) }
    assert_nothing_raised{ @ts.activate(@task) }
  end

  test "activate requires a single argument" do
    assert_raise(ArgumentError){ @ts.activate }
    assert_raise(ArgumentError){ @ts.activate('foo', 'bar') }
  end

  test "activate requires a string argument" do
    assert_raise(TypeError){ @ts.activate(1) }
  end

  test "attempting to activate a bad task results in an error" do
    assert_raise(TaskScheduler::Error){ @ts.activate('bogus') }
  end

  test "application_name basic functionality" do
    assert_respond_to(@ts, :application_name)
    assert_nothing_raised{ @ts.application_name }
  end

  test "application_name returns a string or nil" do
    assert_kind_of([String, NilClass], @ts.application_name)
  end

  test "application name does not accept any arguments" do
    assert_raises(ArgumentError){ @ts.application_name('bogus') }
  end

=begin
   def test_set_application_name
      assert_respond_to(@ts, :application_name=)
      assert_nothing_raised{ @ts.application_name = "notepad.exe" }
   end

   def test_set_application_name_expected_errors
      assert_raise(ArgumentError){ @ts.send(:application_name=) }
      assert_raise(TypeError){ @ts.application_name = 1 }
   end

   def test_get_comment
      assert_respond_to(@ts, :comment)
      assert_nothing_raised{ @ts.comment }
      assert_kind_of(String, @ts.comment)
  end

   def test_get_comment_expected_errors
      assert_raise(ArgumentError){ @ts.comment('test') }
   end

   def test_set_comment
      assert_respond_to(@ts, :comment=)
      assert_nothing_raised{ @ts.comment = "test" }
   end

   def test_set_comment_expected_errors
      assert_raise(ArgumentError){ @ts.send(:comment=) }
      assert_raise(TypeError){ @ts.comment = 1 }
   end

   def test_get_creator
      assert_respond_to(@ts, :creator)
      assert_nothing_raised{ @ts.creator }
      assert_kind_of(String, @ts.creator)
   end

   def test_get_creator_expected_errors
      assert_raise(ArgumentError){ @ts.creator('foo') }
   end

   def test_set_creator
      assert_respond_to(@ts, :creator=)
      assert_nothing_raised{ @ts.creator = "Test Creator" }
   end

   def test_set_creator_expected_errors
      assert_raise(ArgumentError){ @ts.send(:creator=) }
      assert_raise(TypeError){ @ts.creator = 1 }
   end

   def test_delete
      assert_respond_to(@ts, :delete)
      assert_nothing_raised{ @ts.delete(@task) }
   end

   def test_delete_expected_errors
      assert_raise(ArgumentError){ @ts.delete }
      assert_raise(TaskScheduler::Error){ @ts.delete("foofoo") }
   end

   def test_delete_trigger
      assert_respond_to(@ts, :delete_trigger)
      assert_equal(0, @ts.delete_trigger(0))
  end

   # TODO: Figure out why the last two fail
   def test_delete_trigger_expected_errors
      assert_raise(ArgumentError){ @ts.delete }
      #assert_raise(TypeError){ @ts.delete('test') }
      #assert_raise(TaskScheduler::Error){ @ts.delete(-1) }
   end
=end

   test "enum basic functionality" do
     assert_respond_to(@ts, :enum)
     assert_nothing_raised{ @ts.enum }
   end

   test "enum method returns an array of strings" do
     assert_kind_of(Array, @ts.enum)
     assert_kind_of(String, @ts.enum.first)
   end

   test "tasks is an alias for enum" do
     assert_respond_to(@ts, :tasks)
     assert_alias_method(@ts, :tasks, :enum)
   end

   test "enum method does not accept any arguments" do
     assert_raise(ArgumentError){ @ts.enum(1) }
   end

=begin
   def test_exists_basic
      assert_respond_to(@ts, :exists?)
      assert_boolean(@ts.exists?(@task))
   end

   def test_exists
      assert_true(@ts.exists?(@task))
      assert_false(@ts.exists?('bogusXYZ'))
   end

   def test_exists_expected_errors
      assert_raise(ArgumentError){ @ts.exists? }
      assert_raise(ArgumentError){ @ts.exists?('foo', 'bar') }
   end

   def test_exit_code
      assert_respond_to(@ts, :exit_code)
      assert_nothing_raised{ @ts.exit_code }
      assert_kind_of(Fixnum, @ts.exit_code)
   end

   def test_exit_code_expected_errors
      assert_raise(ArgumentError){ @ts.exit_code(true) }
      assert_raise(NoMethodError){ @ts.exit_code = 1 }
   end

   def test_get_flags
      assert_respond_to(@ts, :flags)
      assert_nothing_raised{ @ts.flags }
      assert_kind_of(Fixnum, @ts.flags)
   end

   def test_get_flags_expected_errors
      assert_raise(ArgumentError){ @ts.flags(1) }
   end

   def test_set_flags
      assert_respond_to(@ts, :flags=)
      assert_nothing_raised{ @ts.flags = TaskScheduler::DELETE_WHEN_DONE }
   end

   def test_set_flags_expected_errors
      assert_raise(ArgumentError){ @ts.send(:flags=) }
      assert_raise(TypeError){ @ts.flags = 'test' }
   end

   # TODO: Why does setting the host fail?
   def test_set_machine
      assert_respond_to(@ts, :host=)
      assert_respond_to(@ts, :machine=)
      #assert_nothing_raised{ @ts.machine = @@host }
   end

   def test_get_max_run_time
      assert_respond_to(@ts, :max_run_time)
      assert_nothing_raised{ @ts.max_run_time }
      assert_kind_of(Fixnum, @ts.max_run_time)
   end

   def test_get_max_run_time_expected_errors
      assert_raise(ArgumentError){ @ts.max_run_time(true) }
   end

   def test_set_max_run_time
      assert_respond_to(@ts, :max_run_time=)
      assert_nothing_raised{ @ts.max_run_time = 20000 }
   end

   def test_set_max_run_time_expected_errors
      assert_raise(ArgumentError){ @ts.send(:max_run_time=) }
      assert_raise(TypeError){ @ts.max_run_time = true }
   end

   def test_most_recent_run_time
      assert_respond_to(@ts, :most_recent_run_time)
      assert_nothing_raised{ @ts.most_recent_run_time }
      assert_nil(@ts.most_recent_run_time)
   end

   def test_most_recent_run_time_expected_errors
      assert_raise(ArgumentError){ @ts.most_recent_run_time(true) }
      assert_raise(NoMethodError){ @ts.most_recent_run_time = Time.now }
   end

   def test_new_work_item
      assert_respond_to(@ts, :new_work_item)
      assert_nothing_raised{ @ts.new_work_item('bar', @trigger) }
   end

   def test_new_work_item_expected_errors
      assert_raise(ArgumentError){ @ts.new_work_item('test', {:bogus => 1}) }
   end

   def test_new_task_alias
      assert_respond_to(@ts, :new_task)
      assert_equal(true, @ts.method(:new_task) == @ts.method(:new_work_item))
   end

   def test_new_work_item_expected_errors
      assert_raise(ArgumentError){ @ts.new_work_item }
      assert_raise(ArgumentError){ @ts.new_work_item('bar') }
      assert_raise(TypeError){ @ts.new_work_item(1, 'bar') }
      assert_raise(TypeError){ @ts.new_work_item('bar', 1) }
   end

   def test_next_run_time
      assert_respond_to(@ts, :next_run_time)
      assert_nothing_raised{ @ts.next_run_time }
      assert_kind_of(Time, @ts.next_run_time)
   end

   def test_next_run_time_expected_errors
      assert_raise(ArgumentError){ @ts.next_run_time(true) }
      assert_raise(NoMethodError){ @ts.next_run_time = Time.now }
   end

   def test_get_parameters
      assert_respond_to(@ts, :parameters)
      assert_nothing_raised{ @ts.parameters }
   end

   def test_get_parameters_expected_errors
      assert_raise(ArgumentError){ @ts.parameters('test') }
   end

   def set_parameters
      assert_respond_to(@ts, :parameters=)
      assert_nothing_raised{ @ts.parameters = "somefile.txt" }
   end

   def set_parameters_expected_errors
      assert_raise(ArgumentError){ @ts.send(:parameters=) }
      assert_raise(TypeError){ @ts.parameters = 1 }
   end

   def test_get_priority
      assert_respond_to(@ts, :priority)
      assert_nothing_raised{ @ts.priority }
      assert_kind_of(String, @ts.priority)
   end

   def test_get_priority_expected_errors
      assert_raise(ArgumentError){ @ts.priority(true) }
   end

   def test_set_priority
      assert_respond_to(@ts, :priority=)
      assert_nothing_raised{ @ts.priority = TaskScheduler::NORMAL }
   end

   def test_set_priority_expected_errors
      assert_raise(ArgumentError){ @ts.send(:priority=) }
      assert_raise(TypeError){ @ts.priority = 'alpha' }
   end

   # TODO: Find a harmless way to test this.
   def test_run
      assert_respond_to(@ts, :run)
   end

   def test_run_expected_errors
      assert_raise(ArgumentError){ @ts.run(true) }
      assert_raise(NoMethodError){ @ts.run = true }
   end

   def test_save
      assert_respond_to(@ts, :save)
      assert_nothing_raised{ @ts.save }
   end

   def test_save_custom_file
      assert_nothing_raised{ @ts.save(@job_file) }
      assert_equal(true, File.exists?(@job_file))
   end

   def test_save_expected_errors
      assert_raise(TypeError){ @ts.save(true) }
      assert_raise(ArgumentError){ @ts.save(@job_file, true) }
      assert_raise(NoMethodError){ @ts.save = true }
      assert_raise(TaskScheduler::Error){ @ts.save; @ts.save }
   end

   def test_status
      assert_respond_to(@ts, :status)
      assert_nothing_raised{ @ts.status }
      assert_equal('not scheduled', @ts.status)
   end

   def test_status_expected_errors
      assert_raise(ArgumentError){ @ts.status(true) }
      assert_raise(NoMethodError){ @ts.status = true }
   end

   def test_terminate
      assert_respond_to(@ts, :terminate)
   end

   def test_terminate_expected_errors
      assert_raise(ArgumentError){ @ts.terminate(true) }
      assert_raise(NoMethodError){ @ts.terminate = true }
      assert_raise(TaskScheduler::Error){ @ts.terminate } # It's not running
   end

   def test_get_trigger
      assert_respond_to(@ts, :trigger)
      assert_nothing_raised{ @ts.trigger(0) }
      assert_kind_of(Hash, @ts.trigger(0))
   end

   def test_get_trigger_expected_errors
      assert_raises(ArgumentError){ @ts.trigger }
      assert_raises(TaskScheduler::Error){ @ts.trigger(9999) }
   end

   def test_set_trigger
      assert_respond_to(@ts, :trigger=)
      assert_nothing_raised{ @ts.trigger = @trigger }
   end

   def test_set_trigger_expected_errors
      assert_raises(TypeError){ @ts.trigger = 'blah' }
   end

   def test_add_trigger
      assert_respond_to(@ts, :add_trigger)
      assert_nothing_raised{ @ts.add_trigger(0, @trigger) }
   end

   def test_add_trigger_expected_errors
      assert_raises(ArgumentError){ @ts.add_trigger }
      assert_raises(ArgumentError){ @ts.add_trigger(0) }
      assert_raises(TypeError){ @ts.add_trigger(0, 'foo') }
   end

   def test_trigger_count
      assert_respond_to(@ts, :trigger_count)
      assert_nothing_raised{ @ts.trigger_count }
      assert_kind_of(Fixnum, @ts.trigger_count)
   end

   def test_trigger_count_expected_errors
      assert_raise(ArgumentError){ @ts.trigger_count(true) }
      assert_raise(NoMethodError){ @ts.trigger_count = 1 }
   end

   def test_trigger_delete
      assert_respond_to(@ts, :delete_trigger)
      assert_nothing_raised{ @ts.delete_trigger(0) }
  end

   def test_trigger_delete_expected_errors
      assert_raise(ArgumentError){ @ts.delete_trigger }
      assert_raise(TaskScheduler::Error){ @ts.delete_trigger(9999) }
   end

   def test_trigger_string
      assert_respond_to(@ts, :trigger_string)
      assert_nothing_raised{ @ts.trigger_string(0) }
      assert_equal('At 7:14 AM every day, starting 4/11/2009', @ts.trigger_string(0))
   end

   def test_trigger_string_expected_errors
      assert_raise(ArgumentError){ @ts.trigger_string }
      assert_raise(ArgumentError){ @ts.trigger_string(0, 0) }
      assert_raise(TypeError){ @ts.trigger_string('alpha') }
      assert_raise(TaskScheduler::Error){ @ts.trigger_string(9999) }
   end

   def test_get_working_directory
      assert_respond_to(@ts, :working_directory)
      assert_nothing_raised{ @ts.working_directory }
      assert_kind_of(String, @ts.working_directory)
   end

   def test_get_working_directory_expected_errors
      assert_raise(ArgumentError){ @ts.working_directory(true) }
   end

   def test_set_working_directory
      assert_respond_to(@ts, :working_directory=)
      assert_nothing_raised{ @ts.working_directory = "C:\\" }
   end

   def test_set_working_directory_expected_errors
      assert_raise(ArgumentError){ @ts.send(:working_directory=) }
      assert_raise(TypeError){ @ts.working_directory = 1 }
   end
=end

  test "expected constants are defined" do
    assert_not_nil(TaskScheduler::MONDAY)
    assert_not_nil(TaskScheduler::TUESDAY)
    assert_not_nil(TaskScheduler::WEDNESDAY)
    assert_not_nil(TaskScheduler::THURSDAY)
    assert_not_nil(TaskScheduler::FRIDAY)
    assert_not_nil(TaskScheduler::SATURDAY)
    assert_not_nil(TaskScheduler::SUNDAY)

    assert_not_nil(TaskScheduler::JANUARY)
    assert_not_nil(TaskScheduler::FEBRUARY)
    assert_not_nil(TaskScheduler::MARCH)
    assert_not_nil(TaskScheduler::APRIL)
    assert_not_nil(TaskScheduler::MAY)
    assert_not_nil(TaskScheduler::JUNE)
    assert_not_nil(TaskScheduler::JULY)
    assert_not_nil(TaskScheduler::AUGUST)
    assert_not_nil(TaskScheduler::SEPTEMBER)
    assert_not_nil(TaskScheduler::OCTOBER)
    assert_not_nil(TaskScheduler::NOVEMBER)
    assert_not_nil(TaskScheduler::DECEMBER)

    assert_not_nil(TaskScheduler::ONCE)
    assert_not_nil(TaskScheduler::DAILY)
    assert_not_nil(TaskScheduler::WEEKLY)
    assert_not_nil(TaskScheduler::MONTHLYDATE)
    assert_not_nil(TaskScheduler::MONTHLYDOW)

    assert_not_nil(TaskScheduler::ON_IDLE)
    assert_not_nil(TaskScheduler::AT_SYSTEMSTART)
    assert_not_nil(TaskScheduler::AT_LOGON)

    assert_not_nil(TaskScheduler::INTERACTIVE)
    assert_not_nil(TaskScheduler::DELETE_WHEN_DONE)
    assert_not_nil(TaskScheduler::DISABLED)
    assert_not_nil(TaskScheduler::START_ONLY_IF_IDLE)
    assert_not_nil(TaskScheduler::KILL_ON_IDLE_END)
    assert_not_nil(TaskScheduler::DONT_START_IF_ON_BATTERIES)
    assert_not_nil(TaskScheduler::KILL_IF_GOING_ON_BATTERIES)
    assert_not_nil(TaskScheduler::HIDDEN)
    assert_not_nil(TaskScheduler::RESTART_ON_IDLE_RESUME)
    assert_not_nil(TaskScheduler::SYSTEM_REQUIRED)
    assert_not_nil(TaskScheduler::FLAG_HAS_END_DATE)
    assert_not_nil(TaskScheduler::FLAG_KILL_AT_DURATION_END)
    assert_not_nil(TaskScheduler::FLAG_DISABLED)
    assert_not_nil(TaskScheduler::MAX_RUN_TIMES)

    #assert_not_nil(TaskScheduler::IDLE)
    #assert_not_nil(TaskScheduler::NORMAL)
    #assert_not_nil(TaskScheduler::HIGH)
    #assert_not_nil(TaskScheduler::REALTIME)
    #assert_not_nil(TaskScheduler::ABOVE_NORMAL)
    #assert_not_nil(TaskScheduler::BELOW_NORMAL)
  end

  def teardown
    File.delete(@job_file) if File.exists?(@job_file)
    @ts.delete('foo') rescue nil
    @ts = nil
    @trigger = nil
    @job_file = nil
  end

  def self.shutdown
    @@host = nil
    @@user = nil
  end
end
