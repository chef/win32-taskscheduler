##########################################################################
# test_taskscheduler.rb
#
# Test suite for the win32-taskscheduler package. You should run this
# via the 'rake test' task.
##########################################################################
require "win32/taskscheduler"
require "win32/security"
require "socket"
require "test-unit"
require "etc"
include Win32

class TC_TaskScheduler < Test::Unit::TestCase
  def self.startup
    @@host = Socket.gethostname
    @@user = Etc.getlogin
    @@elevated = Win32::Security.elevated_security?
  end

  def setup
    @task = "foo"

    @trigger = {
      start_year: 2015,
      start_month: 4,
      start_day: 11,
      start_hour: 7,
      start_minute: 14,
      trigger_type: TaskScheduler::DAILY,
      type: { days_interval: 1 },
    }

    @ts = TaskScheduler.new
  end

  # Helper method
  def setup_task
    @ts.new_work_item(@task, @trigger)
    @ts.activate(@task)
  end

  test "version constant is set to expected value" do
    assert_equal("0.3.2", TaskScheduler::VERSION)
  end

  test "account_information method basic functionality" do
    assert_respond_to(@ts, :account_information)
  end

  test "account_information returns the task owner" do
    setup_task
    assert_equal("#{@@host}\\#{@@user}", @ts.account_information)
  end

  test "account_information returns nil if no task has been activated" do
    assert_nil(@ts.account_information)
  end

  test "account_information does not accept any arguments" do
    assert_raise(ArgumentError) { @ts.account_information("foo") }
  end

  test "set_account_information basic functionality" do
    assert_respond_to(@ts, :set_account_information)
  end

  test "set_account_information works as expected" do
    setup_task
    # assert_nothing_raised{ @ts.set_account_information(@@user, 'XXXX') }
    # assert_equal('test', @ts.account_information)
  end

  test "set_account_information requires two arguments" do
    setup_task
    assert_raise(ArgumentError) { @ts.set_account_information }
    assert_raise(ArgumentError) { @ts.set_account_information("x") }
    assert_raise(ArgumentError) { @ts.set_account_information("x", "y", "z") }
  end

  test "arguments to set_account_information must be strings" do
    setup_task
    assert_raise(TypeError) { @ts.set_account_information(1, "XXX") }
    assert_raise(TypeError) { @ts.set_account_information("x", 1) }
  end

  test "activate method basic functionality" do
    assert_respond_to(@ts, :activate)
  end

  test "activate behaves as expected" do
    @ts.new_work_item(@task, @trigger)
    assert_nothing_raised { @ts.activate(@task) }
  end

  test "calling activate on the same object multiple times has no effect" do
    @ts.new_work_item(@task, @trigger)
    assert_nothing_raised { @ts.activate(@task) }
    assert_nothing_raised { @ts.activate(@task) }
  end

  test "activate requires a single argument" do
    assert_raise(ArgumentError) { @ts.activate }
    assert_raise(ArgumentError) { @ts.activate("foo", "bar") }
  end

  test "activate requires a string argument" do
    assert_raise(TypeError) { @ts.activate(1) }
  end

  test "attempting to activate a bad task results in an error" do
    assert_raise(TaskScheduler::Error) { @ts.activate("bogus") }
  end

  test "application_name basic functionality" do
    assert_respond_to(@ts, :application_name)
  end

  test "application_name returns a string or nil" do
    setup_task
    assert_nothing_raised { @ts.application_name }
    assert_kind_of([String, NilClass], @ts.application_name)
  end

  test "application_name does not accept any arguments" do
    assert_raises(ArgumentError) { @ts.application_name("bogus") }
  end

  test "application_name= basic functionality" do
    assert_respond_to(@ts, :application_name=)
  end

  test "application_name= works as expected" do
    setup_task
    assert_nothing_raised { @ts.application_name = "notepad.exe" }
    assert_equal("notepad.exe", @ts.application_name)
  end

  test "application_name= requires a string argument" do
    assert_raise(TypeError) { @ts.application_name = 1 }
  end

  test "comment method basic functionality" do
    assert_respond_to(@ts, :comment)
  end

  test "comment method returns a string" do
    setup_task
    assert_nothing_raised { @ts.comment }
    assert_kind_of(String, @ts.comment)
  end

  test "comment method does not accept any arguments" do
    assert_raise(ArgumentError) { @ts.comment("test") }
  end

  test "comment= method basic functionality" do
    assert_respond_to(@ts, :comment=)
  end

  test "comment= works as expected" do
    setup_task
    assert_nothing_raised { @ts.comment = "test" }
    assert_equal("test", @ts.comment)
  end

  test "comment= method requires a string argument" do
    assert_raise(TypeError) { @ts.comment = 1 }
  end

  test "creator method basic functionality" do
    setup_task
    assert_respond_to(@ts, :creator)
    assert_nothing_raised { @ts.creator }
    assert_kind_of(String, @ts.creator)
  end

  test "creator method returns expected value" do
    setup_task
    assert_equal("", @ts.creator)
  end

  test "creator method does not accept any arguments" do
    assert_raise(ArgumentError) { @ts.creator("foo") }
  end

  test "creator= method basic functionality" do
    assert_respond_to(@ts, :creator=)
  end

  test "creator= method works as expected" do
    setup_task
    assert_nothing_raised { @ts.creator = "Test Creator" }
    assert_equal("Test Creator", @ts.creator)
  end

  test "creator= method requires a string argument" do
    assert_raise(TypeError) { @ts.creator = 1 }
  end

  test "delete method basic functionality" do
    assert_respond_to(@ts, :delete)
  end

  test "delete method works as expected" do
    setup_task
    assert_nothing_raised { @ts.delete(@task) }
    assert_false(@ts.exists?(@task))
  end

  test "delete method requires a single argument" do
    assert_raise(ArgumentError) { @ts.delete }
  end

  test "delete method raises an error if the task does not exist" do
    assert_raise(TaskScheduler::Error) { @ts.delete("foofoo") }
  end

  test "enum basic functionality" do
    assert_respond_to(@ts, :enum)
    assert_nothing_raised { @ts.enum }
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
    assert_raise(ArgumentError) { @ts.enum(1) }
  end

  test "exists? basic functionality" do
    assert_respond_to(@ts, :exists?)
    assert_boolean(@ts.exists?(@task))
  end

  test "exists? returns expected value" do
    setup_task
    assert_true(@ts.exists?(@task))
    assert_false(@ts.exists?("bogusXYZ"))
  end

  test "exists? method requires a single argument" do
    assert_raise(ArgumentError) { @ts.exists? }
    assert_raise(ArgumentError) { @ts.exists?("foo", "bar") }
  end

  test "exit_code basic functionality" do
    setup_task
    assert_respond_to(@ts, :exit_code)
    assert_nothing_raised { @ts.exit_code }
    assert_kind_of(Integer, @ts.exit_code)
  end

  test "exit_code takes no arguments" do
    assert_raise(ArgumentError) { @ts.exit_code(true) }
    assert_raise(NoMethodError) { @ts.exit_code = 1 }
  end

  test "machine= basic functionality" do
    assert_respond_to(@ts, :machine=)
  end

  test "machine= works as expected" do
    omit_unless(@@elevated)
    setup_task
    assert_nothing_raised { @ts.machine = @@host }
    assert_equal(@@host, @ts.machine)
  end

  test "host= is an alias for machine=" do
    assert_alias_method(@ts, :machine=, :host=)
  end

  test "set_machine basic functionality" do
    assert_respond_to(@ts, :set_machine)
  end

  test "set_host is an alias for set_machine" do
    assert_alias_method(@ts, :set_machine, :set_host)
  end

  test "max_run_time basic functionality" do
    assert_respond_to(@ts, :max_run_time)
  end

  test "max_run_time works as expected" do
    setup_task
    assert_nothing_raised { @ts.max_run_time }
    assert_kind_of(Integer, @ts.max_run_time)
  end

  test "max_run_time accepts no arguments" do
    assert_raise(ArgumentError) { @ts.max_run_time(true) }
  end

  test "max_run_time= basic functionality" do
    assert_respond_to(@ts, :max_run_time=)
  end

  test "max_run_time= works as expected" do
    setup_task
    assert_nothing_raised { @ts.max_run_time = 20_000 }
    assert_equal(20_000, @ts.max_run_time)
  end

  test "max_run_time= requires a numeric argument" do
    assert_raise(TypeError) { @ts.max_run_time = true }
  end

  test "most_recent_run_time basic functionality" do
    setup_task
    assert_respond_to(@ts, :most_recent_run_time)
    assert_nothing_raised { @ts.most_recent_run_time }
  end

  test "most_recent_run_time is nil if task hasn't run" do
    setup_task
    assert_nil(@ts.most_recent_run_time)
  end

  test "most_recent_run_time does not accept any arguments" do
    assert_raise(ArgumentError) { @ts.most_recent_run_time(true) }
    assert_raise(NoMethodError) { @ts.most_recent_run_time = Time.now }
  end

  test "new_work_item basic functionality" do
    assert_respond_to(@ts, :new_work_item)
  end

  test "new_work_item accepts a trigger/hash" do
    assert_nothing_raised { @ts.new_work_item("bar", @trigger) }
  end

  test "new_work_item fails if a bogus trigger key is present" do
    assert_raise(ArgumentError) { @ts.new_work_item("test", bogus: 1) }
  end

  test "new_task is an alias for new_work_item" do
    assert_respond_to(@ts, :new_task)
    assert_alias_method(@ts, :new_task, :new_work_item)
  end

  test "new_work_item requires a task name and a trigger" do
    assert_raise(ArgumentError) { @ts.new_work_item }
    assert_raise(ArgumentError) { @ts.new_work_item("bar") }
  end

  test "new_work_item expects a string for the first argument" do
    assert_raise(TypeError) { @ts.new_work_item(1, @trigger) }
  end

  test "new_work_item expects a hash for the second argument" do
    assert_raise(TypeError) { @ts.new_work_item(@task, 1) }
  end

  test "next_run_time basic functionality" do
    assert_respond_to(@ts, :next_run_time)
  end

  test "next_run_time returns a Time object" do
    setup_task
    assert_nothing_raised { @ts.next_run_time }
    assert_kind_of(Time, @ts.next_run_time)
  end

  test "next_run_time does not take any arguments" do
    assert_raise(ArgumentError) { @ts.next_run_time(true) }
    assert_raise(NoMethodError) { @ts.next_run_time = Time.now }
  end

  test "parameters basic functionality" do
    setup_task
    assert_respond_to(@ts, :parameters)
    assert_nothing_raised { @ts.parameters }
  end

  test "parameters method does not take any arguments" do
    assert_raise(ArgumentError) { @ts.parameters("test") }
  end

  test "parameters= basic functionality" do
    assert_respond_to(@ts, :parameters=)
  end

  test "parameters= works as expected" do
    setup_task
    assert_nothing_raised { @ts.parameters = "somefile.txt" }
    assert_equal("somefile.txt", @ts.parameters)
  end

  test "parameters= requires a string argument" do
    assert_raise(ArgumentError) { @ts.send(:parameters=) }
    assert_raise(TypeError) { @ts.parameters = 1 }
  end

  test "priority basic functionality" do
    assert_respond_to(@ts, :priority)
  end

  test "priority returns the expected value" do
    setup_task
    assert_nothing_raised { @ts.priority }
    assert_kind_of(String, @ts.priority)
  end

  test "priority method does not accept any arguments" do
    assert_raise(ArgumentError) { @ts.priority(true) }
  end

  test "priority= basic functionality" do
    assert_respond_to(@ts, :priority=)
  end

  test "priority= works as expected" do
    setup_task
    assert_nothing_raised { @ts.priority = TaskScheduler::NORMAL }
    assert_equal("normal", @ts.priority)
  end

  test "priority= requires a numeric argument" do
    assert_raise(TypeError) { @ts.priority = "alpha" }
  end

  # TODO: Find a harmless way to test this.
  test "run basic functionality" do
    assert_respond_to(@ts, :run)
  end

  test "run does not accept any arguments" do
    assert_raise(ArgumentError) { @ts.run(true) }
    assert_raise(NoMethodError) { @ts.run = true }
  end

  test "save basic functionality" do
    assert_respond_to(@ts, :save)
  end

  test "status basic functionality" do
    setup_task
    assert_respond_to(@ts, :status)
    assert_nothing_raised { @ts.status }
    assert_kind_of(String, @ts.status)
  end

  test "status returns the expected value" do
    setup_task
    assert_equal("ready", @ts.status)
  end

  test "status does not accept any arguments" do
    assert_raise(ArgumentError) { @ts.status(true) }
  end

  test "terminate basic functionality" do
    assert_respond_to(@ts, :terminate)
  end

  test "terminate does not accept any arguments" do
    assert_raise(ArgumentError) { @ts.terminate(true) }
    assert_raise(NoMethodError) { @ts.terminate = true }
  end

  test "calling terminate on a task that isn't running raises an error" do
    assert_raise(TaskScheduler::Error) { @ts.terminate }
  end

  test "trigger basic functionality" do
    setup_task
    assert_respond_to(@ts, :trigger)
    assert_nothing_raised { @ts.trigger(0) }
  end

  test "trigger returns a hash object" do
    setup_task
    assert_kind_of(Hash, @ts.trigger(0))
  end

  test "trigger requires an index" do
    assert_raises(ArgumentError) { @ts.trigger }
  end

  test "trigger raises an error if the index is invalid" do
    assert_raises(TaskScheduler::Error) { @ts.trigger(9999) }
  end

  test "trigger= basic functionality" do
    assert_respond_to(@ts, :trigger=)
  end

  test "trigger= works as expected" do
    setup_task
    assert_nothing_raised { @ts.trigger = @trigger }
  end

  test "trigger= requires a hash argument" do
    assert_raises(TypeError) { @ts.trigger = "blah" }
  end

  test "add_trigger basic functionality" do
    assert_respond_to(@ts, :add_trigger)
  end

  test "add_trigger works as expected" do
    setup_task
    assert_nothing_raised { @ts.add_trigger(0, @trigger) }
  end

  test "add_trigger requires two arguments" do
    assert_raises(ArgumentError) { @ts.add_trigger }
    assert_raises(ArgumentError) { @ts.add_trigger(0) }
  end

  test "add_trigger reqiuires an integer for the first argument" do
    assert_raises(TypeError) { @ts.add_trigger("foo", @trigger) }
  end

  test "add_trigger reqiuires a hash for the second argument" do
    assert_raises(TypeError) { @ts.add_trigger(0, "foo") }
  end

  test "trigger_count basic functionality" do
    setup_task
    assert_respond_to(@ts, :trigger_count)
    assert_nothing_raised { @ts.trigger_count }
    assert_kind_of(Integer, @ts.trigger_count)
  end

  test "trigger_count returns the expected value" do
    setup_task
    assert_equal(1, @ts.trigger_count)
  end

  test "trigger_count does not accept any arguments" do
    assert_raise(ArgumentError) { @ts.trigger_count(true) }
    assert_raise(NoMethodError) { @ts.trigger_count = 1 }
  end

  test "delete_trigger basic functionality" do
    assert_respond_to(@ts, :delete_trigger)
  end

  test "delete_trigger works as expected" do
    setup_task
    assert_nothing_raised { @ts.delete_trigger(0) }
    assert_equal(0, @ts.trigger_count)
  end

  test "passing a bad index to delete_trigger will raise an error" do
    assert_raise(TaskScheduler::Error) { @ts.delete_trigger(9999) }
  end

  test "delete_trigger requires at least one argument" do
    assert_raise(ArgumentError) { @ts.delete_trigger }
  end

  test "trigger_string basic functionality" do
    setup_task
    assert_respond_to(@ts, :trigger_string)
    assert_nothing_raised { @ts.trigger_string(0) }
  end

  test "trigger_string returns the expected value" do
    setup_task
    assert_equal("Starting 2015-04-11T07:14:00", @ts.trigger_string(0))
  end

  test "trigger_string requires a single argument" do
    assert_raise(ArgumentError) { @ts.trigger_string }
    assert_raise(ArgumentError) { @ts.trigger_string(0, 0) }
  end

  test "trigger_string requires a numeric argument" do
    assert_raise(TypeError) { @ts.trigger_string("alpha") }
  end

  test "trigger_string raises an error if the index is invalid" do
    assert_raise(TaskScheduler::Error) { @ts.trigger_string(9999) }
  end

  test "working_directory basic functionality" do
    setup_task
    assert_respond_to(@ts, :working_directory)
    assert_nothing_raised { @ts.working_directory }
    assert_kind_of(String, @ts.working_directory)
  end

  test "working_directory takes no arguments" do
    assert_raise(ArgumentError) { @ts.working_directory(true) }
  end

  test "working_directory= basic functionality" do
    assert_respond_to(@ts, :working_directory=)
  end

  test "working_directory= works as expected" do
    setup_task
    assert_nothing_raised { @ts.working_directory = "C:\\" }
  end

  test "working_directory= requires a string argument" do
    setup_task
    assert_raise(TypeError) { @ts.working_directory = 1 }
  end

  test "expected day constants are defined" do
    assert_not_nil(TaskScheduler::MONDAY)
    assert_not_nil(TaskScheduler::TUESDAY)
    assert_not_nil(TaskScheduler::WEDNESDAY)
    assert_not_nil(TaskScheduler::THURSDAY)
    assert_not_nil(TaskScheduler::FRIDAY)
    assert_not_nil(TaskScheduler::SATURDAY)
    assert_not_nil(TaskScheduler::SUNDAY)
  end

  test "expected month constants are defined" do
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
  end

  test "expected repeat constants are defined" do
    assert_not_nil(TaskScheduler::ONCE)
    assert_not_nil(TaskScheduler::DAILY)
    assert_not_nil(TaskScheduler::WEEKLY)
    assert_not_nil(TaskScheduler::MONTHLYDATE)
    assert_not_nil(TaskScheduler::MONTHLYDOW)
  end

  test "expected start constants are defined" do
    assert_not_nil(TaskScheduler::ON_IDLE)
    assert_not_nil(TaskScheduler::AT_SYSTEMSTART)
    assert_not_nil(TaskScheduler::AT_LOGON)
    assert_not_nil(TaskScheduler::ON_SESSION_STATE_CHANGE)
  end

  test "expected flag constants are defined" do
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
  end

  test "expected priority constants are defined" do
    assert_not_nil(TaskScheduler::IDLE)
    assert_not_nil(TaskScheduler::NORMAL)
    assert_not_nil(TaskScheduler::HIGH)
    assert_not_nil(TaskScheduler::REALTIME)
    assert_not_nil(TaskScheduler::ABOVE_NORMAL)
    assert_not_nil(TaskScheduler::BELOW_NORMAL)
  end

  test "expected session state constants are defined" do
    assert_not_nil(TaskScheduler::TASK_CONSOLE_CONNECT)
    assert_not_nil(TaskScheduler::TASK_CONSOLE_DISCONNECT)
    assert_not_nil(TaskScheduler::TASK_REMOTE_CONNECT)
    assert_not_nil(TaskScheduler::TASK_REMOTE_DISCONNECT)
    assert_not_nil(TaskScheduler::TASK_SESSION_LOCK)
    assert_not_nil(TaskScheduler::TASK_SESSION_UNLOCK)
  end

  def teardown
    @ts.delete(@task) if @ts.exists?(@task)
    @ts.delete("bar") if @ts.exists?("bar")
    @ts = nil
    @task = nil
    @trigger = nil
  end

  def self.shutdown
    @@host = nil
    @@user = nil
    @@elevated = nil
  end
end
