require_relative "taskscheduler/sid"
require_relative "taskscheduler/helper"
require_relative "taskscheduler/time_calc_helper"
require_relative "taskscheduler/constants"
require_relative "taskscheduler/version"
require "win32ole" unless defined?(WIN32OLE)
require "time" unless defined?(Time.zone_offset)
require "socket" unless defined?(Socket)
require "structured_warnings"

# The Win32 module serves as a namespace only
module Win32
  # The TaskScheduler class encapsulates a Windows scheduled task
  class TaskScheduler
    include Win32::TaskScheduler::Helper
    include Win32::TaskScheduler::TaskSchedulerConstants
    include Win32::TaskScheduler::TimeCalcHelper
    include Win32::TaskScheduler::SID

    # The Error class is typically raised if any TaskScheduler methods fail.
    class Error < StandardError; end

    # Shorthand constants

    IDLE = IDLE_PRIORITY_CLASS
    NORMAL = NORMAL_PRIORITY_CLASS
    HIGH = HIGH_PRIORITY_CLASS
    REALTIME = REALTIME_PRIORITY_CLASS
    BELOW_NORMAL = BELOW_NORMAL_PRIORITY_CLASS
    ABOVE_NORMAL = ABOVE_NORMAL_PRIORITY_CLASS

    ONCE = TASK_TIME_TRIGGER_ONCE
    DAILY = TASK_TIME_TRIGGER_DAILY
    WEEKLY = TASK_TIME_TRIGGER_WEEKLY
    MONTHLYDATE = TASK_TIME_TRIGGER_MONTHLYDATE
    MONTHLYDOW = TASK_TIME_TRIGGER_MONTHLYDOW

    ON_IDLE = TASK_EVENT_TRIGGER_ON_IDLE
    AT_SYSTEMSTART = TASK_EVENT_TRIGGER_AT_SYSTEMSTART
    AT_LOGON = TASK_EVENT_TRIGGER_AT_LOGON
    ON_SESSION_STATE_CHANGE = TASK_TRIGGER_SESSION_STATE_CHANGE
    FIRST_WEEK = TASK_FIRST_WEEK
    SECOND_WEEK = TASK_SECOND_WEEK
    THIRD_WEEK = TASK_THIRD_WEEK
    FOURTH_WEEK = TASK_FOURTH_WEEK
    LAST_WEEK = TASK_LAST_WEEK
    SUNDAY = TASK_SUNDAY
    MONDAY = TASK_MONDAY
    TUESDAY = TASK_TUESDAY
    WEDNESDAY = TASK_WEDNESDAY
    THURSDAY = TASK_THURSDAY
    FRIDAY = TASK_FRIDAY
    SATURDAY = TASK_SATURDAY
    JANUARY = TASK_JANUARY
    FEBRUARY = TASK_FEBRUARY
    MARCH = TASK_MARCH
    APRIL = TASK_APRIL
    MAY = TASK_MAY
    JUNE = TASK_JUNE
    JULY = TASK_JULY
    AUGUST = TASK_AUGUST
    SEPTEMBER = TASK_SEPTEMBER
    OCTOBER = TASK_OCTOBER
    NOVEMBER = TASK_NOVEMBER
    DECEMBER = TASK_DECEMBER

    INTERACTIVE = TASK_FLAG_INTERACTIVE
    DELETE_WHEN_DONE = TASK_FLAG_DELETE_WHEN_DONE
    DISABLED = TASK_FLAG_DISABLED
    START_ONLY_IF_IDLE = TASK_FLAG_START_ONLY_IF_IDLE
    KILL_ON_IDLE_END = TASK_FLAG_KILL_ON_IDLE_END
    DONT_START_IF_ON_BATTERIES = TASK_FLAG_DONT_START_IF_ON_BATTERIES
    KILL_IF_GOING_ON_BATTERIES = TASK_FLAG_KILL_IF_GOING_ON_BATTERIES
    RUN_ONLY_IF_DOCKED = TASK_FLAG_RUN_ONLY_IF_DOCKED
    HIDDEN = TASK_FLAG_HIDDEN
    RUN_IF_CONNECTED_TO_INTERNET = TASK_FLAG_RUN_IF_CONNECTED_TO_INTERNET
    RESTART_ON_IDLE_RESUME = TASK_FLAG_RESTART_ON_IDLE_RESUME
    SYSTEM_REQUIRED = TASK_FLAG_SYSTEM_REQUIRED
    RUN_ONLY_IF_LOGGED_ON = TASK_FLAG_RUN_ONLY_IF_LOGGED_ON

    FLAG_HAS_END_DATE = TASK_TRIGGER_FLAG_HAS_END_DATE
    FLAG_KILL_AT_DURATION_END = TASK_TRIGGER_FLAG_KILL_AT_DURATION_END
    FLAG_DISABLED = TASK_TRIGGER_FLAG_DISABLED

    MAX_RUN_TIMES = TASK_MAX_RUN_TIMES

    FIRST = TASK_FIRST
    SECOND = TASK_SECOND
    THIRD = TASK_THIRD
    FOURTH = TASK_FOURTH
    FIFTH = TASK_FIFTH
    SIXTH = TASK_SIXTH
    SEVENTH = TASK_SEVENTH
    EIGHTH = TASK_EIGHTH
    NINETH = TASK_NINETH
    TENTH = TASK_TENTH
    ELEVENTH = TASK_ELEVENTH
    TWELFTH = TASK_TWELFTH
    THIRTEENTH = TASK_THIRTEENTH
    FOURTEENTH = TASK_FOURTEENTH
    FIFTEENTH = TASK_FIFTEENTH
    SIXTEENTH = TASK_SIXTEENTH
    SEVENTEENTH = TASK_SEVENTEENTH
    EIGHTEENTH = TASK_EIGHTEENTH
    NINETEENTH = TASK_NINETEENTH
    TWENTIETH = TASK_TWENTIETH
    TWENTY_FIRST = TASK_TWENTY_FIRST
    TWENTY_SECOND = TASK_TWENTY_SECOND
    TWENTY_THIRD = TASK_TWENTY_THIRD
    TWENTY_FOURTH = TASK_TWENTY_FOURTH
    TWENTY_FIFTH = TASK_TWENTY_FIFTH
    TWENTY_SIXTH = TASK_TWENTY_SIXTH
    TWENTY_SEVENTH = TASK_TWENTY_SEVENTH
    TWENTY_EIGHTH = TASK_TWENTY_EIGHTH
    TWENTY_NINTH = TASK_TWENTY_NINTH
    THIRTYETH = TASK_THIRTYETH
    THIRTY_FIRST = TASK_THIRTY_FIRST
    LAST = TASK_LAST

    # :startdoc:

    attr_accessor :password
    attr_reader :host

    def root_path(path = "\\")
      path
    end

    # Returns a new TaskScheduler object, attached to +folder+. If that
    # folder does not exist, but the +force+ option is set to true, then
    # it will be created. Otherwise an error will be raised. The default
    # is to use the root folder.
    #
    # If +task+ and +trigger+ are present, then a new task is generated
    # as well. This is effectively the same as .new + #new_work_item.
    #
    def initialize(task = nil, trigger = nil, folder = root_path, force = false)
      @folder = folder
      @force  = force

      @host     = Socket.gethostname
      @task     = nil
      @password = nil

      raise ArgumentError, "invalid folder" unless folder.include?("\\")

      unless [TrueClass, FalseClass].include?(force.class)
        raise TypeError, "invalid force value"
      end

      begin
        @service = WIN32OLE.new("Schedule.Service")
      rescue WIN32OLERuntimeError => err
        raise Error, err.inspect
      end

      @service.Connect

      if folder != root_path
        begin
          @root = @service.GetFolder(folder)
        rescue WIN32OLERuntimeError => err
          if force
            @root = @service.GetFolder(root_path)
            @root = @root.CreateFolder(folder)
          else
            raise ArgumentError, "folder '#{folder}' not found"
          end
        end
      else
        @root = @service.GetFolder(folder)
      end

      new_work_item(task, trigger) if task && trigger
    end

    # Returns an array of scheduled task names.
    #
    def enum
      # Get the task folder that contains the tasks.
      taskCollection = @root.GetTasks(0)

      array = []

      taskCollection.each do |registeredTask|
        array << registeredTask.Name
      end

      array
    end

    alias tasks enum

    # Returns whether or not the specified task exists.
    #
    def exists?(full_task_path)
      path = nil
      task_name = nil

      if full_task_path.include?("\\")
        *path, task_name = full_task_path.split("\\")
      else
        task_name = full_task_path
      end

      folder = path.nil? ? root_path : path.join("\\")

      begin
        root = @service.GetFolder(folder)
      rescue WIN32OLERuntimeError => err
        return false
      end

      if root.nil?
        false
      else
        begin
          task = root.GetTask(task_name)
          task && task.Name == task_name
        rescue WIN32OLERuntimeError => err
          false
        end
      end
    end

    # Return the sepcified task if exist
    #
    def get_task(task)
      raise TypeError unless task.is_a?(String)

      begin
        registeredTask = @root.GetTask(task)
        @task = registeredTask
      rescue WIN32OLERuntimeError => err
        raise Error, ole_error("activate", err)
      end
    end

    # Activate the specified task.
    #
    def activate(task)
      raise TypeError unless task.is_a?(String)

      begin
        registeredTask = @root.GetTask(task)
        registeredTask.Enabled = 1
        @task = registeredTask
      rescue WIN32OLERuntimeError => err
        raise Error, ole_error("activate", err)
      end
    end

    # Delete the specified task name.
    #
    def delete(task)
      raise TypeError unless task.is_a?(String)

      begin
        @root.DeleteTask(task, 0)
      rescue WIN32OLERuntimeError => err
        raise Error, ole_error("DeleteTask", err)
      end
    end

    # Execute the current task.
    #
    def run
      check_for_active_task
      @task.run(nil)
    end

    # This method no longer has any effect. It is a no-op that remains for
    # backwards compatibility. It will be removed in 0.4.0.
    #
    def save(_file = nil)
      warn DeprecatedMethodWarning, "this method is no longer necessary"
      check_for_active_task
      # Do nothing, deprecated.
    end

    # Terminate (stop) the current task.
    #
    def terminate
      check_for_active_task
      @task.stop(nil)
    end

    alias stop terminate

    # Set the host on which the various TaskScheduler methods will execute.
    # This method may require administrative privileges.
    #
    def machine=(host)
      raise TypeError unless host.is_a?(String)

      begin
        @service.Connect(host)
      rescue WIN32OLERuntimeError => err
        raise Error, ole_error("Connect", err)
      end

      @host = host
      host
    end

    # Similar to the TaskScheduler#machine= method, this method also allows
    # you to pass a user, domain and password as needed. This method may
    # require administrative privileges.
    #
    def set_machine(host, user = nil, domain = nil, password = nil)
      raise TypeError unless host.is_a?(String)

      begin
        @service.Connect(host, user, domain, password)
      rescue WIN32OLERuntimeError => err
        raise Error, ole_error("Connect", err)
      end

      @host = host
      host
    end

    alias host= machine=
    alias machine host
    alias set_host set_machine

    # Sets the +user+ and +password+ for the given task. If the user and
    # password are set properly then true is returned.
    # throws TypeError if password is not provided for other than system users
    def set_account_information(user_id, password, interactive)
      @interactive ||= interactive
      check_credential_requirements(user_id, password)
      check_for_active_task
      @password = password
      register_task_definition(@task.Definition, @task.Path, user_id, password)
      true
    end

    # Returns the user or group associated with the task or nil if no one has been associated with the task yet
    #
    # @return [String] user or group associated with the task
    #
    def account_information
      if @task.nil?
        nil
      elsif !@task.Definition.Principal.UserId.empty?
        @task.Definition.Principal.UserId
      else
        @task.Definition.Principal.GroupId
      end
    end

    # Returns the name of the application associated with the task. If
    # no application is associated with the task then nil is returned.
    #
    def application_name
      check_for_active_task

      app = nil

      @task.Definition.Actions.each do |action|
        if action.Type == 0 # TASK_ACTION_EXEC
          app = action.Path
          break
        end
      end

      app
    end

    # Sets the name of the application associated with the task.
    #
    def application_name=(app)
      raise TypeError unless app.is_a?(String)

      check_for_active_task

      definition = @task.Definition

      definition.Actions.each do |action|
        action.Path = app if action.Type == 0
      end

      register_task_definition(definition)

      app
    end

    # Returns the command line parameters for the task.
    #
    def parameters
      check_for_active_task

      param = nil

      @task.Definition.Actions.each do |action|
        param = action.Arguments if action.Type == 0
      end

      param
    end

    # Sets the parameters for the task. These parameters are passed as command
    # line arguments to the application the task will run. To clear the command
    # line parameters set it to an empty string.
    #--
    # NOTE: Again, it seems the task must be reactivated to be picked up.
    #
    def parameters=(param)
      raise TypeError unless param.is_a?(String)

      check_for_active_task

      definition = @task.Definition

      definition.Actions.each do |action|
        action.Arguments = param if action.Type == 0
      end

      register_task_definition(definition)

      param
    end

    # Returns the working directory for the task.
    #
    def working_directory
      check_for_active_task

      dir = nil

      @task.Definition.Actions.each do |action|
        dir = action.WorkingDirectory if action.Type == 0
      end

      dir
    end

    # Sets the working directory for the task.
    #--
    # TODO: Why do I have to reactivate the task to see the change?
    #
    def working_directory=(dir)
      raise TypeError unless dir.is_a?(String)

      check_for_active_task

      definition = @task.Definition

      definition.Actions.each do |action|
        action.WorkingDirectory = dir if action.Type == 0
      end

      register_task_definition(definition)

      dir
    end

    # Returns the task's priority level. Possible values are 'idle', 'lowest'.
    # 'below_normal_8', 'below_normal_7', 'normal_6', 'normal_5', 'normal_4',
    # 'above_normal_3', 'above_normal_2', 'highest', 'critical' and 'unknown'.
    #
    def priority
      check_for_active_task

      priority = case @task.Definition.Settings.Priority
                 when 0
                   "critical"
                 when 1
                   "highest"
                 when 2
                   "above_normal_2"
                 when 3
                   "above_normal_3"
                 when 4
                   "normal_4"
                 when 5
                   "normal_5"
                 when 6
                   "normal_6"
                 when 7
                   "below_normal_7"
                 when 8
                   "below_normal_8"
                 when 9
                   "lowest"
                 when 10
                   "idle"
                 else
                   "unknown"
                 end

      priority
    end

    # Sets the priority of the task. The +priority+ should be a numeric
    # priority constant value.
    #
    def priority=(priority)
      raise TypeError unless priority.is_a?(Numeric)

      check_for_active_task

      definition = @task.Definition
      definition.Settings.Priority = priority

      register_task_definition(definition)

      priority
    end

    # Creates a new work item (scheduled job) with the given +trigger+. The
    # trigger variable is a hash of options that define when the scheduled
    # job should run.
    #
    def new_work_item(task, trigger, userinfo = { user: nil, password: nil, interactive: false })
      raise TypeError unless userinfo.is_a?(Hash) && task.is_a?(String) && trigger.is_a?(Hash)

      # If user ID is not given, consider it as a 'SYSTEM' user
      userinfo[:user] = SERVICE_ACCOUNT_USERS.first if userinfo[:user].to_s.empty?
      @password = userinfo[:password]
      @interactive = userinfo[:interactive]

      check_credential_requirements(userinfo[:user], userinfo[:password])

      taskDefinition = @service.NewTask(0)
      taskDefinition.RegistrationInfo.Description = ""
      taskDefinition.RegistrationInfo.Author = ""
      taskDefinition.Settings.StartWhenAvailable = false
      taskDefinition.Settings.Enabled = true
      taskDefinition.Settings.Hidden = false

      unless trigger.empty?
        raise ArgumentError, "Unknown trigger type" unless valid_trigger_option(trigger[:trigger_type])

        validate_trigger(trigger)

        startTime = format("%04d-%02d-%02dT%02d:%02d:00", trigger[:start_year], trigger[:start_month], trigger[:start_day], trigger[:start_hour], trigger[:start_minute])

        # Set defaults
        trigger[:end_year]  ||= 0
        trigger[:end_month] ||= 0
        trigger[:end_day]   ||= 0

        endTime = format("%04d-%02d-%02dT00:00:00", trigger[:end_year], trigger[:end_month], trigger[:end_day])

        trig = taskDefinition.Triggers.Create(trigger[:trigger_type].to_i)
        trig.Id = "RegistrationTriggerId#{taskDefinition.Triggers.Count}"
        trig.StartBoundary = startTime if startTime != "0000-00-00T00:00:00"
        trig.EndBoundary = endTime if endTime != "0000-00-00T00:00:00"
        trig.Enabled = true

        repetitionPattern = trig.Repetition

        if trigger[:minutes_duration].to_i > 0
          repetitionPattern.Duration = "PT#{trigger[:minutes_duration] || 0}M"
        end

        if trigger[:minutes_interval].to_i > 0
          repetitionPattern.Interval = "PT#{trigger[:minutes_interval] || 0}M"
        end

        tmp = trigger[:type]
        tmp = nil unless tmp.is_a?(Hash)

        case trigger[:trigger_type]
        when TASK_TIME_TRIGGER_DAILY
          trig.DaysInterval = tmp[:days_interval] if tmp && tmp[:days_interval]
          if trigger[:random_minutes_interval].to_i > 0
            trig.RandomDelay = "PT#{trigger[:random_minutes_interval]}M"
          end
        when TASK_TIME_TRIGGER_WEEKLY
          trig.DaysOfWeek = tmp[:days_of_week] if tmp && tmp[:days_of_week]
          trig.WeeksInterval = tmp[:weeks_interval] if tmp && tmp[:weeks_interval]
          if trigger[:random_minutes_interval].to_i > 0
            trig.RandomDelay = "PT#{trigger[:random_minutes_interval] || 0}M"
          end
        when TASK_TIME_TRIGGER_MONTHLYDATE
          trig.MonthsOfYear = tmp[:months] if tmp && tmp[:months]
          trig.DaysOfMonth = tmp[:days] if tmp && tmp[:days]
          if trigger[:random_minutes_interval].to_i > 0
            trig.RandomDelay = "PT#{trigger[:random_minutes_interval] || 0}M"
          end
          trig.RunOnLastDayOfMonth = trigger[:run_on_last_day_of_month] if trigger[:run_on_last_day_of_month]
        when TASK_TIME_TRIGGER_MONTHLYDOW
          trig.MonthsOfYear = tmp[:months] if tmp && tmp[:months]
          trig.DaysOfWeek = tmp[:days_of_week] if tmp && tmp[:days_of_week]
          trig.WeeksOfMonth = tmp[:weeks_of_month] if tmp && tmp[:weeks_of_month]
          if trigger[:random_minutes_interval].to_i > 0
            trig.RandomDelay = "PT#{trigger[:random_minutes_interval] || 0}M"
          end
          trig.RunOnLastWeekOfMonth = trigger[:run_on_last_week_of_month] if trigger[:run_on_last_week_of_month]
        when TASK_TIME_TRIGGER_ONCE
          if trigger[:random_minutes_interval].to_i > 0
            trig.RandomDelay = "PT#{trigger[:random_minutes_interval] || 0}M"
          end
        when TASK_EVENT_TRIGGER_AT_SYSTEMSTART
          trig.Delay = "PT#{trigger[:delay_duration] || 0}M"
        when TASK_EVENT_TRIGGER_AT_LOGON
          trig.UserId = trigger[:user_id] if trigger[:user_id]
          trig.Delay = "PT#{trigger[:delay_duration] || 0}M"
        when TASK_TRIGGER_SESSION_STATE_CHANGE
          trig.UserId = trigger[:user_id] if trigger[:user_id]
          trig.Delay = "PT#{trigger[:delay_duration] || 0}M"
          trig.StateChange = trigger[:state_change]
        end
      end

      act = taskDefinition.Actions.Create(0)
      act.Path = "cmd"

      register_task_definition(taskDefinition, task, userinfo[:user], userinfo[:password])

      @task = @root.GetTask(task)
    end

    alias new_task new_work_item

    # Returns the number of triggers associated with the active task.
    #
    def trigger_count
      raise Error, "No currently active task" if @task.nil?

      @task.Definition.Triggers.Count
    end

    # Returns a string that describes the current trigger at the specified
    # index for the active task.
    #
    # Example: "At 7:14 AM every day, starting 4/11/2015"
    #
    def trigger_string(index)
      raise TypeError unless index.is_a?(Numeric)

      check_for_active_task
      index += 1 # first item index is 1

      begin
        trigger = @task.Definition.Triggers.Item(index)
      rescue WIN32OLERuntimeError
        raise Error, "No trigger found at index '#{index}'"
      end

      "Starting #{trigger.StartBoundary}"
    end

    # Deletes the trigger at the specified index.
    #--
    # TODO: Fix.
    #
    def delete_trigger(index)
      raise TypeError unless index.is_a?(Numeric)

      check_for_active_task
      index += 1 # first item index is 1

      definition = @task.Definition
      definition.Triggers.Remove(index)
      register_task_definition(definition)

      index
    end

    # Returns a hash that describes the trigger at the given index for the
    # current task.
    #
    def trigger(index)
      raise TypeError unless index.is_a?(Numeric)

      check_for_active_task
      index += 1 # first item index is 1

      begin
        trig = @task.Definition.Triggers.Item(index)
      rescue WIN32OLERuntimeError => err
        raise Error, ole_error("Item", err)
      end

      trigger = {}

      case trig.Type
      when TASK_TIME_TRIGGER_DAILY
        tmp = {}
        tmp[:days_interval] = trig.DaysInterval
        trigger[:type] = tmp
        trigger[:random_minutes_interval] = time_in_minutes(trig.RandomDelay)
      when TASK_TIME_TRIGGER_WEEKLY
        tmp = {}
        tmp[:weeks_interval] = trig.WeeksInterval
        tmp[:days_of_week] = trig.DaysOfWeek
        trigger[:type] = tmp
        trigger[:random_minutes_interval] = time_in_minutes(trig.RandomDelay)
      when TASK_TIME_TRIGGER_MONTHLYDATE
        tmp = {}
        tmp[:months] = trig.MonthsOfYear
        tmp[:days] = trig.DaysOfMonth
        trigger[:type] = tmp
        trigger[:random_minutes_interval] = time_in_minutes(trig.RandomDelay)
        trigger[:run_on_last_day_of_month] = trig.RunOnLastDayOfMonth
      when TASK_TIME_TRIGGER_MONTHLYDOW
        tmp = {}
        tmp[:months] = trig.MonthsOfYear
        tmp[:days_of_week] = trig.DaysOfWeek
        tmp[:weeks_of_month] = trig.WeeksOfMonth
        trigger[:type] = tmp
        trigger[:random_minutes_interval] = time_in_minutes(trig.RandomDelay)
        trigger[:run_on_last_week_of_month] = trig.RunOnLastWeekOfMonth
      when TASK_TIME_TRIGGER_ONCE
        tmp = {}
        tmp[:once] = nil
        trigger[:type] = tmp
        trigger[:random_minutes_interval] = time_in_minutes(trig.RandomDelay)
      when TASK_EVENT_TRIGGER_AT_SYSTEMSTART
        trigger[:delay_duration] = time_in_minutes(trig.Delay)
      when TASK_EVENT_TRIGGER_AT_LOGON
        trigger[:user_id] = trig.UserId if trig.UserId.to_s != ""
        trigger[:delay_duration] = time_in_minutes(trig.Delay)
      when TASK_EVENT_TRIGGER_ON_IDLE
        trigger[:execution_time_limit] = time_in_minutes(trig.ExecutionTimeLimit)
      when TASK_TRIGGER_SESSION_STATE_CHANGE
        trigger[:user_id] = trig.UserId if trig.UserId.to_s != ""
        trigger[:delay_duration] = time_in_minutes(trig.Delay)
        # trigger[:state_change] = trig.SessionState
      else
        raise Error, "Unknown trigger type"
      end

      trigger[:start_year], trigger[:start_month], trigger[:start_day],
      trigger[:start_hour], trigger[:start_minute] = trig.StartBoundary.scan(/(\d+)-(\d+)-(\d+)T(\d+):(\d+)/).first

      trigger[:end_year], trigger[:end_month],
      trigger[:end_day] = trig.EndBoundary.scan(/(\d+)-(\d+)-(\d+)T/).first

      trigger[:minutes_duration] = time_in_minutes(trig.Repetition.Duration)
      trigger[:minutes_interval] = time_in_minutes(trig.Repetition.Interval)
      trigger[:trigger_type] = trig.Type

      trigger
    end

    # Sets the trigger for the currently active task. The +trigger+ is a hash
    # with the following possible options:
    #
    # * days
    # * days_interval
    # * days_of_week
    # * end_day
    # * end_month
    # * end_year
    # * flags
    # * minutes_duration
    # * minutes_interval
    # * months
    # * random_minutes_interval
    # * start_day
    # * start_hour
    # * start_minute
    # * start_month
    # * start_year
    # * trigger_type
    # * type
    # * weeks
    # * weeks_interval
    #
    def trigger=(trigger)
      raise TypeError unless trigger.is_a?(Hash)
      raise ArgumentError, "Unknown trigger type" unless valid_trigger_option(trigger[:trigger_type])

      check_for_active_task

      validate_trigger(trigger)

      definition = @task.Definition
      definition.Triggers.Clear()

      startTime = format("%04d-%02d-%02dT%02d:%02d:00", trigger[:start_year], trigger[:start_month], trigger[:start_day], trigger[:start_hour], trigger[:start_minute])

      endTime = format("%04d-%02d-%02dT00:00:00", trigger[:end_year], trigger[:end_month], trigger[:end_day])

      trig = definition.Triggers.Create(trigger[:trigger_type].to_i)
      trig.Id = "RegistrationTriggerId#{definition.Triggers.Count}"
      trig.StartBoundary = startTime if startTime != "0000-00-00T00:00:00"
      trig.EndBoundary = endTime if endTime != "0000-00-00T00:00:00"
      trig.Enabled = true

      repetitionPattern = trig.Repetition

      if trigger[:minutes_duration].to_i > 0
        repetitionPattern.Duration = "PT#{trigger[:minutes_duration] || 0}M"
      end

      if trigger[:minutes_interval].to_i > 0
        repetitionPattern.Interval = "PT#{trigger[:minutes_interval] || 0}M"
      end

      tmp = trigger[:type]
      tmp = nil unless tmp.is_a?(Hash)

      case trigger[:trigger_type]
      when TASK_TIME_TRIGGER_DAILY
        trig.DaysInterval = tmp[:days_interval] if tmp && tmp[:days_interval]
        if trigger[:random_minutes_interval].to_i > 0
          trig.RandomDelay = "PT#{trigger[:random_minutes_interval]}M"
        end
      when TASK_TIME_TRIGGER_WEEKLY
        trig.DaysOfWeek = tmp[:days_of_week] if tmp && tmp[:days_of_week]
        trig.WeeksInterval = tmp[:weeks_interval] if tmp && tmp[:weeks_interval]
        if trigger[:random_minutes_interval].to_i > 0
          trig.RandomDelay = "PT#{trigger[:random_minutes_interval] || 0}M"
        end
      when TASK_TIME_TRIGGER_MONTHLYDATE
        trig.MonthsOfYear = tmp[:months] if tmp && tmp[:months]
        trig.DaysOfMonth = tmp[:days] if tmp && tmp[:days]
        if trigger[:random_minutes_interval].to_i > 0
          trig.RandomDelay = "PT#{trigger[:random_minutes_interval] || 0}M"
        end
        trig.RunOnLastDayOfMonth = trigger[:run_on_last_day_of_month] if trigger[:run_on_last_day_of_month]
      when TASK_TIME_TRIGGER_MONTHLYDOW
        trig.MonthsOfYear = tmp[:months] if tmp && tmp[:months]
        trig.DaysOfWeek = tmp[:days_of_week] if tmp && tmp[:days_of_week]
        trig.WeeksOfMonth = tmp[:weeks_of_month] if tmp && tmp[:weeks_of_month]
        if trigger[:random_minutes_interval].to_i > 0
          trig.RandomDelay = "PT#{trigger[:random_minutes_interval] || 0}M"
        end
        trig.RunOnLastWeekOfMonth = trigger[:run_on_last_week_of_month] if trigger[:run_on_last_week_of_month]
      when TASK_TIME_TRIGGER_ONCE
        if trigger[:random_minutes_interval].to_i > 0
          trig.RandomDelay = "PT#{trigger[:random_minutes_interval] || 0}M"
        end
      when TASK_EVENT_TRIGGER_AT_SYSTEMSTART
        trig.Delay = "PT#{trigger[:delay_duration] || 0}M"
      when TASK_EVENT_TRIGGER_AT_LOGON
        trig.UserId = trigger[:user_id] if trigger[:user_id]
        trig.Delay = "PT#{trigger[:delay_duration] || 0}M"
        trig.SessionState = trigger[:session_state]
      when TASK_EVENT_TRIGGER_ON_IDLE
        # for setting execution time limit Ref : https://msdn.microsoft.com/en-us/library/windows/desktop/aa380724(v=vs.85).aspx
        if trigger[:execution_time_limit].to_i > 0
          trig.ExecutionTimeLimit = "PT#{trigger[:execution_time_limit] || 0}M"
        end
      end

      register_task_definition(definition)

      trigger
    end

    # Adds a trigger at the specified index.
    #
    def add_trigger(index, trigger)
      raise TypeError unless index.is_a?(Numeric)
      raise TypeError unless trigger.is_a?(Hash)
      raise ArgumentError, "Unknown trigger type" unless valid_trigger_option(trigger[:trigger_type])

      check_for_active_task

      definition = @task.Definition

      startTime = format("%04d-%02d-%02dT%02d:%02d:00", trigger[:start_year], trigger[:start_month], trigger[:start_day], trigger[:start_hour], trigger[:start_minute])

      # Set defaults
      trigger[:end_year]  ||= 0
      trigger[:end_month] ||= 0
      trigger[:end_day]   ||= 0

      endTime = format("%04d-%02d-%02dT00:00:00", trigger[:end_year], trigger[:end_month], trigger[:end_day])

      trig = definition.Triggers.Create(trigger[:trigger_type].to_i)
      trig.Id = "RegistrationTriggerId#{definition.Triggers.Count}"
      trig.StartBoundary = startTime if startTime != "0000-00-00T00:00:00"
      trig.EndBoundary = endTime if endTime != "0000-00-00T00:00:00"
      trig.Enabled = true

      repetitionPattern = trig.Repetition

      if trigger[:minutes_duration].to_i > 0
        repetitionPattern.Duration = "PT#{trigger[:minutes_duration] || 0}M"
      end

      if trigger[:minutes_interval].to_i > 0
        repetitionPattern.Interval = "PT#{trigger[:minutes_interval] || 0}M"
      end

      tmp = trigger[:type]
      tmp = nil unless tmp.is_a?(Hash)

      case trigger[:trigger_type]
      when TASK_TIME_TRIGGER_DAILY
        trig.DaysInterval = tmp[:days_interval] if tmp && tmp[:days_interval]
        if trigger[:random_minutes_interval].to_i > 0
          trig.RandomDelay = "PT#{trigger[:random_minutes_interval]}M"
        end
      when TASK_TIME_TRIGGER_WEEKLY
        trig.DaysOfWeek = tmp[:days_of_week] if tmp && tmp[:days_of_week]
        trig.WeeksInterval = tmp[:weeks_interval] if tmp && tmp[:weeks_interval]
        if trigger[:random_minutes_interval].to_i > 0
          trig.RandomDelay = "PT#{trigger[:random_minutes_interval] || 0}M"
        end
      when TASK_TIME_TRIGGER_MONTHLYDATE
        trig.MonthsOfYear = tmp[:months] if tmp && tmp[:months]
        trig.DaysOfMonth = tmp[:days] if tmp && tmp[:days]
        if trigger[:random_minutes_interval].to_i > 0
          trig.RandomDelay = "PT#{trigger[:random_minutes_interval] || 0}M"
        end
        trig.RunOnLastDayOfMonth = trigger[:run_on_last_day_of_month] if trigger[:run_on_last_day_of_month]
      when TASK_TIME_TRIGGER_MONTHLYDOW
        trig.MonthsOfYear = tmp[:months] if tmp && tmp[:months]
        trig.DaysOfWeek = tmp[:days_of_week] if tmp && tmp[:days_of_week]
        trig.WeeksOfMonth = tmp[:weeks_of_month] if tmp && tmp[:weeks_of_month]
        if trigger[:random_minutes_interval].to_i > 0
          trig.RandomDelay = "PT#{trigger[:random_minutes_interval] || 0}M"
        end
        trig.RunOnLastWeekOfMonth = trigger[:run_on_last_week_of_month] if trigger[:run_on_last_week_of_month]
      when TASK_TIME_TRIGGER_ONCE
        if trigger[:random_minutes_interval].to_i > 0
          trig.RandomDelay = "PT#{trigger[:random_minutes_interval] || 0}M"
        end
      when TASK_EVENT_TRIGGER_AT_SYSTEMSTART
        trig.Delay = "PT#{trigger[:delay_duration] || 0}M"
      when TASK_EVENT_TRIGGER_AT_LOGON
        trig.UserId = trigger[:user_id] if trigger[:user_id]
        trig.Delay = "PT#{trigger[:delay_duration] || 0}M"
      when TASK_TRIGGER_SESSION_STATE_CHANGE
        trig.UserId = trigger[:user_id] if trigger[:user_id]
        trig.Delay = "PT#{trigger[:delay_duration] || 0}M"
        trig.SessionState = trigger[:session_state]
      end

      register_task_definition(definition)

      true
    end

    # Returns the status of the currently active task. Possible values are
    # 'ready', 'running', 'not scheduled' or 'unknown'.
    #
    def status
      check_for_active_task

      status = case @task.State
               when 3
                 "ready"
               when 4
                 "running"
               when 2
                 "queued"
               when 1
                 "not scheduled"
               else
                 "unknown"
               end

      status
    end

    # Returns true if current task is enabled
    def enabled?
      check_for_active_task
      @task.enabled
    end

    # Returns the exit code from the last scheduled run.
    #
    def exit_code
      check_for_active_task
      @task.LastTaskResult
    end

    # Returns the comment associated with the task, if any.
    #
    def comment
      check_for_active_task
      @task.Definition.RegistrationInfo.Description
    end

    alias description comment

    # Sets the comment for the task.
    #
    def comment=(comment)
      raise TypeError unless comment.is_a?(String)

      check_for_active_task

      definition = @task.Definition
      definition.RegistrationInfo.Description = comment
      register_task_definition(definition)

      comment
    end

    alias description= comment=

    # Returns the name of the user who created the task.
    #
    def creator
      check_for_active_task
      @task.Definition.RegistrationInfo.Author
    end

    alias author creator

    # Sets the creator for the task.
    #
    def creator=(creator)
      raise TypeError unless creator.is_a?(String)

      check_for_active_task

      definition = @task.Definition
      definition.RegistrationInfo.Author = creator
      register_task_definition(definition)

      creator
    end

    alias author= creator=

    # Returns a Time object that indicates the next time the task will run.
    #
    def next_run_time
      check_for_active_task
      @task.NextRunTime
    end

    # Returns a Time object indicating the most recent time the task ran or
    # nil if the task has never run.
    #
    def most_recent_run_time
      check_for_active_task

      time = nil

      begin
        time = Time.parse(@task.LastRunTime)
      rescue StandardError
        # Ignore
      end

      time
    end

    # Returns the execution time limit for current active task
    #
    def execution_time_limit
      check_for_active_task
      @task.Definition.Settings.ExecutionTimeLimit
    end

    # Returns the maximum length of time, in milliseconds, that the task
    # will run before terminating.
    #
    def max_run_time
      check_for_active_task

      t = @task.Definition.Settings.ExecutionTimeLimit
      year = t.scan(/(\d+?)Y/).flatten.first
      month = t.scan(/(\d+?)M/).flatten.first
      day = t.scan(/(\d+?)D/).flatten.first
      hour = t.scan(/(\d+?)H/).flatten.first
      min = t.scan(/T.*(\d+?)M/).flatten.first
      sec = t.scan(/(\d+?)S/).flatten.first

      time = 0
      time += year.to_i * 365 if year
      time += month.to_i * 30 if month
      time += day.to_i if day
      time *= 24
      time += hour.to_i if hour
      time *= 60
      time += min.to_i if min
      time *= 60
      time += sec.to_i if sec
      time *= 1000

      time
    end

    # Sets the maximum length of time, in milliseconds, that the task can run
    # before terminating. Returns the value you specified if successful.
    #
    def max_run_time=(max_run_time)
      raise TypeError unless max_run_time.is_a?(Numeric)

      check_for_active_task

      t = max_run_time
      t /= 1000
      limit = "PT#{t}S"

      definition = @task.Definition
      definition.Settings.ExecutionTimeLimit = limit
      register_task_definition(definition)

      max_run_time
    end

    # The Idle settings of a task
    #
    # @see https://docs.microsoft.com/en-us/windows/desktop/TaskSchd/idlesettings#properties
    #
    IdleSettings = %i{idle_duration restart_on_idle stop_on_idle_end wait_timeout}.freeze

    # Configures tasks settings
    #
    # @param [Hash] settings_hash The settings to configure a task
    # @option settings_hash [Boolean] :allow_demand_start The subject
    # @option settings_hash [Boolean] :allow_hard_terminate
    # @option settings_hash [Boolean] :disallow_start_if_on_batteries
    # @option settings_hash [Boolean] :disallow_start_on_remote_app_session
    # @option settings_hash [Boolean] :enabled
    # @option settings_hash [Boolean] :hidden
    # @option settings_hash [Boolean] :run_only_if_idle
    # @option settings_hash [Boolean] :run_only_if_network_available
    # @option settings_hash [Boolean] :start_when_available
    # @option settings_hash [Boolean] :stop_if_going_on_batteries
    # @option settings_hash [Boolean] :use_unified_scheduling_engine
    # @option settings_hash [Boolean] :volatile
    # @option settings_hash [Boolean] :wake_to_run
    # @option settings_hash [Boolean] :restart_on_idle The Idle Setting
    # @option settings_hash [Boolean] :stop_on_idle_end The Idle Setting
    # @option settings_hash [Integer] :compatibility
    # @option settings_hash [Integer] :multiple_instances
    # @option settings_hash [Integer] :priority
    # @option settings_hash [Integer] :restart_count
    # @option settings_hash [String] :delete_expired_task_after
    # @option settings_hash [String] :execution_time_limit
    # @option settings_hash [String] :restart_interval
    # @option settings_hash [String] :idle_duration The Idle Setting
    # @option settings_hash [String] :wait_timeout The Idle Setting
    #
    # @return [Hash] User input
    #
    # @see https://msdn.microsoft.com/en-us/library/windows/desktop/aa383480(v=vs.85).aspx#properties
    #
    def configure_settings(settings_hash)
      raise TypeError, "User input settings are required in hash" unless settings_hash.is_a?(Hash)

      check_for_active_task
      definition = @task.Definition

      # Check for invalid setting
      invalid_settings = settings_hash.keys - valid_settings_options
      raise TypeError, "Invalid setting passed: #{invalid_settings.join(", ")}" unless invalid_settings.empty?

      # Some modification is required in user input
      hash = settings_hash.dup

      # Conversion of few settings
      hash[:execution_time_limit] = hash[:max_run_time] unless hash[:max_run_time].nil?
      %i{execution_time_limit idle_duration restart_interval wait_timeout}.each do |setting|
        hash[setting] = "PT#{hash[setting]}M" unless hash[setting].nil?
      end

      task_settings = definition.Settings

      # Some Idle setting needs to be configured
      if IdleSettings.any? { |setting| hash.key?(setting) }
        idle_settings = task_settings.IdleSettings
        IdleSettings.each do |setting|
          next if hash[setting].nil?

          idle_settings.setproperty(camelize(setting.to_s), hash[setting])
          # This setting is not required to be configured now
          hash.delete(setting)
        end
      end

      # XML settings are not to be configured
      %i{xml_text xml}.map { |x| hash.delete(x) }

      hash.each do |setting, value|
        setting = camelize(setting.to_s)
        definition.Settings.setproperty(setting, value)
      end

      register_task_definition(definition)

      settings_hash
    end

    # Set registration information options. The possible options are:
    #
    # * author
    # * date
    # * description (or comment)
    # * documentation
    # * security_descriptor (should be a Win32::Security::SID)
    # * source
    # * uri
    # * version
    # * xml_text (or xml)
    #
    # Note that most of these options have standalone methods as well,
    # e.g. calling ts.configure_registration_info(:author => 'Dan') is
    # the same as calling ts.author = 'Dan'.
    #
    def configure_registration_info(hash)
      raise TypeError unless hash.is_a?(Hash)

      check_for_active_task

      definition = @task.Definition

      author = hash[:author]
      date = hash[:date]
      description = hash[:description] || hash[:comment]
      documentation = hash[:documentation]
      security_descriptor = hash[:security_descriptor]
      source = hash[:source]
      uri = hash[:uri]
      version = hash[:version]
      xml_text = hash[:xml_text] || hash[:xml]

      definition.RegistrationInfo.Author = author if author
      definition.RegistrationInfo.Date = date if date
      definition.RegistrationInfo.Description = description if description
      definition.RegistrationInfo.Documentation = documentation if documentation
      definition.RegistrationInfo.SecurityDescriptor = security_descriptor if security_descriptor
      definition.RegistrationInfo.Source = source if source
      definition.RegistrationInfo.URI = uri if uri
      definition.RegistrationInfo.Version = version if version
      definition.RegistrationInfo.XmlText = xml_text if xml_text

      register_task_definition(definition)

      hash
    end

    # Sets the principals for current active task. The principal is hash with following possible options.
    # Expected principal hash: { id: STRING, display_name: STRING, user_id: STRING,
    # logon_type: INTEGER, group_id: STRING, run_level: INTEGER }
    #
    def configure_principals(principals)
      raise TypeError unless principals.is_a?(Hash)

      check_for_active_task
      definition = @task.Definition
      definition.Principal.Id = principals[:id] if principals[:id].to_s != ""
      definition.Principal.DisplayName = principals[:display_name] if principals[:display_name].to_s != ""
      definition.Principal.UserId = principals[:user_id] if principals[:user_id].to_s != ""
      definition.Principal.LogonType = principals[:logon_type] if principals[:logon_type].to_s != ""
      definition.Principal.GroupId = principals[:group_id] if principals[:group_id].to_s != ""
      definition.Principal.RunLevel = principals[:run_level] if principals[:run_level].to_s != ""
      @interactive = true if principals[:logon_type] == TASK_LOGON_INTERACTIVE_TOKEN
      register_task_definition(definition)
      principals
    end

    # Returns a hash containing all the principal information of the current task
    def principals
      check_for_active_task
      principals_hash = {}
      @task.Definition.Principal.ole_get_methods.each do |principal|
        principals_hash[principal.name] = @task.Definition.Principal._getproperty(principal.dispid, [], [])
      end
      symbolize_keys(principals_hash)
    end

    # Returns a hash containing all settings of the current task
    def settings
      check_for_active_task
      settings_hash = {}
      @task.Definition.Settings.ole_get_methods.each do |setting|
        next if setting.name == "XmlText" # not needed

        settings_hash[setting.name] = @task.Definition.Settings._getproperty(setting.dispid, [], [])
      end

      settings_hash["IdleSettings"] = idle_settings
      settings_hash["NetworkSettings"] = network_settings
      symbolize_keys(settings_hash)
    end

    # Returns a hash of idle settings of the current task
    def idle_settings
      check_for_active_task
      settings_hash = {}
      @task.Definition.Settings.IdleSettings.ole_get_methods.each do |setting|
        settings_hash[setting.name] = @task.Definition.Settings.IdleSettings._getproperty(setting.dispid, [], [])
      end
      symbolize_keys(settings_hash)
    end

    # Returns a hash of network settings of the current task
    def network_settings
      check_for_active_task
      settings_hash = {}
      @task.Definition.Settings.NetworkSettings.ole_get_methods.each do |setting|
        settings_hash[setting.name] = @task.Definition.Settings.NetworkSettings._getproperty(setting.dispid, [], [])
      end
      symbolize_keys(settings_hash)
    end

    # Returns the user or group associated with the task. If no one is associated, it returns the default user i.e., 'SYSTEM'
    #
    # @param [WIN32OLE] definition
    #
    # @return [String] user_id
    #
    def task_user_id(definition)
      user_id = definition.Principal.UserId.to_s
      user_id = definition.Principal.GroupId.to_s if user_id.empty?
      user_id = SERVICE_ACCOUNT_USERS.first if user_id.empty?
      user_id
    end

    private

    # Returns a camle-case string to its underscore format
    def underscore(string)
      string.gsub(/([a-z\d])([A-Z])/, '\1_\2'.freeze).downcase
    end

    # Converts a snake-case string to camel-case format
    #
    # @param [String] str
    #
    # @return [String] In camel case format
    #
    def camelize(str)
      str.split("_").map(&:capitalize).join
    end

    # Converts all the keys of a hash to underscored-symbol format
    def symbolize_keys(hash)
      hash.each_with_object({}) do |(k, v), h|
        h[underscore(k.to_s).to_sym] = v.is_a?(Hash) ? symbolize_keys(v) : v
      end
    end

    def valid_trigger_option(trigger_type)
      [TASK_TIME_TRIGGER_ONCE, TASK_TIME_TRIGGER_DAILY, TASK_TIME_TRIGGER_WEEKLY,
       TASK_TIME_TRIGGER_MONTHLYDATE, TASK_TIME_TRIGGER_MONTHLYDOW, TASK_EVENT_TRIGGER_ON_IDLE,
       TASK_EVENT_TRIGGER_AT_SYSTEMSTART, TASK_EVENT_TRIGGER_AT_LOGON,
       TASK_TRIGGER_SESSION_STATE_CHANGE].include?(trigger_type.to_i)
    end

    def validate_trigger(hash)
      %i{start_year start_month start_day}.each do |key|
        raise ArgumentError, "#{key} must be set" unless hash[key]
      end
    end

    # Configurable settings options
    #
    # @note Logically, this is summation of
    #  * Settings
    #  * IdleSettings - [:idle_settings]
    #  * :max_run_time, :xml
    #
    # @return [Array]
    #
    def valid_settings_options
      %i{allow_demand_start allow_hard_terminate compatibility delete_expired_task_after
         disallow_start_if_on_batteries disallow_start_on_remote_app_session enabled
         execution_time_limit hidden idle_duration maintenance_settings max_run_time
         multiple_instances network_settings priority restart_count restart_interval
         restart_on_idle run_only_if_idle run_only_if_network_available
         start_when_available stop_if_going_on_batteries stop_on_idle_end
         use_unified_scheduling_engine volatile wait_timeout wake_to_run xml xml_text}
    end

    def check_for_active_task
      raise Error, "No currently active task" if @task.nil?
    end

    # Checks if the user belongs to service accounts category
    #
    # @return  [Boolean] True or False
    #
    def service_account_user?(user)
      SERVICE_ACCOUNT_USERS.include?(user.to_s.upcase)
    end

    # Checks if the user belongs to group accounts category
    #
    # @return [Boolean] True or False
    #
    def group_user?(user)
      BUILT_IN_GROUPS.include?(user.to_s.upcase)
    end

    # Checks if the user belongs to system users category
    #
    # @return  [Boolean] True or False
    #
    def system_user?(user)
      SYSTEM_USERS.include?(user.to_s.upcase)
    end

    # System users will not require a password
    # Other users will require a password if the task is non-interactive.
    #
    # @param [String] user_id
    # @param [String] password
    #
    def check_credential_requirements(user_id, password)
      user_id = user_id.to_s
      password = password.to_s

      if password.empty?
        unless system_user?(user_id) || @interactive
          raise ArgumentError, "Password is required for non-system users"
        end
      else
        if system_user?(user_id)
          raise ArgumentError, "Password is not required for system users"
        end
      end
    end

    # Returns the applicable flag as per the given users and groups which is used while
    # RegisterTaskDefinition
    #
    # @param [String] user_id
    # @param [String] password
    #
    # @return [Integer] Logon Types
    #
    def logon_type(user_id, password)
      if service_account_user?(user_id)
        TASK_LOGON_SERVICE_ACCOUNT
      elsif group_user?(user_id)
        TASK_LOGON_GROUP
      elsif !user_id.to_s.empty? && !password.to_s.empty?
        if @interactive
          TASK_LOGON_INTERACTIVE_TOKEN
        else
          TASK_LOGON_PASSWORD
        end
      else
        TASK_LOGON_INTERACTIVE_TOKEN
      end
    end

    # Uses RegisterTaskDefinition and creates or updates the task
    #
    # @param [WIN32OLE] definition
    # @param [String] path
    # @param [String] user_id
    # @param [String] password
    #
    # @return [Integer] Logon Types
    #
    def register_task_definition(definition, path = nil, user_id = nil, password = nil)
      user_id ||= task_user_id(definition)
      password ||= @password
      path ||= @task.Path
      @task = @root.RegisterTaskDefinition(
        path, # Path (name) of the task
        definition, # definition of the task
        TASK_CREATE_OR_UPDATE, # Equivalent to TASK_CREATE | TASK_UPDATE
        user_id,
        password,
        logon_type(user_id, password)
      )
    rescue WIN32OLERuntimeError => err
      method_name = caller_locations(1, 1)[0].label
      raise Error, ole_error(method_name, err)
    end
  end
  # :stopdoc:
end
