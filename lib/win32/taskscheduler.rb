require_relative 'windows/helper'
require 'win32ole'
require 'socket'
require 'time'
require 'structured_warnings'

# The Win32 module serves as a namespace only
module Win32

  # The TaskScheduler class encapsulates a Windows scheduled task
  class TaskScheduler
    include Windows::TaskSchedulerHelper

    # The version of the win32-taskscheduler library
    VERSION = '0.3.2'.freeze

    # The Error class is typically raised if any TaskScheduler methods fail.
    class Error < StandardError; end

    # Triggers

    # Trigger is set to run the task a single time
    TASK_TIME_TRIGGER_ONCE = 1

    # Trigger is set to run the task on a daily interval
    TASK_TIME_TRIGGER_DAILY = 2

    # Trigger is set to run the task on specific days of a specific week & month
    TASK_TIME_TRIGGER_WEEKLY = 3

    # Trigger is set to run the task on specific day(s) of the month
    TASK_TIME_TRIGGER_MONTHLYDATE = 4

    # Trigger is set to run the task on specific day(s) of the month
    TASK_TIME_TRIGGER_MONTHLYDOW = 5

    # Trigger is set to run the task if the system remains idle for the amount
    # of time specified by the idle wait time of the task
    TASK_EVENT_TRIGGER_ON_IDLE = 6

    TASK_TRIGGER_REGISTRATION = 7

    # Trigger is set to run the task at system startup
    TASK_EVENT_TRIGGER_AT_SYSTEMSTART = 8

    # Trigger is set to run the task when a user logs on
    TASK_EVENT_TRIGGER_AT_LOGON = 9

    TASK_TRIGGER_SESSION_STATE_CHANGE = 11

    # Daily Tasks

    # The task will run on Sunday
    TASK_SUNDAY = 0x1

    # The task will run on Monday
    TASK_MONDAY = 0x2

    # The task will run on Tuesday
    TASK_TUESDAY = 0x4

    # The task will run on Wednesday
    TASK_WEDNESDAY = 0x8

    # The task will run on Thursday
    TASK_THURSDAY = 0x10

    # The task will run on Friday
    TASK_FRIDAY = 0x20

    # The task will run on Saturday
    TASK_SATURDAY = 0x40

    # Weekly tasks

    # The task will run between the 1st and 7th day of the month
    TASK_FIRST_WEEK = 1

    # The task will run between the 8th and 14th day of the month
    TASK_SECOND_WEEK = 2

    # The task will run between the 15th and 21st day of the month
    TASK_THIRD_WEEK = 3

    # The task will run between the 22nd and 28th day of the month
    TASK_FOURTH_WEEK = 4

    # The task will run the last seven days of the month
    TASK_LAST_WEEK = 5

    # Monthly tasks

    # The task will run in January
    TASK_JANUARY = 0x1

    # The task will run in February
    TASK_FEBRUARY = 0x2

    # The task will run in March
    TASK_MARCH = 0x4

    # The task will run in April
    TASK_APRIL = 0x8

    # The task will run in May
    TASK_MAY = 0x10

    # The task will run in June
    TASK_JUNE = 0x20

    # The task will run in July
    TASK_JULY = 0x40

    # The task will run in August
    TASK_AUGUST = 0x80

    # The task will run in September
    TASK_SEPTEMBER = 0x100

    # The task will run in October
    TASK_OCTOBER = 0x200

    # The task will run in November
    TASK_NOVEMBER = 0x400

    # The task will run in December
    TASK_DECEMBER = 0x800

    # Flags

    # Used when converting AT service jobs into work items
    TASK_FLAG_INTERACTIVE = 0x1

    # The work item will be deleted when there are no more scheduled run times
    TASK_FLAG_DELETE_WHEN_DONE = 0x2

    # The work item is disabled. Useful for temporarily disabling a task
    TASK_FLAG_DISABLED = 0x4

    # The work item begins only if the computer is not in use at the scheduled
    # start time
    TASK_FLAG_START_ONLY_IF_IDLE = 0x10

    # The work item terminates if the computer makes an idle to non-idle
    # transition while the work item is running
    TASK_FLAG_KILL_ON_IDLE_END = 0x20

    # The work item does not start if the computer is running on battery power
    TASK_FLAG_DONT_START_IF_ON_BATTERIES = 0x40

    # The work item ends, and the associated application quits, if the computer
    # switches to battery power
    TASK_FLAG_KILL_IF_GOING_ON_BATTERIES = 0x80

    # The work item starts only if the computer is in a docking station
    TASK_FLAG_RUN_ONLY_IF_DOCKED = 0x100

    # The work item created will be hidden
    TASK_FLAG_HIDDEN = 0x200

    # The work item runs only if there is a valid internet connection
    TASK_FLAG_RUN_IF_CONNECTED_TO_INTERNET = 0x400

    # The work item starts again if the computer makes a non-idle to idle
    # transition
    TASK_FLAG_RESTART_ON_IDLE_RESUME = 0x800

    # The work item causes the system to be resumed, or awakened, if the
    # system is running on batter power
    TASK_FLAG_SYSTEM_REQUIRED = 0x1000

    # The work item runs only if a specified account is logged on interactively
    TASK_FLAG_RUN_ONLY_IF_LOGGED_ON = 0x2000

    # Triggers

    # The task will stop at some point in time
    TASK_TRIGGER_FLAG_HAS_END_DATE = 0x1

    # The task can be stopped at the end of the repetition period
    TASK_TRIGGER_FLAG_KILL_AT_DURATION_END = 0x2

    # The task trigger is disabled
    TASK_TRIGGER_FLAG_DISABLED = 0x4

    # :stopdoc:

    TASK_MAX_RUN_TIMES = 1440
    TASKS_TO_RETRIEVE  = 5

    # Task creation

    TASK_VALIDATE_ONLY = 0x1
    TASK_CREATE = 0x2
    TASK_UPDATE = 0x4
    TASK_CREATE_OR_UPDATE = 0x6
    TASK_DISABLE = 0x8
    TASK_DONT_ADD_PRINCIPAL_ACE = 0x10
    TASK_IGNORE_REGISTRATION_TRIGGERS = 0x20

    # Task logon types

    TASK_LOGON_NONE = 0
    TASK_LOGON_PASSWORD = 1
    TASK_LOGON_S4U = 2
    TASK_LOGON_INTERACTIVE_TOKEN = 3
    TASK_LOGON_GROUP = 4
    TASK_LOGON_SERVICE_ACCOUNT = 5
    TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD = 6

    # Priority classes

    REALTIME_PRIORITY_CLASS     = 0
    HIGH_PRIORITY_CLASS         = 1
    ABOVE_NORMAL_PRIORITY_CLASS = 2 # Or 3
    NORMAL_PRIORITY_CLASS       = 4 # Or 5, 6
    BELOW_NORMAL_PRIORITY_CLASS = 7 # Or 8
    IDLE_PRIORITY_CLASS         = 9 # Or 10

    CLSCTX_INPROC_SERVER  = 0x1
    CLSID_CTask =  [0x148BD520,0xA2AB,0x11CE,0xB1,0x1F,0x00,0xAA,0x00,0x53,0x05,0x03].pack('LSSC8')
    CLSID_CTaskScheduler =  [0x148BD52A,0xA2AB,0x11CE,0xB1,0x1F,0x00,0xAA,0x00,0x53,0x05,0x03].pack('LSSC8')
    IID_ITaskScheduler = [0x148BD527,0xA2AB,0x11CE,0xB1,0x1F,0x00,0xAA,0x00,0x53,0x05,0x03].pack('LSSC8')
    IID_ITask = [0x148BD524,0xA2AB,0x11CE,0xB1,0x1F,0x00,0xAA,0x00,0x53,0x05,0x03].pack('LSSC8')
    IID_IPersistFile = [0x0000010b,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46].pack('LSSC8')

    # Days of month

    TASK_FIRST = 0x01
    TASK_SECOND = 0x02
    TASK_THIRD = 0x04
    TASK_FOURTH = 0x08
    TASK_FIFTH = 0x10
    TASK_SIXTH = 0x20
    TASK_SEVENTH = 0x40
    TASK_EIGHTH = 0x80
    TASK_NINETH = 0x100
    TASK_TENTH = 0x200
    TASK_ELEVENTH = 0x400
    TASK_TWELFTH = 0x800
    TASK_THIRTEENTH = 0x1000
    TASK_FOURTEENTH = 0x2000
    TASK_FIFTEENTH = 0x4000
    TASK_SIXTEENTH = 0x8000
    TASK_SEVENTEENTH = 0x10000
    TASK_EIGHTEENTH = 0x20000
    TASK_NINETEENTH = 0x40000
    TASK_TWENTIETH = 0x80000
    TASK_TWENTY_FIRST = 0x100000
    TASK_TWENTY_SECOND = 0x200000
    TASK_TWENTY_THIRD = 0x400000
    TASK_TWENTY_FOURTH = 0x800000
    TASK_TWENTY_FIFTH = 0x1000000
    TASK_TWENTY_SIXTH = 0x2000000
    TASK_TWENTY_SEVENTH = 0x4000000
    TASK_TWENTY_EIGHTH = 0x8000000
    TASK_TWENTY_NINTH = 0x10000000
    TASK_THIRTYETH = 0x20000000
    TASK_THIRTY_FIRST = 0x40000000
    TASK_LAST = 0x80000000

    # :startdoc:

    attr_accessor :password
    attr_reader :host

    # Returns a new TaskScheduler object, attached to +folder+. If that
    # folder does not exist, but the +force+ option is set to true, then
    # it will be created. Otherwise an error will be raised. The default
    # is to use the root folder.
    #
    # If +task+ and +trigger+ are present, then a new task is generated
    # as well. This is effectively the same as .new + #new_work_item.
    #
    def initialize(task = nil, trigger = nil, folder = "\\", force = false)
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
        @service = WIN32OLE.new('Schedule.Service')
      rescue WIN32OLERuntimeError => err
        raise Error, err.inspect
      end

      @service.Connect

      if folder != "\\"
        begin
          @root = @service.GetFolder(folder)
        rescue WIN32OLERuntimeError => err
          if force
            @root.CreateFolder(folder)
            @root = @service.GetFolder(folder)
          else
            raise ArgumentError, "folder '#{folder}' not found"
          end
        end
      else
        @root = @service.GetFolder(folder)
      end

      if task && trigger
        new_work_item(task, trigger)
      end
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
    def exists?(task)
      enum.include?(task)
    end

    def get_task(task)
      raise TypeError unless task.is_a?(String)

      begin
        registeredTask = @root.GetTask(task)
        @task = registeredTask
      rescue WIN32OLERuntimeError => err
        raise Error, ole_error('activate', err)
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
        raise Error, ole_error('activate', err)
      end
    end

    # Delete the specified task name.
    #
    def delete(task)
      raise TypeError unless task.is_a?(String)

      begin
        @root.DeleteTask(task, 0)
      rescue WIN32OLERuntimeError => err
        raise Error, ole_error('DeleteTask', err)
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
    def save(file = nil)
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
        raise Error, ole_error('Connect', err)
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
        raise Error, ole_error('Connect', err)
      end

      @host = host
      host
    end

    alias host= machine=
    alias machine host
    alias set_host set_machine

    SYSTEM_USERS = ['NT AUTHORITY\SYSTEM', "SYSTEM", 'NT AUTHORITY\LOCALSERVICE', 'NT AUTHORITY\NETWORKSERVICE', 'BUILTIN\USERS', "USERS"].freeze

    # Sets the +user+ and +password+ for the given task. If the user and
    # password are set properly then true is returned.
    # throws TypeError if password is not provided for other than system users
    def set_account_information(user, password)
      raise TypeError unless user.is_a?(String)
      unless SYSTEM_USERS.include?(user.upcase)
        raise TypeError unless password.is_a?(String)
      end
      check_for_active_task

      @password = password

      begin
        @task = @root.RegisterTaskDefinition(
          @task.Path,
          @task.Definition,
          TASK_CREATE_OR_UPDATE,
          user,
          password,
          TASK_LOGON_PASSWORD
        )
      rescue WIN32OLERuntimeError => err
        raise Error, ole_error('RegisterTaskDefinition', err)
      end

      true
    end

    # Returns the user associated with the task or nil if no user has yet
    # been associated with the task.
    #
    def account_information
      @task.nil? ? nil : @task.Definition.Principal.UserId
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

      update_task_definition(definition)

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

      update_task_definition(definition)

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

      update_task_definition(definition)

      dir
    end

    # Returns the task's priority level. Possible values are 'idle',
    # 'normal', 'high', 'realtime', 'below_normal', 'above_normal',
    # and 'unknown'.
    #
    def priority
      check_for_active_task

      case @task.Definition.Settings.Priority
        when 0
          priority = 'critical'
        when 1
          priority = 'highest'
        when 2
          priority = 'above_normal'
        when 3
          priority = 'above_normal'
        when 4
          priority = 'normal'
        when 5
          priority = 'normal'
        when 6
          priority = 'normal'
        when 7
          priority = 'below_normal'
        when 8
          priority = 'below_normal'
        when 9
          priority = 'lowest'
        when 10
          priority = 'idle'
        else
          priority = 'unknown'
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

      update_task_definition(definition)

      priority
    end

    # Creates a new work item (scheduled job) with the given +trigger+. The
    # trigger variable is a hash of options that define when the scheduled
    # job should run.
    #
    def new_work_item(task, trigger)
      raise TypeError unless task.is_a?(String)
      raise TypeError unless trigger.is_a?(Hash)
      raise ArgumentError, 'Unknown trigger type' unless valid_trigger_option(trigger[:trigger_type])

      validate_trigger(trigger)

      taskDefinition = @service.NewTask(0)
      taskDefinition.RegistrationInfo.Description = ''
      taskDefinition.RegistrationInfo.Author = ''
      taskDefinition.Settings.StartWhenAvailable = true
      taskDefinition.Settings.Enabled  = true
      taskDefinition.Settings.Hidden = false
      

      startTime = "%04d-%02d-%02dT%02d:%02d:00" % [
        trigger[:start_year], trigger[:start_month], trigger[:start_day],
        trigger[:start_hour], trigger[:start_minute]
      ]

      # Set defaults
      trigger[:end_year]  ||= 0
      trigger[:end_month] ||= 0
      trigger[:end_day]   ||= 0

      endTime = "%04d-%02d-%02dT00:00:00" % [
        trigger[:end_year], trigger[:end_month], trigger[:end_day]
      ]

      trig = taskDefinition.Triggers.Create(trigger[:trigger_type].to_i)
      trig.Id = "RegistrationTriggerId#{taskDefinition.Triggers.Count}"
      trig.StartBoundary = startTime
      trig.EndBoundary = endTime if endTime != '0000-00-00T00:00:00'
      trig.Enabled = true

      repetitionPattern = trig.Repetition

      if trigger[:minutes_duration].to_i > 0
        repetitionPattern.Duration = "PT#{trigger[:minutes_duration]||0}M"
      end

      if trigger[:minutes_interval].to_i > 0
        repetitionPattern.Interval = "PT#{trigger[:minutes_interval]||0}M"
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
            trig.RandomDelay = "PT#{trigger[:random_minutes_interval]||0}M"
          end
        when TASK_TIME_TRIGGER_MONTHLYDATE
          trig.MonthsOfYear = tmp[:months] if tmp && tmp[:months]
          trig.DaysOfMonth = tmp[:days] if tmp && tmp[:days]
          if trigger[:random_minutes_interval].to_i > 0
            trig.RandomDelay = "PT#{trigger[:random_minutes_interval]||0}M"
          end
        when TASK_TIME_TRIGGER_MONTHLYDOW
          trig.MonthsOfYear = tmp[:months] if tmp && tmp[:months]
          trig.DaysOfWeek = tmp[:days_of_week] if tmp && tmp[:days_of_week]
          trig.WeeksOfMonth = tmp[:weeks_of_month] if tmp && tmp[:weeks_of_month]
          if trigger[:random_minutes_interval].to_i>0
            trig.RandomDelay = "PT#{trigger[:random_minutes_interval]||0}M"
          end
        when TASK_TIME_TRIGGER_ONCE
          if trigger[:random_minutes_interval].to_i > 0
            trig.RandomDelay = "PT#{trigger[:random_minutes_interval]||0}M"
          end
        when TASK_EVENT_TRIGGER_AT_SYSTEMSTART
          trig.Delay = "PT#{trigger[:delay_duration]||0}M"          
        when TASK_EVENT_TRIGGER_AT_LOGON
          trig.UserId = trigger[:user_id] if trigger[:user_id]
          trig.Delay = "PT#{trigger[:delay_duration]||0}M"
      end

      act = taskDefinition.Actions.Create(0)
      act.Path = 'cmd'

      begin
        @task = @root.RegisterTaskDefinition(
          task,
          taskDefinition,
          TASK_CREATE_OR_UPDATE,
          nil,
          nil,
          TASK_LOGON_INTERACTIVE_TOKEN
        )
      rescue WIN32OLERuntimeError => err
        raise Error, ole_error('RegisterTaskDefinition', err)
      end

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
      index += 1  # first item index is 1

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
      index += 1  # first item index is 1

      definition = @task.Definition
      definition.Triggers.Remove(index)
      update_task_definition(definition)

      index
    end

    # Returns a hash that describes the trigger at the given index for the
    # current task.
    #
    def trigger(index)
      raise TypeError unless index.is_a?(Numeric)
      check_for_active_task
      index += 1  # first item index is 1

      begin
        trig = @task.Definition.Triggers.Item(index)
      rescue WIN32OLERuntimeError => err
        raise Error, ole_error('Item', err)
      end

      trigger = {}

      case trig.Type
        when TASK_TIME_TRIGGER_DAILY
          tmp = {}
          tmp[:days_interval] = trig.DaysInterval
          trigger[:type] = tmp
          trigger[:random_minutes_interval] = trig.RandomDelay.scan(/(\d+)M/)[0][0].to_i if trig.RandomDelay != ""
        when TASK_TIME_TRIGGER_WEEKLY
          tmp = {}
          tmp[:weeks_interval] = trig.WeeksInterval
          tmp[:days_of_week] = trig.DaysOfWeek
          trigger[:type] = tmp
          trigger[:random_minutes_interval] = trig.RandomDelay.scan(/(\d+)M/)[0][0].to_i if trig.RandomDelay != ""
        when TASK_TIME_TRIGGER_MONTHLYDATE
          tmp = {}
          tmp[:months] = trig.MonthsOfYear
          tmp[:days] = trig.DaysOfMonth
          trigger[:type] = tmp
          trigger[:random_minutes_interval] = trig.RandomDelay.scan(/(\d+)M/)[0][0].to_i if trig.RandomDelay != ""
        when TASK_TIME_TRIGGER_MONTHLYDOW
          tmp = {}
          tmp[:months] = trig.MonthsOfYear
          tmp[:days_of_week] = trig.DaysOfWeek
          tmp[:weeks_of_month] = trig.WeeksOfMonth
          trigger[:type] = tmp
          trigger[:random_minutes_interval] = trig.RandomDelay.scan(/(\d+)M/)[0][0].to_i if trig.RandomDelay != ""
        when TASK_TIME_TRIGGER_ONCE
          tmp = {}
          tmp[:once] = nil
          trigger[:type] = tmp
          trigger[:random_minutes_interval] = trig.RandomDelay.scan(/(\d+)M/)[0][0].to_i if trig.RandomDelay != ""
        when TASK_EVENT_TRIGGER_AT_SYSTEMSTART
          trigger[:delay_duration] = trig.Delay.scan(/(\d+)M/)[0][0].to_i if trig.Delay != ""
        when TASK_EVENT_TRIGGER_AT_LOGON
          trigger[:user_id] = trig.UserId if trig.UserId.to_s != ""
          trigger[:delay_duration] = trig.Delay.scan(/(\d+)M/)[0][0].to_i if trig.Delay != ""
        else
          raise Error, 'Unknown trigger type'
      end
      
      trigger[:start_year], trigger[:start_month],
      trigger[:start_day],  trigger[:start_hour],
      trigger[:start_minute] = trig.StartBoundary.scan(/(\d+)-(\d+)-(\d+)T(\d+):(\d+)/).first

      trigger[:end_year], trigger[:end_month],
      trigger[:end_day] = trig.EndBoundary.scan(/(\d+)-(\d+)-(\d+)T/).first

      if trig.Repetition.Duration != ""
        trigger[:minutes_duration] = trig.Repetition.Duration.scan(/(\d+)M/)[0][0].to_i
      end

      if trig.Repetition.Interval != ""
        trigger[:minutes_interval] = trig.Repetition.Interval.scan(/(\d+)M/)[0][0].to_i
      end
      
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
      raise ArgumentError, 'Unknown trigger type' unless valid_trigger_option(trigger[:trigger_type])
      
      check_for_active_task

      validate_trigger(trigger)

      definition = @task.Definition
      definition.Triggers.Clear()

      startTime = "%04d-%02d-%02dT%02d:%02d:00" % [
        trigger[:start_year], trigger[:start_month],
        trigger[:start_day], trigger[:start_hour], trigger[:start_minute]
      ]

      endTime = "%04d-%02d-%02dT00:00:00" % [
        trigger[:end_year], trigger[:end_month], trigger[:end_day]
      ]

      trig = definition.Triggers.Create(type)
      trig.Id = "RegistrationTriggerId#{definition.Triggers.Count}"
      trig.StartBoundary = startTime
      trig.EndBoundary = endTime if endTime != '0000-00-00T00:00:00'
      trig.Enabled = true

      repetitionPattern = trig.Repetition

      if trigger[:minutes_duration].to_i > 0
        repetitionPattern.Duration = "PT#{trigger[:minutes_duration]||0}M"
      end

      if trigger[:minutes_interval].to_i > 0
        repetitionPattern.Interval  = "PT#{trigger[:minutes_interval]||0}M"
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
            trig.RandomDelay = "PT#{trigger[:random_minutes_interval]||0}M"
          end
        when TASK_TIME_TRIGGER_MONTHLYDATE
          trig.MonthsOfYear = tmp[:months] if tmp && tmp[:months]
          trig.DaysOfMonth = tmp[:days] if tmp && tmp[:days]
          if trigger[:random_minutes_interval].to_i > 0
            trig.RandomDelay = "PT#{trigger[:random_minutes_interval]||0}M"
          end
        when TASK_TIME_TRIGGER_MONTHLYDOW
          trig.MonthsOfYear = tmp[:months] if tmp && tmp[:months]
          trig.DaysOfWeek = tmp[:days_of_week] if tmp && tmp[:days_of_week]
          trig.WeeksOfMonth = tmp[:weeks] if tmp && tmp[:weeks]
          if trigger[:random_minutes_interval].to_i > 0
            trig.RandomDelay = "PT#{trigger[:random_minutes_interval]||0}M"
          end
        when TASK_TIME_TRIGGER_ONCE
          if trigger[:random_minutes_interval].to_i > 0
            trig.RandomDelay = "PT#{trigger[:random_minutes_interval]||0}M"
          end
        when TASK_EVENT_TRIGGER_AT_SYSTEMSTART
          trig.Delay = "PT#{trigger[:delay_duration]||0}M"
        when TASK_EVENT_TRIGGER_AT_LOGON
          trig.UserId = trigger[:user_id] if trigger[:user_id]
          trig.Delay = "PT#{trigger[:delay_duration]||0}M"
      end

      update_task_definition(definition)

      trigger
    end

    # Adds a trigger at the specified index.
    #
    def add_trigger(index, trigger)
      raise TypeError unless index.is_a?(Numeric)
      raise TypeError unless trigger.is_a?(Hash)
      raise ArgumentError, 'Unknown trigger type' unless valid_trigger_option(trigger[:trigger_type])

      check_for_active_task

      definition = @task.Definition

      startTime = "%04d-%02d-%02dT%02d:%02d:00" % [
        trigger[:start_year], trigger[:start_month], trigger[:start_day],
        trigger[:start_hour], trigger[:start_minute]
      ]

      # Set defaults
      trigger[:end_year]  ||= 0
      trigger[:end_month] ||= 0
      trigger[:end_day]   ||= 0

      endTime = "%04d-%02d-%02dT00:00:00" % [
        trigger[:end_year], trigger[:end_month], trigger[:end_day]
      ]

      trig = definition.Triggers.Create( trigger[:trigger_type].to_i )
      trig.Id = "RegistrationTriggerId#{definition.Triggers.Count}"
      trig.StartBoundary = startTime
      trig.EndBoundary = endTime if endTime != '0000-00-00T00:00:00'
      trig.Enabled = true

      repetitionPattern = trig.Repetition

      if trigger[:minutes_duration].to_i > 0
        repetitionPattern.Duration = "PT#{trigger[:minutes_duration]||0}M"
      end

      if trigger[:minutes_interval].to_i > 0
        repetitionPattern.Interval  = "PT#{trigger[:minutes_interval]||0}M"
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
          trig.RandomDelay = "PT#{trigger[:random_minutes_interval]||0}M"
          end
        when TASK_TIME_TRIGGER_MONTHLYDATE
          trig.MonthsOfYear = tmp[:months] if tmp && tmp[:months]
          trig.DaysOfMonth = tmp[:days] if tmp && tmp[:days]
          if trigger[:random_minutes_interval].to_i > 0
          trig.RandomDelay = "PT#{trigger[:random_minutes_interval]||0}M"
          end
        when TASK_TIME_TRIGGER_MONTHLYDOW
          trig.MonthsOfYear  = tmp[:months] if tmp && tmp[:months]
          trig.DaysOfWeek  = tmp[:days_of_week] if tmp && tmp[:days_of_week]
          trig.WeeksOfMonth  = tmp[:weeks] if tmp && tmp[:weeks]
          if trigger[:random_minutes_interval].to_i > 0
          trig.RandomDelay = "PT#{trigger[:random_minutes_interval]||0}M"
          end
        when TASK_TIME_TRIGGER_ONCE
          if trigger[:random_minutes_interval].to_i > 0
          trig.RandomDelay = "PT#{trigger[:random_minutes_interval]||0}M"
          end
        when TASK_EVENT_TRIGGER_AT_SYSTEMSTART
          trig.Delay = "PT#{trigger[:delay_duration]||0}M"
        when TASK_EVENT_TRIGGER_AT_LOGON
          trig.UserId = trigger[:user_id] if trigger[:user_id]
          trig.Delay = "PT#{trigger[:delay_duration]||0}M"
      end

      update_task_definition(definition)

      true
    end

    # Returns the status of the currently active task. Possible values are
    # 'ready', 'running', 'not scheduled' or 'unknown'.
    #
    def status
      check_for_active_task

      case @task.State
        when 3
          status = 'ready'
        when 4
          status = 'running'
        when 1
          status = 'not scheduled'
        else
          status = 'unknown'
      end

      status
    end

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
      update_task_definition(definition)

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
      update_task_definition(definition)

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
      rescue
        # Ignore
      end

      time
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
      limit ="PT#{t}S"

      definition = @task.Definition
      definition.Settings.ExecutionTimeLimit = limit
      update_task_definition(definition)

      max_run_time
    end

    # Accepts a hash that lets you configure various task definition settings.
    # The possible options are:
    #
    # * allow_demand_start
    # * allow_hard_terminate
    # * compatibility
    # * delete_expired_task_after
    # * disallowed_start_if_on_batteries
    # * enabled
    # * execution_time_limit (or max_run_time)
    # * hidden
    # * idle_settings
    # * network_settings
    # * priority
    # * restart_count
    # * restart_interval
    # * run_only_if_idle
    # * run_only_if_network_available
    # * start_when_available
    # * stop_if_going_on_batteries
    # * wake_to_run
    # * xml_text (or xml)
    #
    def configure_settings(hash)
      raise TypeError unless hash.is_a?(Hash)
      check_for_active_task

      definition = @task.Definition

      allow_demand_start = hash[:allow_demand_start]
      allow_hard_terminate = hash[:allow_hard_terminate]
      compatibility = hash[:compatibility]
      delete_expired_task_after = hash[:delete_expired_task_after]
      disallow_start_if_on_batteries = hash[:disallow_start_if_on_batteries]
      enabled = hash[:enabled]
      execution_time_limit = hash[:execution_time_limit] || hash[:max_run_time]
      hidden = hash[:hidden]
      idle_settings = hash[:idle_settings]
      network_settings = hash[:network_settings]
      priority = hash[:priority]
      restart_count = hash[:restart_count]
      restart_interval = hash[:restart_interval]
      run_only_if_idle = hash[:run_only_if_idle]
      run_only_if_network_available = hash[:run_only_if_network_available]
      start_when_available = hash[:start_when_available]
      stop_if_going_on_batteries = hash[:stop_if_going_on_batteries]
      wake_to_run = hash[:wake_to_run]
      xml_text = hash[:xml_text] || hash[:xml]

      definition.Settings.AllowDemandStart = allow_demand_start if allow_demand_start
      definition.Settings.AllowHardTerminate = allow_hard_terminate if allow_hard_terminate
      definition.Settings.Compatibility = compatibility if compatibility
      definition.Settings.DeleteExpiredTaskAfter = delete_expired_task_after if delete_expired_task_after
      definition.Settings.DisallowStartIfOnBatteries = disallow_start_if_on_batteries if disallow_start_if_on_batteries
      definition.Settings.Enabled = enabled if enabled
      definition.Settings.ExecutionTimeLimit = execution_time_limit if execution_time_limit
      definition.Settings.Hidden = hidden if hidden
      definition.Settings.IdleSettings = idle_settings if idle_settings
      definition.Settings.NetworkSettings = network_settings if network_settings
      definition.Settings.Priority = priority if priority
      definition.Settings.RestartCount = restart_count if restart_count
      definition.Settings.RestartInterval = restart_interval if restart_interval
      definition.Settings.RunOnlyIfIdle = run_only_if_idle if run_only_if_idle
      definition.Settings.RunOnlyIfNetworkAvailable = run_only_if_network_available if run_only_if_network_available
      definition.Settings.StartWhenAvailable = start_when_available if start_when_available
      definition.Settings.StopIfGoingOnBatteries = stop_if_going_on_batteries if stop_if_going_on_batteries
      definition.Settings.WakeToRun = wake_to_run if wake_to_run
      definition.Settings.XmlText = xml_text if xml_text

      update_task_definition(definition)

      hash
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

      update_task_definition(definition)

      hash
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

    private

    # Returns a camle-case string to its underscore format
    def underscore(string)
      string.gsub(/([a-z\d])([A-Z])/, '\1_\2'.freeze).downcase
    end

    # Converts all the keys of a hash to underscored-symbol format
    def symbolize_keys(hash)
      hash.each_with_object({}) do |(k, v), h|
        h[underscore(k.to_s).to_sym] = v.is_a?(Hash) ? symbolize_keys(v) : v 
      end
    end    

    def valid_trigger_option(trigger_type)
      [ TASK_TIME_TRIGGER_ONCE, TASK_TIME_TRIGGER_DAILY, TASK_TIME_TRIGGER_WEEKLY,
        TASK_TIME_TRIGGER_MONTHLYDATE, TASK_TIME_TRIGGER_MONTHLYDOW, TASK_EVENT_TRIGGER_ON_IDLE,
        TASK_EVENT_TRIGGER_AT_SYSTEMSTART, TASK_EVENT_TRIGGER_AT_LOGON ].include?(trigger_type.to_i)
    end


    def validate_trigger(hash)
      [:start_year, :start_month, :start_day].each{ |key|
        raise ArgumentError, "#{key} must be set" unless hash[key]
      }
    end

    def check_for_active_task
      raise Error, 'No currently active task' if @task.nil?
    end

    def update_task_definition(definition)
      user = definition.Principal.UserId

      @task = @root.RegisterTaskDefinition(
        @task.Path,
        definition,
        TASK_CREATE_OR_UPDATE,
        user,
        @password,
        @password ? TASK_LOGON_PASSWORD : TASK_LOGON_INTERACTIVE_TOKEN
      )
    rescue WIN32OLERuntimeError => err
      method_name = caller_locations(1,1)[0].label
      raise Error, ole_error(method_name, err)
    end
  end
end
