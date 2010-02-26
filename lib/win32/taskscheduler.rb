require 'windows/process'
require 'win32ole'
include Windows::Process

# The Win32 module serves as a namespace only
module Win32

  # The TaskScheduler class encapsulates a Windows scheduled task
  class TaskScheduler

    # The version of the win32-taskscheduler library
    VERSION = '0.3.0'

    # The Error class is typically raised if any TaskScheduler methods fail.
    class Error < StandardError; end

    # Triggers

    # Trigger is set to run the task a single tim
    TASK_TIME_TRIGGER_ONCE = 0

    # Trigger is set to run the task on a daily interval
    TASK_TIME_TRIGGER_DAILY = 1

    # Trigger is set to run the task on specific days of a specific week & month
    TASK_TIME_TRIGGER_WEEKLY = 2

    # Trigger is set to run the task on specific day(s) of the month
    TASK_TIME_TRIGGER_MONTHLYDATE = 3

    # Trigger is set to run the task on specific day(s) of the month
    TASK_TIME_TRIGGER_MONTHLYDOW = 4

    # Trigger is set to run the task if the system remains idle for the amount
    # of time specified by the idle wait time of the task
    TASK_EVENT_TRIGGER_ON_IDLE = 5

    # Trigger is set to run the task at system startup
    TASK_EVENT_TRIGGER_AT_SYSTEMSTART = 6

    # Trigger is set to run the task when a user logs on
    TASK_EVENT_TRIGGER_AT_LOGON = 7

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

    CLSCTX_INPROC_SERVER  = 0x1
    CLSID_CTask =  [0x148BD520,0xA2AB,0x11CE,0xB1,0x1F,0x00,0xAA,0x00,0x53,0x05,0x03].pack('LSSC8')
    CLSID_CTaskScheduler =  [0x148BD52A,0xA2AB,0x11CE,0xB1,0x1F,0x00,0xAA,0x00,0x53,0x05,0x03].pack('LSSC8')
    IID_ITaskScheduler = [0x148BD527,0xA2AB,0x11CE,0xB1,0x1F,0x00,0xAA,0x00,0x53,0x05,0x03].pack('LSSC8')
    IID_ITask = [0x148BD524,0xA2AB,0x11CE,0xB1,0x1F,0x00,0xAA,0x00,0x53,0x05,0x03].pack('LSSC8')
    IID_IPersistFile = [0x0000010b,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46].pack('LSSC8')

    # :startdoc:

    attr_accessor :password

    # Returns a new TaskScheduler object, attached to +folder+. If that
    # folder does not exist, but the +force+ option is set to true, then
    # it will be created. Otherwise an error will be raised.
    #
    def initialize(folder = "\\", force = false)
      @task   = nil
      @folder = folder
      @force  = force

      raise ArgumentError, "invalid folder" unless folder.include?("\\")

      unless [TrueClass, FalseClass].include?(force.class)
        raise TypeError, "invalid force value" 
      end

      begin
        @service = WIN32OLE.new('Schedule.Service')
      rescue WIN32OLERuntimeError => err
        raise Error, err.inspect
      end

      unless @service.FolderExists(folder)
        if force
          @service.CreateFolder(folder)
        else
          raise ArgumentError, "folder '#{folder}' could not be found"
        end
      end

      @service.Connect
      @root = @service.GetFolder(folder)
      @password = nil
    end

    # Returns an array of scheduled task names.
    #
    def enum
      # Get the task folder that contains the tasks.
      taskCollection = @root.GetTasks(0)
      array = []
      taskCollection.each do |registeredTask|
        array.push(registeredTask.Name)
      end
      array
    end

    # Activate the specified task.
    #
    def activate(task)
      raise TypeError unless task.is_a?(String)

      begin
        registeredTask = @root.GetTask(task)
        registeredTask.Enabled = 1
      rescue WIN32OLERuntimeError
        raise Error, "Access Denied"
      end
      @task = registeredTask
      self
    end

    # Delete the specified task name.
    #
    def delete(task)
      raise TypeError unless task.is_a?(String)

      begin
        @root.DeleteTask(task, 0)
      rescue
        raise Error, "Access Denied"
      end

      self
    end

    # Execute the current task.
    #
    def run
      raise Error, 'null task' if @task.nil?

      task = @task.run(nil)
      self
    end

    # Saves the current task. Tasks must be saved before they can be activated.
    # The .job file itself is typically stored in the C:\WINDOWS\Tasks folder.
    #
    # If +file+ (an absolute path) is specified then the job is saved to that
    # file instead. A '.job' extension is recommended but not enforced.
    #
    # Note that calling TaskScheduler#save also resets the TaskScheduler object
    # so that there is no currently active task.
    #
    def save(file = nil)
      raise Error, 'null task' if @task.nil?
      # do nothing
      self
    end

    # Terminate the current task.
    #
    def terminate
      raise Error, 'null task' if @task.nil?

      task = @task.stop(nil)
      self
    end

    # Set the host on which the various TaskScheduler methods will execute.
    #
    def machine=(host)
      raise TypeError unless host.is_a?(String)
      @service.Connect(host)
      host
    end

    alias :host= :machine=

    # Sets the +user+ and +password+ for the given task. If the user and
    # password are set properly then true is returned.
    #
    # In some cases the job may be created, but the account information was
    # bad. In this case the task is created but a warning is generated and
    # false is returned.
    #
    def set_account_information(user, password)
      raise Error, 'No currently active task' if @task.nil?
      @password = password

      @root.RegisterTaskDefinition(
        @task.Path,
        @task.Definition,
        TASK_CREATE_OR_UPDATE,
        user,
        password,
        TASK_LOGON_PASSWORD
      )

      true
    end

    # Returns the user associated with the task or nil if no user has yet
    # been associated with the task.
    #
    def account_information
      raise Error, 'No currently active task' if @task.nil?
      user = @task.Definition.Principal.UserId
    end

    # Returns the name of the application associated with the task.
    #
    def application_name
      raise Error, 'No currently active task' if @task.nil?

      app = nil

      @task.Definition.Actions.each do |action|
        app = action.Path if action.Type == 0
      end

      app
    end

    # Returns the command line parameters for the task.
    #
    def parameters
      raise Error, 'No currently active task' if @task.nil?

      param = nil

      @task.Definition.Actions.each do |action|
        param = action.Arguments if action.Type == 0
      end

      param
    end

    # Sets the parameters for the task. These parameters are passed as command
    # line arguments to the application the task will run. To clear the command
    # line parameters set it to an empty string.
    #
    def parameters=(param)
      raise Error, 'No currently active task' if @task.nil?
      raise TypeError unless param.is_a?(String)

      definition = @task.Definition
      definition.Actions.each do |action|
        action.Arguments = param if action.Type == 0
      end
      user = definition.Principal.UserId

      @root.RegisterTaskDefinition(
        @task.Path,
        definition,
        TASK_CREATE_OR_UPDATE,
        user,
        @password,
        @password ? TASK_LOGON_PASSWORD : TASK_LOGON_INTERACTIVE_TOKEN
      )

      param
    end

    # Returns the working directory for the task.
    #
    def working_directory
      raise Error,"No currently active task" if @task.nil?

      dir = nil

      @task.Definition.Actions.each do |action|
        dir = action.WorkingDirectory if action.Type == 0
      end

      dir
    end

    # Sets the working directory for the task.
    #
    def working_directory=(dir)
      raise Error, 'No currently active task' if @task.nil?
      raise TypeError unless dir.is_a?(String)

      definition = @task.Definition

      definition.Actions.each do |action|
        action.WorkingDirectory = dir if action.Type == 0
      end

      user = definition.Principal.UserId

      @root.RegisterTaskDefinition(
        @task.Path,
        definition,
        0x6,
        user,
        @password,
        @password ? TASK_LOGON_PASSWORD : TASK_LOGON_INTERACTIVE_TOKEN
      )

      dir
    end

    # Returns the task's priority level. Possible values are 'idle',
    # 'normal', 'high', 'realtime', 'below_normal', 'above_normal',
    # and 'unknown'.
    #
    def priority
      raise Error, 'No currently active task' if @task.nil?

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
      raise Error, 'No currently active task' if @task.nil?
      raise TypeError unless priority.is_a?(Numeric)

      definition = @task.Definition
      definition.Settings.Priority = priority
      user = definition.Principal.UserId

      @root.RegisterTaskDefinition(
        @task.Path,
        definition,
        TASK_CREATE_OR_UPDATE,
        user,
        @password,
        @password ? TASK_LOGON_PASSWORD : TASK_LOGON_INTERACTIVE_TOKEN
      )

      priority
    end

    # Creates a new work item (scheduled job) with the given +trigger+. The
    # trigger variable is a hash of options that define when the scheduled
    # job should run.
    #
    def new_work_item(task, trigger)
      raise TypeError unless trigger.is_a?(Hash)

      taskDefinition = @service.NewTask(0)
      taskDefinition.RegistrationInfo.Description = ''
      taskDefinition.RegistrationInfo.Author  = ''
      taskDefinition.Settings.StartWhenAvailable = true
      taskDefinition.Settings.Enabled  = true
      taskDefinition.Settings.Hidden = false

      case trigger['trigger_type']
        when TASK_TIME_TRIGGER_DAILY
          type = 2
        when TASK_TIME_TRIGGER_WEEKLY
          type = 3
        when TASK_TIME_TRIGGER_MONTHLYDATE
          type = 4
        when TASK_TIME_TRIGGER_MONTHLYDOW
          type = 5
        when TASK_TIME_TRIGGER_ONCE
          type = 1
        else
          raise Error, 'Unknown trigger type'
      end

      startTime = "%04d-%02d-%02dT%02d:%02d:00" % [trigger['start_year'],trigger['start_month'],trigger['start_day'],trigger['start_hour'], trigger['start_minute']]
      endTime = "%04d-%02d-%02dT00:00:00" % [trigger['end_year'],trigger['end_month'],trigger['end_day']]

      trig = taskDefinition.Triggers.Create(type)
      trig.Id = "RegistrationTriggerId"
      trig.StartBoundary = startTime
      trig.EndBoundary = endTime if endTime != '0000-00-00T00:00:00'
      trig.Enabled = true

      repetitionPattern = trig.Repetition
      repetitionPattern.Duration = "PT#{trigger['minutes_duration']||0}M" if trigger['minutes_duration'].to_i>0
      repetitionPattern.Interval  = "PT#{trigger['minutes_interval']||0}M" if trigger['minutes_interval'].to_i>0

      tmp = trigger['type']
      tmp = nil unless tmp.is_a?(Hash)

      case trigger['trigger_type']
        when TASK_TIME_TRIGGER_DAILY
          trig.DaysInterval =tmp['days_interval'] if tmp && tmp['days_interval']
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']}M" if trigger['random_minutes_interval'].to_i>0
        when TASK_TIME_TRIGGER_WEEKLY
          trig.DaysOfWeek  = tmp['days_of_week'] if tmp && tmp['days_of_week']
          trig.WeeksInterval  = tmp['weeks_interval'] if tmp && tmp['weeks_interval']
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']||0}M" if trigger['random_minutes_interval'].to_i>0
        when TASK_TIME_TRIGGER_MONTHLYDATE
          trig.MonthsOfYear  = tmp['months'] if tmp && tmp['months']
          trig.DaysOfMonth  = tmp['days'] if tmp && tmp['days']
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']||0}M" if trigger['random_minutes_interval'].to_i>0
        when TASK_TIME_TRIGGER_MONTHLYDOW
          trig.MonthsOfYear  = tmp['months'] if tmp && tmp['months']
          trig.DaysOfWeek  = tmp['days_of_week'] if tmp && tmp['days_of_week']
          trig.WeeksOfMonth  = tmp['weeks'] if tmp && tmp['weeks']
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']||0}M" if trigger['random_minutes_interval'].to_i>0
        when TASK_TIME_TRIGGER_ONCE
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']||0}M" if trigger['random_minutes_interval'].to_i>0
      end

      act = taskDefinition.Actions.Create(0)
      act.Path = 'cmd'

      @root.RegisterTaskDefinition(
        task,
        taskDefinition,
        TASK_CREATE_OR_UPDATE,
        nil,
        nil,
        TASK_LOGON_INTERACTIVE_TOKEN
      )
      @task = @root.GetTask(task)
      self
    end

    alias :new_task :new_work_item

    # Returns the number of triggers associated with the active task.
    #
    def trigger_count
      raise Error, "No currently active task" if @task.nil?

      count = @task.Definition.Triggers.Count
    end

    # Returns a string that describes the current trigger at the specified
    # index for the active task.
    #
    # Example: "At 7:14 AM every day, starting 4/11/2009"
    #
    def trigger_string(index)
      raise Error, 'No currently active task' if @task.nil?
      raise TypeError unless index.is_a?(Numeric)

      begin
        trigger = @task.Definition.Triggers.Item(index)
      rescue WIN32OLERuntimeError
        raise Error, "No trigger found at index '#{index}'"
      end

      "Starting #{trigger.StartBoundary}"
    end

    # Deletes the trigger at the specified index.
    #
    def delete_trigger(index)
      raise Error, 'No currently active task' if @task.nil?

      definition = @task.Definition
      definition.Triggers.Remove(index)
      user = definition.Principal.UserId

      @root.RegisterTaskDefinition(
        @task.Path,
        definition,
        TASK_CREATE_OR_UPDATE,
        user,
        @password,
        @password ? TASK_LOGON_PASSWORD : TASK_LOGON_INTERACTIVE_TOKEN
      )

      index
    end

    # Returns a hash that describes the trigger at the given index for the
    # current task.
    #
    def trigger(index)
      raise Error, 'No currently active task' if @task.nil?

      trig = @task.Definition.Triggers.Item(index)
      trigger = {}
      trigger['start_year'],trigger['start_month'],
      trigger['start_day'],trigger['start_hour'],
      trigger['start_minute'] = trig.StartBoundary.scan(/(\d+)-(\d+)-(\d+)T(\d+):(\d+)/).first

      trigger['end_year'],trigger['end_month'],
      trigger['end_day'] = trig.StartBoundary.scan(/(\d+)-(\d+)-(\d+)T/).first

      trigger['minutes_duration'] = trig.Repetition.Duration.scan(/(\d+)M/)[0][0].to_i if trig.Repetition.Duration != ""
      trigger['minutes_interval'] = trig.Repetition.Interval.scan(/(\d+)M/)[0][0].to_i if trig.Repetition.Interval != ""
      trigger['random_minutes_interval'] = trig.RandomDelay.scan(/(\d+)M/)[0][0].to_i if trig.RandomDelay != ""

      case trig.Type
        when 2
          trigger['trigger_type'] = TASK_TIME_TRIGGER_DAILY
          tmp = {}
          tmp['days_interval'] = trig.DaysInterval
          trigger['type'] = tmp
        when 3
          trigger['trigger_type'] = TASK_TIME_TRIGGER_WEEKLY
          tmp = {}
          tmp['weeks_interval'] = trig.WeeksInterval
          tmp['days_of_week'] =trig.DaysOfWeek
          trigger['type'] = tmp
        when 4
          trigger['trigger_type'] = TASK_TIME_TRIGGER_MONTHLYDATE
          tmp = {}
          tmp['months'] = trig.MonthsOfYear
          tmp['days'] = trig.DaysOfMonth
          trigger['type'] = tmp
        when 5
          trigger['trigger_type'] =TASK_TIME_TRIGGER_MONTHLYDOW
          tmp = {}
          tmp['months'] = trig.MonthsOfYear
          tmp['days_of_week'] = trig.DaysOfWeek
          tmp['weeks'] = trig.weeks
          trigger['type'] = tmp
        when 1
          trigger['trigger_type'] =TASK_TIME_TRIGGER_ONCE
          tmp = {}
          tmp['once'] = nil
          trigger['type'] = tmp
        else
          raise Error, 'Unknown trigger type'
      end
      trigger
    end

    # Sets the trigger for the currently active task.
    #
    def trigger=(trigger)
      raise Error, 'No currently active task' if @task.nil?
      raise TypeError unless trigger.is_a?(Hash)

      definition = @task.Definition
      definition.Triggers.Clear()

      case trigger['trigger_type']
        when TASK_TIME_TRIGGER_DAILY
          type = 2
        when TASK_TIME_TRIGGER_WEEKLY
          type = 3
        when TASK_TIME_TRIGGER_MONTHLYDATE
          type = 4
        when TASK_TIME_TRIGGER_MONTHLYDOW
          type = 5
        when TASK_TIME_TRIGGER_ONCE
          type = 1
        else
          raise Error, 'Unknown trigger type'
      end

      startTime = "%04d-%02d-%02dT%02d:%02d:00" % [trigger['start_year'],trigger['start_month'],trigger['start_day'],trigger['start_hour'], trigger['start_minute']]
      endTime = "%04d-%02d-%02dT00:00:00" % [trigger['end_year'],trigger['end_month'],trigger['end_day']]

      trig = definition.Triggers.Create(type)
      trig.Id = "RegistrationTriggerId"
      trig.StartBoundary = startTime
      trig.EndBoundary = endTime if endTime != '0000-00-00T00:00:00'
      trig.Enabled = true

      repetitionPattern = trig.Repetition
      repetitionPattern.Duration = "PT#{trigger['minutes_duration']||0}M" if trigger['minutes_duration'].to_i>0
      repetitionPattern.Interval  = "PT#{trigger['minutes_interval']||0}M" if trigger['minutes_interval'].to_i>0

      tmp = trigger['type']
      tmp = nil unless tmp.is_a?(Hash)

      case trigger['trigger_type']
        when TASK_TIME_TRIGGER_DAILY
          trig.DaysInterval =tmp['days_interval'] if tmp && tmp['days_interval']
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']}M" if trigger['random_minutes_interval'].to_i>0
        when TASK_TIME_TRIGGER_WEEKLY
          trig.DaysOfWeek  = tmp['days_of_week'] if tmp && tmp['days_of_week']
          trig.WeeksInterval  = tmp['weeks_interval'] if tmp && tmp['weeks_interval']
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']||0}M" if trigger['random_minutes_interval'].to_i>0
        when TASK_TIME_TRIGGER_MONTHLYDATE
          trig.MonthsOfYear  = tmp['months'] if tmp && tmp['months']
          trig.DaysOfMonth  = tmp['days'] if tmp && tmp['days']
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']||0}M" if trigger['random_minutes_interval'].to_i>0
        when TASK_TIME_TRIGGER_MONTHLYDOW
          trig.MonthsOfYear  = tmp['months'] if tmp && tmp['months']
          trig.DaysOfWeek  = tmp['days_of_week'] if tmp && tmp['days_of_week']
          trig.WeeksOfMonth  = tmp['weeks'] if tmp && tmp['weeks']
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']||0}M" if trigger['random_minutes_interval'].to_i>0
        when TASK_TIME_TRIGGER_ONCE
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']||0}M" if trigger['random_minutes_interval'].to_i>0
      end

      user = definition.Principal.UserId

      @root.RegisterTaskDefinition(
        @task.Path,
        definition,
        TASK_CREATE_OR_UPDATE,
        user,
        @password,
        @password ? TASK_LOGON_PASSWORD : TASK_LOGON_INTERACTIVE_TOKEN
      )

      trigger
    end

    # Adds a trigger at the specified index.
    #
    def add_trigger(index, trigger)
      raise Error, 'No currently active task' if @task.nil?
      raise TypeError unless trigger.is_a?(Hash)

      definition = @task.Definition
      case trigger['trigger_type']
        when TASK_TIME_TRIGGER_DAILY
          type = 2
        when TASK_TIME_TRIGGER_WEEKLY
          type = 3
        when TASK_TIME_TRIGGER_MONTHLYDATE
          type = 4
        when TASK_TIME_TRIGGER_MONTHLYDOW
          type = 5
        when TASK_TIME_TRIGGER_ONCE
          type = 1
        else
          raise Error, 'Unknown trigger type'
      end

      startTime = "%04d-%02d-%02dT%02d:%02d:00" % [trigger['start_year'],trigger['start_month'],trigger['start_day'],trigger['start_hour'], trigger['start_minute']]
      endTime = "%04d-%02d-%02dT00:00:00" % [trigger['end_year'],trigger['end_month'],trigger['end_day']]

      trig = definition.Triggers.Create(type)
      trig.Id = "RegistrationTriggerId"
      trig.StartBoundary = startTime
      trig.EndBoundary = endTime if endTime != '0000-00-00T00:00:00'
      trig.Enabled = true

      repetitionPattern = trig.Repetition
      repetitionPattern.Duration = "PT#{trigger['minutes_duration']||0}M" if trigger['minutes_duration'].to_i>0
      repetitionPattern.Interval  = "PT#{trigger['minutes_interval']||0}M" if trigger['minutes_interval'].to_i>0

      tmp = trigger['type']
      tmp = nil unless tmp.is_a?(Hash)

      case trigger['trigger_type']
        when TASK_TIME_TRIGGER_DAILY
          trig.DaysInterval =tmp['days_interval'] if tmp && tmp['days_interval']
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']}M"
        when TASK_TIME_TRIGGER_WEEKLY
          trig.DaysOfWeek  = tmp['days_of_week'] if tmp && tmp['days_of_week']
          trig.WeeksInterval  = tmp['weeks_interval'] if tmp && tmp['weeks_interval']
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']||0}M"
        when TASK_TIME_TRIGGER_MONTHLYDATE
          trig.MonthsOfYear  = tmp['months'] if tmp && tmp['months']
          trig.DaysOfMonth  = tmp['days'] if tmp && tmp['days']
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']||0}M"
        when TASK_TIME_TRIGGER_MONTHLYDOW
          trig.MonthsOfYear  = tmp['months'] if tmp && tmp['months']
          trig.DaysOfWeek  = tmp['days_of_week'] if tmp && tmp['days_of_week']
          trig.WeeksOfMonth  = tmp['weeks'] if tmp && tmp['weeks']
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']||0}M"
        when TASK_TIME_TRIGGER_ONCE
          trig.RandomDelay = "PT#{trigger['random_minutes_interval']||0}M"
      end

      user = definition.Principal.UserId

      begin
        @root.RegisterTaskDefinition(
          @task.Path,
          definition,
          TASK_CREATE_OR_UPDATE,
          user,
          @password,
          @password ? TASK_LOGON_PASSWORD : TASK_LOGON_INTERACTIVE_TOKEN
        )
      rescue
        raise Error, 'add_trigger failed'
      end

      true
    end

    # Returns the flags (integer) that modify the behavior of the work item. You
    # must OR the return value to determine the flags yourself.
    #--
    # Not supported
    #
    def flags
      raise Error, 'null task' if @task.nil?

      flags = 0
    end

    # Sets an OR'd value of flags that modify the behavior of the work item.
    #--
    # Not supported
    #
    def flags=(flags)
      raise Error, 'No currently active task' if @task.nil?

      flags = 0
    end

    # Returns the status of the currently active task. Possible values are
    # 'ready', 'running', 'not scheduled' or 'unknown'.
    #
    def status
      raise Error, 'No currently active task' if @task.nil?

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

    # Returns the exit code from the last scheduled run.
    #
    def exit_code
      raise Error, 'No currently active task' if @task.nil?

      exit_code = @task.LastTaskResult
    end

    # Returns the comment associated with the task, if any.
    #
    def comment
      raise Error, 'No currently active task' if @task.nil?

      @task.Definition.RegistrationInfo.Description
    end

    # Sets the comment for the task.
    #
    def comment=(comment)
      raise Error, 'No currently active task' if @task.nil?
      raise TypeError unless comment.is_a?(String)

      definition = @task.Definition
      definition.RegistrationInfo.Description = comment

      user = definition.Principal.UserId
      @root.RegisterTaskDefinition(@task.Path, definition,0x6,user,@password,@password ? 1 : 3)

      comment
    end

    # Returns the name of the user who created the task.
    #
    def creator
      raise Error, 'No currently active task' if @task.nil?

      @task.Definition.RegistrationInfo.Author
    end

    # Sets the creator for the task.
    #
    def creator=(creator)
      raise Error, 'No currently active task' if @task.nil?
      raise TypeError unless creator.is_a?(String)

      definition = @task.Definition
      definition.RegistrationInfo.Author = creator

      user = definition.Principal.UserId
      @root.RegisterTaskDefinition(@task.Path, definition,0x6,user,@password,@password ? 1 : 3)

      creator
    end

    # Returns a Time object that indicates the next time the task will run.
    #
    def next_run_time
      raise Error, 'No currently active task' if @task.nil?

      Time.parse(@task.NextRunTime)
    end

    # Returns a Time object indicating the most recent time the task ran or
    # nil if the task has never run.
    #
    def most_recent_run_time
      raise Error, 'No currently active task' if @task.nil?

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
      raise Error, 'No currently active task' if @task.nil?

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
      time += sec if sec
      time *= 1000

      time
    end

    # Sets the maximum length of time, in milliseconds, that the task can run
    # before terminating. Returns the value you specified if successful.
    #
    def max_run_time=(max_run_time)
      raise Error, 'No currently active task' if @task.nil?
      raise TypeError unless max_run_time.is_a?(Numeric)

      t = max_run_time
      t /= 1000
      limit ="PT#{t}S"

      definition = @task.Definition
      definition.Settings.ExecutionTimeLimit = limit
      user = definition.Principal.UserId

      @root.RegisterTaskDefinition(
        @task.Path,
        definition,
        TASK_CREATE_OR_UPDATE,
        user,
        @password,
        @password ? TASK_LOGON_PASSWORD : TASK_LOGON_INTERACTIVE_TOKEN
      )

      max_run_time
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
   end
end
