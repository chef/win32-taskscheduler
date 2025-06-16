libx = File.expand_path("lib", __dir__)
$LOAD_PATH.unshift(libx) unless $LOAD_PATH.include?(libx)

require "date"

# In order to test some dummy task and folder, we are creating
# 'Test' folder and consider it as 'root'.
# After the test case, all the above created elements
# i.e, 'Test', Tasks and nested folders will be deleted.
# It will ensure No Harm be caused to any existing tasks at root

# Creating folder 'Test'; This will be treated as root
def create_test_folder
  @service ||= service
  @root_path = '\\'
  @test_path = '\\Test'
  @root_folder = @service.GetFolder(@root_path)
  @test_folder = @root_folder.CreateFolder(@test_path)
  @ts.instance_variable_set(:@root, @test_folder) if @ts
end

# Deleting the test folder its nested folder(if any)
# And the tasks within(if any)
def clear_them
  if @test_folder
    delete_all(@test_folder)
  else
    @test_folder = nil
  end
end

# Logic to delete tasks and folder using Win32OLE methods
def delete_all(folder)
  delete_tasks_in(folder)
  folder.GetFolders(0).each do |nested_folder|
    delete_all(nested_folder)
  end
  @root_folder.DeleteFolder(folder.Path, 0)
end

# Deleting the task in specified folder using Win32OLE methods
def delete_tasks_in(folder)
  folder.GetTasks(0).each do |task|
    folder.DeleteTask(task.Name, 0)
  end
end

# Returns Win32 Service object
def service
  service = WIN32OLE.new("Schedule.Service")
  service.Connect
  service
end

# Returns no of task in the specified folder
def no_of_tasks(folder = @test_folder)
  if folder.is_a?(String)
    folder = @test_path + folder
    folder = @service.GetFolder(folder)
  end
  folder.GetTasks(0).Count
end

# Steps to create (Register) a task
# This is an example activity to test
# the related functionalities over a task in RSpecs
def create_task
  return nil unless @service

  @task_definition = @service.NewTask(0)
  task_registration
  task_prinicipals
  task_settings
  task_triggers
  task_action
  register_task
end

def task_registration
  reg_info = @task_definition.RegistrationInfo
  reg_info.Description = "Sample task for testing purpose"
  reg_info.Author = "Rspec"
end

def task_prinicipals
  principal = @task_definition.Principal
  principal.LogonType = 3 # Interactive Logon
end

def task_settings
  settings = @task_definition.Settings
  settings.Enabled = true
  settings.StartWhenAvailable = true
  settings.Hidden = false
end

def task_triggers
  triggers = @task_definition.Triggers
  trigger = triggers.Create(1) # Time
  start_time, end_time = start_end_time
  trigger.StartBoundary = start_time
  trigger.EndBoundary = end_time
  trigger.ExecutionTimeLimit = "PT5M" # Five minutes
  trigger.Id = "TimeTriggerId"
  trigger.Enabled = true
end

def task_action
  action = @task_definition.Actions.Create(0)
  action.Path = @app
end

# Registering(Creating) a task in test folder
def register_task(user = 'SYSTEM', password = nil)
  @current_task = @test_folder.RegisterTaskDefinition(
    @task,
    @task_definition,
    6,
    user,
    password,
    3
  )
  @ts.instance_variable_set(:@task, @current_task) if @ts
end

def start_end_time
  t = Date.new(2010)
  start_time = t.strftime("%FT%T")
  end_time = (t + 5).strftime("%FT%T") # 5 Days
  [start_time, end_time]
end

def tasksch_err
  Win32::TaskScheduler::Error
end

# Methods to build all types of triggers and their values

def all_triggers
  all_triggers = {}

  %w{
    ONCE
    DAILY
    WEEKLY
    MONTHLYDATE
    MONTHLYDOW
    ON_IDLE
    AT_SYSTEMSTART
    AT_LOGON
  }.each do |trig_type|
    trigger = {}
    trigger[:trigger_type] = Win32::TaskScheduler.class_eval(trig_type)
    start_end_params(trigger)
    other_trigger_params(trig_type, trigger)
    all_triggers[trig_type] = trigger
  end

  all_triggers
end

def start_end_params(trigger)
  %i{start_year end_year}.each do |t|
    trigger[t] = "2030"
  end

  %i{start_month start_day start_hour start_minute}.each do |t|
    trigger[t] = "02"
  end

  %i{end_day end_month}.each do |t|
    trigger[t] = "03"
  end

  %i{minutes_duration minutes_interval}.each do |t|
    trigger[t] = 2
  end
end

def other_trigger_params(trig_type, trigger)
  type = {}
  case trig_type
  when "ONCE"
    trigger[:type] = type
    type[:once] = nil
    trigger[:random_minutes_interval] = 2
  when "DAILY"
    trigger[:type] = type
    type[:days_interval] = 2
    trigger[:random_minutes_interval] = 2
  when "WEEKLY"
    trigger[:type] = type
    type[:weeks_interval] = 2
    type[:days_of_week] = sunday
    trigger[:random_minutes_interval] = 2
  when "MONTHLYDATE"
    trigger[:type] = type
    type[:months] = january
    type[:days] = first_day
    trigger[:run_on_last_day_of_month] = false
    trigger[:random_minutes_interval] = 2
  when "MONTHLYDOW"
    trigger[:type] = type
    type[:months] = january
    type[:days_of_week] = sunday
    type[:weeks_of_month] = first_week
    trigger[:run_on_last_week_of_month] = false
    trigger[:random_minutes_interval] = 2
  when "ON_IDLE"
    trigger[:execution_time_limit] = 2
  when "AT_SYSTEMSTART"
    trigger[:delay_duration] = 2
  when "AT_LOGON"
    trigger[:user_id] = "SYSTEM"
    trigger[:delay_duration] = 2
  end
end

def sunday
  Win32::TaskScheduler::SUNDAY
end

def january
  Win32::TaskScheduler::JANUARY
end

def first_day
  Win32::TaskScheduler::FIRST
end

def first_week
  Win32::TaskScheduler::FIRST_WEEK
end
