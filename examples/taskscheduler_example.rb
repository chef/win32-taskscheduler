#######################################################################
# taskscheduler_example.rb
#
# A test script for general futzing. You can run this example via the
# 'example' rake task.
#
# Modify as you see fit.
#######################################################################
require "win32/taskscheduler"
require "fileutils"
require "pp"
include Win32

puts "VERSION: " + TaskScheduler::VERSION

ts = TaskScheduler.new

trigger = {
  start_year: 2009,
  start_month: 4,
  start_day: 11,
  start_hour: 7,
  start_minute: 14,
  trigger_type: TaskScheduler::DAILY,
  type: { "days_interval" => 1 },
}

if ts.enum.grep(/foo/).empty?
  ts.new_work_item("foo", trigger)
  ts.application_name = "notepad.exe"
  puts "Task Added"
end

ts.activate("foo")
ts.priority = TaskScheduler::IDLE
ts.working_directory = 'C:\\'

puts "App name: " + ts.application_name
puts "Creator: " + ts.creator
puts "Exit code: " + ts.exit_code.to_s
puts "Max run time: " + ts.max_run_time.to_s
puts "Next run time: " + ts.next_run_time.to_s
puts "Parameters: " + ts.parameters
puts "Priority: " + ts.priority.to_s
puts "Status: " + ts.status
puts "Trigger count: " + ts.trigger_count.to_s
puts "Trigger string: " + ts.trigger_string(0)
puts "Working directory: " + ts.working_directory
puts "Trigger: "

pp ts.trigger(0)

ts.delete("foo")
puts "Task deleted"
