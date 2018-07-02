require 'rubygems'
require_relative 'lib/win32/taskscheduler/version'

Gem::Specification.new do |spec|
  spec.name       = 'win32-taskscheduler'
  spec.version    = Win32::TaskScheduler::VERSION
  spec.authors    = ['Park Heesob', 'Daniel J. Berger']
  spec.license    = 'Artistic-2.0'
  spec.email      = 'djberg96@gmail.com'
  spec.homepage   = 'http://github.com/chef/win32-taskscheduler'
  spec.summary    = 'A library for the Windows task scheduler'
  spec.test_files = Dir['test/test*']
  spec.files      = Dir['**/*'].reject{ |f| f.include?('git') }

  spec.add_dependency('ffi')
  spec.add_dependency('structured_warnings')

  spec.add_development_dependency('test-unit')
  spec.add_development_dependency('rake')
  spec.add_development_dependency('win32-security')
  spec.add_development_dependency "rspec", "~>3.0"

  spec.extra_rdoc_files = [
    "README.md",
    "CHANGELOG.md",
    "RELEASE_NOTES.md",
  ]

  spec.description = <<-EOF
    The win32-taskscheduler library provides an interface to the MS Windows
    Task Scheduler. With this interface you can create new scheduled tasks,
    configure existing tasks, or delete tasks.
  EOF
end
