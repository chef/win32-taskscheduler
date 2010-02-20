require 'rubygems'

Gem::Specification.new do |spec|
  spec.name       = 'win32-taskscheduler'
  spec.version    = '0.2.0'
  spec.authors    = ['Park Heesob', 'Daniel J. Berger']
  spec.license    = 'Artistic 2.0'
  spec.email      = 'djberg96@gmail.com'
  spec.homepage   = 'http://www.rubyforge.org/projects/win32utils'
  spec.platform   = Gem::Platform::RUBY
  spec.summary    = 'A library for the Windows task scheduler'
  spec.has_rdoc   = true
  spec.test_files = Dir['test/test*']
  spec.files      = Dir['**/*'].reject{ |f| f.include?('git') }

  spec.rubyforge_project = 'win32utils'

  spec.extra_rdoc_files = [
    'README',
    'CHANGES',
    'MANIFEST',
    'doc/taskscheduler.txt'
  ]

  spec.description = <<-EOF
    The win32-taskscheduler library provides an interface to the MS Windows
    Task Scheduler. With this interface you can create new scheduled tasks,
    configure existing tasks, or delete tasks.
  EOF
end
