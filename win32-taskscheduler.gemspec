require 'rubygems'

spec = Gem::Specification.new do |gem|
   gem.name              = 'win32-taskscheduler'
   gem.version           = '0.2.0'
   gem.authors           = ['Park Heesob', 'Daniel J. Berger']
   gem.email             = 'djberg96@gmail.com'
   gem.homepage          = 'http://www.rubyforge.org/projects/win32utils'
   gem.rubyforge_project = 'win32utils'
   gem.platform          = Gem::Platform::RUBY
   gem.summary           = 'A library for the Windows task scheduler'
   gem.has_rdoc          = true
   gem.test_files        = Dir['test/test*']
   gem.files             = Dir['**/*'].delete_if{ |f| f.include?('CVS') || f.include?('ext') }
   gem.license           = 'Artistic 2.0'

   gem.extra_rdoc_files = [
      'README',
      'CHANGES',
      'MANIFEST',
      'doc/taskscheduler.txt'
   ]

   gem.required_ruby_version = '>= 1.8.0'

   gem.description = <<-EOF
      The win32-taskscheduler library provides an interface to the MS Windows
      Task Scheduler. With this interface you can create new scheduled tasks,
      configure existing tasks, or delete tasks.
   EOF
end

if $0 == __FILE__
   Gem.manage_gems if Gem::RubyGemsVersion.to_f < 1.0
   Gem::Builder.new(spec).build
end
