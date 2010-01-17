require 'rake'
require 'rake/clean'
require 'rake/testtask'
require 'rbconfig'
include Config

=begin
desc "Cleans up the C related files created during the build"
task :clean do
   Dir.chdir('ext') do
      if File.exists?('taskscheduler.o') || File.exists?('taskscheduler.so')
         sh 'nmake distclean'
      end

      if File.exists?('win32/taskscheduler.so')
         File.delete('win32/taskscheduler.so')
      end         
   end
end

desc "Builds, but does not install, the win32-taskscheduler library"
task :build => [:clean] do
   Dir.chdir('ext') do
      ruby 'extconf.rb'
      sh 'nmake'
      FileUtils.cp('taskscheduler.so', 'win32/taskscheduler.so')      
   end  
end
=end

desc "Install the win32-taskscheduler library (non-gem)"
task :install => [:build] do
   dir = File.join(Config::CONFIG['sitelibdir'], 'win32')
   Dir.mkdir(dir) unless File.exists?(dir)
   FileUtils.cp('lib/win32/taskscheduler.rb', dir)
end

desc 'Install the win32-taskscheduler library as a gem'
task :install_gem do
   ruby 'win32-taskscheduler.gemspec'
   file = Dir['win32-taskscheduler*.gem'].first
   sh "gem install #{file}"
end

desc 'Run the example code'
task :example do
   ruby '-Iib examples/taskscheduler_example.rb'
end

desc "Run the test suite for the win32-taskscheduler library"
Rake::TestTask.new do |t|
   t.verbose = true
   t.warning = true
end
