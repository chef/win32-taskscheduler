require 'rake'
require 'rake/testtask'
require 'rbconfig'
include Config

namespace 'gem' do
  desc 'Remove any existing .gem files'
  task :clean do
    Dir['*.gem'].each{ |f| File.delete(f) }
  end

  desc 'Build the win32-taskscheduler gem'
  task :build => [:clean] do
    spec = eval(IO.read('win32-taskscheduler.gemspec'))
    Gem::Builder.new(spec).build
  end

  desc 'Install the win32-taskscheduler library as a gem'
  task :install => [:build] do
    file = Dir['win32-taskscheduler*.gem'].first
    sh "gem install #{file}"
  end
end

desc 'Run the example code'
task :example do
  ruby '-Iib examples/taskscheduler_example.rb'
end

desc 'Run the test suite for the win32-taskscheduler library'
Rake::TestTask.new do |t|
  t.verbose = true
  t.warning = true
end
