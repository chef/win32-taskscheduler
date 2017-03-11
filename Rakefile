require 'rake'
require 'rake/clean'
require 'rake/testtask'

CLEAN.include("**/*.gem", "**/*.rbc")

namespace 'gem' do
  desc 'Build the win32-taskscheduler gem'
  task :create => [:clean] do
    require 'rubygems/package'
    spec = eval(IO.read('win32-taskscheduler.gemspec'))
    Gem::Package.build(spec, true)
  end

  desc 'Install the win32-taskscheduler library as a gem'
  task :install => [:create] do
    file = Dir['win32-taskscheduler*.gem'].first
    sh "gem install -l #{file}"
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

task :default => :test
