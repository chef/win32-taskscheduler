require 'rake'
require 'rake/clean'
require 'rake/testtask'
require 'rspec/core/rake_task'

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

RSpec::Core::RakeTask.new(:spec) do |spec|
  spec.pattern = FileList["spec/**/*_spec.rb", "spec/**/**/*_spec.rb"].to_a
end

desc 'Run the test suite for the win32-taskscheduler library'
Rake::TestTask.new do |t|
  t.verbose = true
  t.warning = true
end

begin
  require "yard"
  YARD::Rake::YardocTask.new(:docs)
rescue LoadError
  puts "yard is not available. bundle install first to make sure all dependencies are installed."
end

task :console do
  require "irb"
  require "irb/completion"
  ARGV.clear
  IRB.start
end

task :default => :test
