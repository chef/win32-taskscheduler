require "rake"
require "rake/clean"
require "rake/testtask"
require "rspec/core/rake_task"

CLEAN.include("**/*.gem", "**/*.rbc")

namespace "gem" do
  desc "Build the win32-taskscheduler gem"
  task create: [:clean] do
    require "rubygems/package"
    spec = eval(IO.read("win32-taskscheduler.gemspec"))
    Gem::Package.build(spec, true)
  end

  desc "Install the win32-taskscheduler library as a gem"
  task install: [:create] do
    file = Dir["win32-taskscheduler*.gem"].first
    sh "gem install -l #{file}"
  end
end

desc "Run the example code"
task :example do
  ruby "-Iib examples/taskscheduler_example.rb"
end

begin
  require "rspec/core/rake_task"

  RSpec::Core::RakeTask.new do |t|
    t.pattern = "spec/**/*_spec.rb"
  end
rescue LoadError
  desc "rspec is not installed, this task is disabled"
  task :spec do
    abort "rspec is not installed. bundle install first to make sure all dependencies are installed."
  end
end

desc "Check Linting and code style."
task :style do
  require "rubocop/rake_task"
  require "cookstyle/chefstyle"

  if RbConfig::CONFIG["host_os"] =~ /mswin|mingw|cygwin/
    # Windows-specific command, rubocop erroneously reports the CRLF in each file which is removed when your PR is uploaeded to GitHub.
    # This is a workaround to ignore the CRLF from the files before running cookstyle.
    sh "cookstyle --chefstyle -c .rubocop.yml --except Layout/EndOfLine"
  else
    sh "cookstyle --chefstyle -c .rubocop.yml"
  end
rescue LoadError
  puts "Rubocop or Cookstyle gems are not installed. bundle install first to make sure all dependencies are installed."
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

task default: %i{style spec}
