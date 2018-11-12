source "https://rubygems.org"

gemspec

group :development do
  gem "chefstyle"
  gem "rake"
  gem "rspec", "~> 3.0"
end

group :ci do
  gem "rspec_junit_formatter"
end

group :docs do
  gem "github-markup"
  gem "redcarpet"
  gem "yard"
end

group :debug do
  gem "pry"
  gem "pry-byebug"
  gem "pry-stack_explorer"
end

instance_eval(ENV["GEMFILE_MOD"]) if ENV["GEMFILE_MOD"]

# If you want to load debugging tools into the bundle exec sandbox,
# add these additional dependencies into Gemfile.local
eval_gemfile(__FILE__ + ".local") if File.exist?(__FILE__ + ".local")
